"""Pure calculation and import code; independent of the GUI and user files."""
from __future__ import annotations

import csv
import io
import json
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

from i18n import LocalizedError, Message, translate

CATEGORIES = {'electricity': '电费', 'water': '水费', 'gas': '燃气费'}
FIELDS = ('ReadDatetime', 'ReadValue', 'MeterSerial', 'MeterCode', 'QualityMethod')
REQUIRED = FIELDS[:3]
EXTENSIONS = {'.csv', '.xlsx', '.xls', '.xlsm', '.xlsb'}
UNITS = {'kWh': ('energy', Decimal(1)), 'MJ': ('energy', Decimal(1) / Decimal('3.6')),
         'L': ('volume', Decimal(1)), 'kL': ('volume', Decimal(1000)), 'm³': ('volume', Decimal(1000))}


def number(value, label):
    try:
        result = Decimal(str(value).strip())
    except (InvalidOperation, ValueError):
        raise LocalizedError('{label}必须为数字。', label=Message(label)) from None
    if not result.is_finite() or result < 0:
        raise LocalizedError('{label}必须为非负有限数值。', label=Message(label))
    return result


def money(value):
    return str(value.quantize(Decimal('.01'), rounding=ROUND_HALF_UP))


@dataclass(frozen=True)
class Tariff:
    usage_rate: Decimal
    daily_rate: Decimal
    data_unit: str
    rate_unit: str
    interval_minutes: int = 0

    @classmethod
    def parse(cls, data):
        usage = number(data['usage_rate'], '用量单价')
        daily = number(data['daily_rate'], '每日固定费')
        source, target = data['data_unit'], data['rate_unit']
        if source not in UNITS or target not in UNITS:
            raise LocalizedError('不支持的单位。')
        if UNITS[source][0] != UNITS[target][0]:
            raise LocalizedError('用量单位和计价单位不能跨体积 / 能量转换；燃气体积转能量需要供应商换算系数。')
        interval = number(data.get('interval_minutes', 0), '采样间隔')
        if interval != int(interval) or (interval and (interval > 1440 or 1440 % int(interval))):
            raise LocalizedError('采样间隔应为0，或能整除1440的正整数分钟数。')
        return cls(usage, daily, source, target, int(interval))

    @property
    def factor(self):
        if self.data_unit == self.rate_unit:
            return Decimal(1)
        if self.data_unit == 'kWh' and self.rate_unit == 'MJ':
            return Decimal('3.6')
        return UNITS[self.data_unit][1] / UNITS[self.rate_unit][1]

    def as_dict(self):
        return {'usage_rate': str(self.usage_rate), 'daily_rate': str(self.daily_rate),
                'data_unit': self.data_unit, 'rate_unit': self.rate_unit,
                'interval_minutes': self.interval_minutes}


def load_tariffs(path):
    data = json.loads(Path(path).read_text(encoding='utf-8-sig'))
    if not isinstance(data, dict) or data.get('version') != 1:
        raise LocalizedError('费率文件 version 必须为1。')
    currency = data.get('currency', '')
    if not isinstance(currency, str) or len(currency) != 3 or not currency.isascii() or not currency.isalpha():
        raise LocalizedError('currency 应为三个英文字母，如 AUD、CNY。')
    categories = data.get('categories')
    if not isinstance(categories, dict) or set(categories) != set(CATEGORIES):
        raise LocalizedError('categories 必须包含且仅包含 electricity、water、gas。')
    return currency.upper(), {k: Tariff.parse(v) for k, v in categories.items()}


@dataclass(frozen=True)
class Source:
    path: Path
    sheet: str | None = None

    @property
    def label(self):
        return self.path.name + (f' / {self.sheet}' if self.sheet is not None else '')


def workbook(path):
    try:
        from python_calamine import CalamineWorkbook
    except ImportError:
        raise LocalizedError('Excel 导入需要依赖，请运行 python -m pip install -r requirements.txt') from None
    return CalamineWorkbook.from_path(str(path))


def sources(path):
    path = Path(path)
    if path.suffix.lower() == '.csv':
        return [Source(path)]
    if path.suffix.lower() not in EXTENSIONS:
        raise LocalizedError('不支持的文件扩展名。')
    with workbook(path) as book:
        return [Source(path, name) for name in book.sheet_names]


def read_table(source):
    if source.path.suffix.lower() == '.csv':
        raw = source.path.read_bytes()
        for encoding in ('utf-8-sig', 'gb18030'):
            try:
                text = raw.decode(encoding)
                break
            except UnicodeDecodeError:
                pass
        else:
            raise LocalizedError('CSV 编码应为 UTF-8 或 GB18030。')
        return list(csv.reader(io.StringIO(text)))
    if source.sheet is None:
        raise LocalizedError('请先选择 Excel 工作表。')
    with workbook(source.path) as book:
        return book.get_sheet_by_name(source.sheet).to_python(skip_empty_area=False)


def timestamp(value):
    if isinstance(value, datetime):
        result = value
    elif isinstance(value, date):
        result = datetime.combine(value, datetime.min.time())
    elif isinstance(value, str):
        try:
            result = datetime.fromisoformat(value.strip())
        except ValueError:
            raise LocalizedError('日期应为 YYYY-MM-DD 或 YYYY-MM-DD HH:MM:SS。') from None
    else:
        raise LocalizedError('日期应为 ISO 文本或 Excel 日期单元格，不能是无日期格式的序列号。')
    if result.tzinfo is not None:
        raise LocalizedError('请先将带时区日期转换为账单当地时间，并移除时区标记。')
    return result


def read_records(source, mapping):
    table = read_table(source)
    if not table:
        raise LocalizedError('{source}：空表。', source=source.label)
    header = [str(v).strip() if v is not None else '' for v in table[0]]
    active = {k: str(mapping.get(k, '')).strip() for k in FIELDS}
    if any(not active[k] for k in REQUIRED):
        raise LocalizedError('日期、用量、电表号的列名映射不能为空。')
    names = [v for v in active.values() if v]
    if len(names) != len(set(names)):
        raise LocalizedError('不同字段不能映射到同一列。')
    for field, name in active.items():
        if name and header.count(name) > 1:
            raise LocalizedError('{source}：列名 {name} 重复。', source=source.label, name=name)
        if field in REQUIRED and name not in header:
            raise LocalizedError('{source}：缺少列 {name}；请检查第一行和列名映射。', source=source.label, name=name)
    indices = {k: header.index(v) for k, v in active.items() if v and v in header}
    count = 0
    for line, row in enumerate(table[1:], 2):
        if not any(v is not None and str(v).strip() for v in row):
            continue
        if any(v is not None and str(v).strip() for v in row[len(header):]):
            raise LocalizedError('{source} 第{line}行：数据列数超过表头。', source=source.label, line=line)
        def cell(key):
            index = indices.get(key)
            return row[index] if index is not None and index < len(row) else None
        def text(key):
            value = cell(key)
            return '' if value is None else str(value).strip()
        try:
            stamp = timestamp(cell('ReadDatetime'))
            usage = number(cell('ReadValue'), '用量')
            meter = text('MeterSerial')
            if not meter:
                raise LocalizedError('表号不能为空。')
            yield stamp, usage, meter, text('MeterCode'), text('QualityMethod')
            count += 1
        except ValueError as exc:
            raise LocalizedError('{source} 第{line}行：{error}', source=source.label, line=line, error=exc.message if isinstance(exc, LocalizedError) else str(exc)) from exc
    if not count:
        raise LocalizedError('{source}：没有用量记录。', source=source.label)


@dataclass(frozen=True)
class Daily:
    category: str
    day: date
    count: int
    usage: Decimal
    unit: str
    usage_cost: Decimal
    daily_cost: Decimal
    status: str
    status_en: str = ''

    @property
    def total(self):
        return self.usage_cost + self.daily_cost


def calculate(category, selected, tariff, mapping, start=None, end=None):
    if start and end and start > end:
        raise LocalizedError('开始日期不能晚于结束日期。')
    records, duplicates = {}, 0
    for source in selected:
        for stamp, usage, meter, code, quality in read_records(source, mapping):
            if (start and stamp.date() < start) or (end and stamp.date() > end):
                continue
            key = (meter, code, stamp)
            if key in records:
                old_usage, qualities = records[key]
                if old_usage != usage:
                    raise LocalizedError('{stamp}：相同表号/通道存在冲突用量。', stamp=stamp)
                qualities.add(quality)
                duplicates += 1
            else:
                records[key] = (usage, {quality})
    if len({key[:2] for key in records}) > 1:
        raise LocalizedError('同一分类所选期间含多个表号或通道，请分开计算。')
    groups = defaultdict(list)
    for (_, _, stamp), (usage, quality) in records.items():
        groups[stamp.date()].append((stamp, usage, quality))
    result = []
    for day, values in sorted(groups.items()):
        usage = sum((v[1] for v in values), Decimal(0)) * tariff.factor
        warnings = []
        interval = tariff.interval_minutes
        if interval:
            slots = {v[0].hour * 60 + v[0].minute for v in values if not v[0].second and not v[0].microsecond}
            if len(values) != 1440 // interval or slots != set(range(0, 1440, interval)):
                warnings.append(Message('覆盖不完整或时间不对齐（{count}条），仍计1天固定费', count=len(values)))
        else:
            warnings.append('未检查全天覆盖')
        if any(q != 'A' for v in values for q in v[2]):
            warnings.append('非A或缺失质量标记')
        result.append(Daily(category, day, len(values), usage, tariff.rate_unit,
                            usage * tariff.usage_rate, tariff.daily_rate, '；'.join(warnings) or '完整', '; '.join(translate(w, 'en') for w in warnings) or 'Complete'))
    return result, duplicates
