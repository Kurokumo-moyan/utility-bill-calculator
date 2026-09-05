"""Small, explicit Chinese/English catalog; never translates imported user values."""

EN = {
    'Utility Bill Calculator · 水电燃气计算器': 'Utility Bill Calculator',
    '水电燃气计算器': 'Utility Bill Calculator',
    '选择文件或工作表 → 设置费率和列名 → 计算。输入必须是时段用量，不是累计表数。': 'Select files or sheets → Set rates and columns → Calculate. Use interval consumption, not cumulative meter readings.',
    '币种': 'Currency', '导入费率 JSON': 'Import rates JSON', '保存设置': 'Save settings', '刷新文件夹': 'Refresh folders',
    '电费': 'Electricity', '水费': 'Water', '燃气费': 'Gas', '选择文件夹': 'Choose folder', '添加文件': 'Add files',
    '用量单价 / 计价单位': 'Rate / billing unit', '固定费 / 天': 'Daily charge', '数据单位': 'Data unit', '计价单位': 'Billing unit',
    '间隔分钟（0不检查）': 'Interval min (0=off)',
    '列名映射：填写文件第一行实际列名（区分大小写）；后两项可留空': 'Column mapping: enter exact first-row headers (case-sensitive); last two fields are optional',
    '日期时间 *': 'Date/time *', '时段用量 *': 'Interval usage *', '表号 *': 'Meter ID *', '通道（可选）': 'Channel (optional)', '质量（可选）': 'Quality (optional)',
    '全选': 'Select all', '取消选择': 'Clear selection', '查看选中表的列名': 'Show selected headers',
    'Excel 各工作表分别列出，可 Ctrl / Shift 多选。': 'Excel sheets are listed separately. Ctrl / Shift selects multiple items.',
    '开始日期': 'Start date', '结束日期': 'End date', 'YYYY-MM-DD，留空不限': 'YYYY-MM-DD; blank = no limit',
    '计算选中数据': 'Calculate', '导出结果 CSV': 'Export CSV',
    '分类': 'Category', '日期': 'Date', '记录数': 'Records', '用量': 'Usage', '单位': 'Unit', '用量费': 'Usage cost', '固定费': 'Daily cost', '合计': 'Total', '数据检查': 'Data checks',
    '先输入费率或导入费率文件。': 'Enter rates or import a rates file first.',
    '固定费按各分类实际有记录的日期收取；缺失日期不补计。费率需自行包含适用税费。\n仅按天计费请将用量单价设为0。不同费率期间请分开计算；没有分摊功能。': 'Daily charges apply only to dates with records in each category; missing dates are not billed. Enter tax-inclusive rates.\nFor daily-only billing, set the usage rate to 0. Calculate different tariff periods separately. Bill splitting is not included.',
    '选择或设置已更改，请重新计算。': 'Selection or settings changed. Please calculate again.',
    '工作表': 'Sheet', '部分文件未读取': 'Some files could not be read',
    '请先选中一个文件或工作表。': 'Select a file or sheet first.', '第一行列名': 'First-row headers', '空表': 'Empty sheet', '读取失败': 'Read failed',
    '费率 JSON': 'Rates JSON', '费率已导入；如需下次保留，请点击保存设置。': 'Rates imported. Click Save settings to keep them for next time.',
    '费率导入失败': 'Rates import failed', '设置': 'Settings', '设置已保存到 settings.local.json。': 'Settings saved to settings.local.json.',
    '保存失败': 'Save failed', '设置格式不正确': 'Invalid settings format', '分类设置格式不正确': 'Invalid category settings format', '设置读取失败': 'Could not restore settings',
    '币种应为三个英文字母，如 AUD。': 'Currency must contain three English letters, e.g. AUD.',
    '开始日期不能晚于结束日期。': 'Start date must not be after end date.',
    '同一文件 / 工作表被分到多个类别，请取消重复选择。': 'The same file/sheet is selected in multiple categories. Remove the duplicate selection.',
    '没有可计算记录，请检查文件选择与日期。': 'No records to calculate. Check the selected files and dates.',
    '无法计算': 'Calculation failed', '导出': 'Export', '请先计算。': 'Calculate first.',
    '总额按未舍入金额求和后四舍五入至分': 'Total rounded to cents after summing unrounded amounts', '导出失败': 'Export failed',
    '{category} {days}天，缺{missing}天，去重{duplicates}条': '{category}: {days} days, {missing} missing days, {duplicates} duplicates removed',
    '{category}：所选范围无记录': '{category}: no records in selected range',
    '总计 {total} {currency}': 'Total {total} {currency}', '缺失日期未计固定费。': 'Daily charges for missing dates are excluded.',
    '{label}必须为数字。': '{label} must be a number.', '{label}必须为非负有限数值。': '{label} must be a finite, non-negative number.',
    '用量单价': 'Usage rate', '每日固定费': 'Daily charge', '采样间隔': 'Sampling interval',
    '不支持的单位。': 'Unsupported unit.',
    '用量单位和计价单位不能跨体积 / 能量转换；燃气体积转能量需要供应商换算系数。': 'Cannot convert between volume and energy; gas volume-to-energy conversion needs a supplier-specific factor.',
    '采样间隔应为0，或能整除1440的正整数分钟数。': 'Interval must be 0 or a positive whole number of minutes that divides 1440.',
    '费率文件 version 必须为1。': 'Rates file version must be 1.', 'currency 应为三个英文字母，如 AUD、CNY。': 'currency must contain three English letters, e.g. AUD or CNY.',
    'categories 必须包含且仅包含 electricity、water、gas。': 'categories must contain exactly electricity, water and gas.',
    'Excel 导入需要依赖，请运行 python -m pip install -r requirements.txt': 'Excel import requires dependencies. Run python -m pip install -r requirements.txt',
    '不支持的文件扩展名。': 'Unsupported file extension.', 'CSV 编码应为 UTF-8 或 GB18030。': 'CSV encoding must be UTF-8 or GB18030.',
    '请先选择 Excel 工作表。': 'Select an Excel sheet first.', '日期应为 YYYY-MM-DD 或 YYYY-MM-DD HH:MM:SS。': 'Use YYYY-MM-DD or YYYY-MM-DD HH:MM:SS for dates.',
    '日期应为 ISO 文本或 Excel 日期单元格，不能是无日期格式的序列号。': 'Use ISO date text or an Excel date cell, not an unformatted date serial number.',
    '请先将带时区日期转换为账单当地时间，并移除时区标记。': 'Convert timezone-aware dates to the billing location time and remove timezone markers first.',
    '{source}：空表。': '{source}: empty table.', '日期、用量、电表号的列名映射不能为空。': 'Date, usage and meter ID mappings are required.',
    '不同字段不能映射到同一列。': 'Different fields cannot map to the same column.',
    '{source}：列名 {name} 重复。': '{source}: duplicate header {name}.',
    '{source}：缺少列 {name}；请检查第一行和列名映射。': '{source}: missing column {name}; check the first row and column mapping.',
    '{source} 第{line}行：数据列数超过表头。': '{source}, row {line}: more data columns than headers.',
    '表号不能为空。': 'Meter ID cannot be empty.', '{source} 第{line}行：{error}': '{source}, row {line}: {error}',
    '{source}：没有用量记录。': '{source}: no usage records.', '{stamp}：相同表号/通道存在冲突用量。': '{stamp}: conflicting usage for the same meter/channel.',
    '同一分类所选期间含多个表号或通道，请分开计算。': 'Multiple meters or channels in one category. Calculate them separately.',
    '覆盖不完整或时间不对齐（{count}条），仍计1天固定费': 'Incomplete coverage or misaligned timestamps ({count} records); one daily charge still applies',
    '未检查全天覆盖': 'Full-day coverage not checked', '非A或缺失质量标记': 'Non-A or missing quality flag', '完整': 'Complete',
}


class Message(str):
    """Keep a Chinese-compatible string while retaining translation parameters."""
    def __new__(cls, key, **params):
        obj = super().__new__(cls, key.format(**params))
        obj.key, obj.params = key, params
        return obj

    def render(self, language):
        template = EN.get(self.key, self.key) if language == 'en' else self.key
        return template.format(**{k: translate(v, language) if isinstance(v, Message) else v for k, v in self.params.items()})


def translate(text, language='zh'):
    if isinstance(text, Message):
        return text.render(language)
    return EN.get(text, text) if language == 'en' else text


class LocalizedError(ValueError):
    def __init__(self, key, **params):
        self.message = Message(key, **params)
        super().__init__(self.message)
