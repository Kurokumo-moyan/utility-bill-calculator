"""Local desktop UI for CSV and Excel utility bills."""
import csv
import json
import tkinter as tk
from datetime import date
from decimal import Decimal
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from i18n import Message, LocalizedError, translate

from billing import (CATEGORIES, EXTENSIONS, FIELDS, UNITS, Tariff, calculate,
                     load_tariffs, money, read_table, sources)

BASE = Path(__file__).resolve().parent
SETTINGS = BASE / 'settings.local.json'


class App:
    def __init__(self, root):
        self.root, self.panels, self.results = root, {}, []
        self.language = 'zh'
        self.notes = []
        self.summary_message = '先输入费率或导入费率文件。'
        root.title('Utility Bill Calculator · 水电燃气计算器')
        root.geometry('1320x900')
        body = ttk.Frame(root, padding=12)
        body.pack(fill='both', expand=True)
        ttk.Label(body, text='水电燃气计算器', font=('TkDefaultFont', 17, 'bold')).pack(anchor='w')
        ttk.Label(body, text='选择文件或工作表 → 设置费率和列名 → 计算。输入必须是时段用量，不是累计表数。').pack(anchor='w', pady=5)
        bar = ttk.Frame(body); bar.pack(fill='x')
        self.currency = tk.StringVar(value='AUD')
        ttk.Label(bar, text='币种').pack(side='left')
        ttk.Entry(bar, textvariable=self.currency, width=7).pack(side='left', padx=5)
        for title, action in [('导入费率 JSON', self.import_rates), ('保存设置', self.save), ('刷新文件夹', self.refresh)]:
            ttk.Button(bar, text=title, command=action).pack(side='left', padx=4)
        self.language_button = ttk.Button(bar, text='English', command=self.toggle_language)
        self.language_button.pack(side='right')
        tabs = self.tabs = ttk.Notebook(body); tabs.pack(fill='x', pady=8)
        for category, label in CATEGORIES.items():
            f = ttk.Frame(tabs, padding=8); tabs.add(f, text=label)
            unit = 'kWh' if category == 'electricity' else 'L' if category == 'water' else 'MJ'
            panel = {'sources': []}
            for key, default in {'folder': str(BASE / 'data' / category), 'usage_rate': '', 'daily_rate': '',
                                 'data_unit': unit, 'rate_unit': unit, 'interval_minutes': '5' if category == 'electricity' else '0'}.items():
                panel[key] = tk.StringVar(value=default)
            row = ttk.Frame(f); row.pack(fill='x')
            ttk.Entry(row, textvariable=panel['folder'], width=60).pack(side='left', fill='x', expand=True)
            ttk.Button(row, text='选择文件夹', command=lambda c=category: self.choose_folder(c)).pack(side='left', padx=4)
            ttk.Button(row, text='添加文件', command=lambda c=category: self.add_files(c)).pack(side='left')
            row = ttk.Frame(f); row.pack(fill='x', pady=8)
            for key, label in [('usage_rate', '用量单价 / 计价单位'), ('daily_rate', '固定费 / 天'), ('data_unit', '数据单位'), ('rate_unit', '计价单位'), ('interval_minutes', '间隔分钟（0不检查）')]:
                ttk.Label(row, text=label).pack(side='left')
                widget = ttk.Combobox(row, textvariable=panel[key], values=list(UNITS), state='readonly', width=6) if key.endswith('unit') else ttk.Entry(row, textvariable=panel[key], width=9)
                widget.pack(side='left', padx=(4, 8))
            mapping = ttk.LabelFrame(f, text='列名映射：填写文件第一行实际列名（区分大小写）；后两项可留空', padding=6)
            mapping.pack(fill='x')
            panel['mapping'] = {}
            for i, (field, label) in enumerate(zip(FIELDS, ('日期时间 *', '时段用量 *', '表号 *', '通道（可选）', '质量（可选）'))):
                ttk.Label(mapping, text=label).grid(row=0, column=i, sticky='w', padx=4)
                variable = tk.StringVar(value=field)
                panel['mapping'][field] = variable
                ttk.Entry(mapping, textvariable=variable, width=23).grid(row=1, column=i, padx=4, sticky='ew')
                mapping.columnconfigure(i, weight=1)
            wrapper = ttk.Frame(f); wrapper.pack(fill='x', pady=6)
            listing = tk.Listbox(wrapper, selectmode='extended', exportselection=False, height=7)
            listing.pack(side='left', fill='both', expand=True)
            scroll = ttk.Scrollbar(wrapper, command=listing.yview); scroll.pack(side='right', fill='y')
            listing.configure(yscrollcommand=scroll.set)
            listing.bind('<<ListboxSelect>>', self.invalidate)
            panel['list'] = listing
            row = ttk.Frame(f); row.pack(fill='x')
            ttk.Button(row, text='全选', command=lambda c=category: self.select_all(c)).pack(side='left')
            ttk.Button(row, text='取消选择', command=lambda c=category: self.select_all(c, False)).pack(side='left', padx=6)
            ttk.Button(row, text='查看选中表的列名', command=lambda c=category: self.show_header(c)).pack(side='left')
            ttk.Label(row, text='Excel 各工作表分别列出，可 Ctrl / Shift 多选。').pack(side='left', padx=12)
            self.panels[category] = panel
        row = ttk.Frame(body); row.pack(fill='x')
        self.start, self.end = tk.StringVar(), tk.StringVar()
        for title, variable in [('开始日期', self.start), ('结束日期', self.end)]:
            ttk.Label(row, text=title).pack(side='left')
            ttk.Entry(row, textvariable=variable, width=13).pack(side='left', padx=6)
        ttk.Label(row, text='YYYY-MM-DD，留空不限').pack(side='left')
        ttk.Button(row, text='计算选中数据', command=self.run).pack(side='left', padx=12)
        ttk.Button(row, text='导出结果 CSV', command=self.export).pack(side='left')
        columns = ('分类', '日期', '记录数', '用量', '单位', '用量费', '固定费', '合计', '数据检查')
        self.columns = columns
        wrapper = ttk.Frame(body); wrapper.pack(fill='both', expand=True, pady=8)
        self.table = ttk.Treeview(wrapper, columns=columns, show='headings', height=9)
        for col in columns:
            self.table.heading(col, text=col)
            self.table.column(col, width=85 if col != '数据检查' else 380, minwidth=70)
        self.table.grid(row=0, column=0, sticky='nsew')
        vs = ttk.Scrollbar(wrapper, command=self.table.yview); vs.grid(row=0, column=1, sticky='ns')
        hs = ttk.Scrollbar(wrapper, orient='horizontal', command=self.table.xview); hs.grid(row=1, column=0, sticky='ew')
        self.table.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        wrapper.columnconfigure(0, weight=1); wrapper.rowconfigure(0, weight=1)
        self.summary = tk.StringVar(value='先输入费率或导入费率文件。')
        ttk.Label(body, textvariable=self.summary, wraplength=1280).pack(anchor='w', pady=5)
        ttk.Label(body, text='固定费按各分类实际有记录的日期收取；缺失日期不补计。费率需自行包含适用税费。\n仅按天计费请将用量单价设为0。不同费率期间请分开计算；没有分摊功能。').pack(anchor='w')
        self.restore()
        for variable in (self.currency, self.start, self.end):
            variable.trace_add('write', self.invalidate)
        for p in self.panels.values():
            for key in ('folder', 'usage_rate', 'daily_rate', 'data_unit', 'rate_unit', 'interval_minutes'):
                p[key].trace_add('write', self.invalidate)
            for variable in p['mapping'].values():
                variable.trace_add('write', self.invalidate)
        self.refresh()
        self.static_text = []
        self.capture_text(body)

    def tr(self, text):
        return translate(text, self.language)

    def error(self, exc):
        return self.tr(exc.message) if isinstance(exc, LocalizedError) else str(exc)

    def notify(self, kind, title, message):
        getattr(messagebox, kind)(self.tr(title), self.tr(message))

    def set_summary(self, message):
        self.summary_message = message
        self.summary.set(self.tr(message))

    def capture_text(self, widget):
        if widget is not self.language_button and 'text' in widget.keys() and not ('textvariable' in widget.keys() and widget.cget('textvariable')):
            self.static_text.append((widget, widget.cget('text')))
        for child in widget.winfo_children():
            self.capture_text(child)

    def status(self, row):
        return row.status_en if self.language == 'en' else row.status

    def toggle_language(self):
        self.language = 'en' if self.language == 'zh' else 'zh'
        self.language_button.configure(text='中文' if self.language == 'en' else 'English')
        self.root.title(self.tr('Utility Bill Calculator · 水电燃气计算器'))
        for widget, key in self.static_text:
            widget.configure(text=self.tr(key))
        for i, label in enumerate(CATEGORIES.values()):
            self.tabs.tab(i, text=self.tr(label))
        for col in self.columns:
            self.table.heading(col, text=self.tr(col))
        # Replace list item labels in place; keep selections, scroll positions and results.
        for p in self.panels.values():
            selected = p['list'].curselection()
            scroll = p['list'].yview()[0]
            p['list'].delete(0, 'end')
            for source in p['sources']:
                label = str(source.path) + (f'  [{self.tr("工作表")}: {source.sheet}]' if source.sheet is not None else '')
                p['list'].insert('end', label)
            for index in selected:
                p['list'].selection_set(index)
            p['list'].yview_moveto(scroll)
        if self.results:
            self.render_results()
        else:
            self.summary.set(self.tr(self.summary_message))

    def render_results(self):
        self.table.delete(*self.table.get_children())
        for r in self.results:
            self.table.insert('', 'end', values=(self.tr(CATEGORIES[r.category]), str(r.day), r.count, str(r.usage), r.unit,
                                                 str(r.usage_cost), str(r.daily_cost), money(r.total), self.status(r)))
        totals = {c: sum((r.total for r in self.results if r.category == c), Decimal(0)) for c in CATEGORIES}
        parts = [f'{self.tr(CATEGORIES[c])} {money(v)} {self.result_currency}' for c, v in totals.items()]
        parts.append(self.tr(Message('总计 {total} {currency}', total=money(sum(totals.values())), currency=self.result_currency)))
        self.summary.set(' | '.join(parts) + '\n' + '; '.join(self.tr(n) for n in self.notes) + ' ' + self.tr('缺失日期未计固定费。'))

    def invalidate(self, *_):
        self.results = []
        self.table.delete(*self.table.get_children())
        self.set_summary('选择或设置已更改，请重新计算。')

    def select_all(self, category, selected=True):
        listing = self.panels[category]['list']
        if selected:
            listing.selection_set(0, 'end')
        else:
            listing.selection_clear(0, 'end')
        self.invalidate()

    def set_sources(self, category, entries):
        p = self.panels[category]
        p['sources'] = sorted(set(entries), key=lambda s: (str(s.path), s.sheet or ''))
        p['list'].delete(0, 'end')
        for source in p['sources']:
            p['list'].insert('end', f'{source.path}'+ (f'  [{self.tr("工作表")}: {source.sheet}]' if source.sheet is not None else ''))

    def discover(self, category, paths, existing=()):
        entries, errors = list(existing), []
        for path in paths:
            try:
                entries.extend(sources(path))
            except Exception as exc:
                errors.append(f'{path.name}: {self.error(exc)}')
        self.set_sources(category, entries)
        return errors

    def refresh(self):
        self.invalidate()
        errors = []
        for category, p in self.panels.items():
            folder = Path(p['folder'].get())
            try:
                paths = [f for f in folder.rglob('*') if f.is_file() and f.suffix.lower() in EXTENSIONS and not f.name.startswith('~$')] if folder.is_dir() else []
                errors.extend(self.discover(category, paths))
            except Exception as exc:
                errors.append(self.error(exc))
        if errors:
            self.notify('showwarning', '部分文件未读取', '\n'.join(errors))

    def choose_folder(self, category):
        chosen = filedialog.askdirectory(initialdir=BASE, title=self.tr('选择文件夹'))
        if chosen:
            self.panels[category]['folder'].set(chosen)
            self.refresh()

    def add_files(self, category):
        chosen = filedialog.askopenfilenames(initialdir=BASE, title=self.tr('添加文件'), filetypes=[('CSV / Excel', '*.csv *.xlsx *.xls *.xlsm *.xlsb')])
        if chosen:
            self.invalidate()
            errors = self.discover(category, [Path(p) for p in chosen], self.panels[category]['sources'])
            if errors:
                self.notify('showwarning', '部分文件未读取', '\n'.join(errors))

    def show_header(self, category):
        p = self.panels[category]
        try:
            indices = p['list'].curselection()
            if not indices:
                raise LocalizedError('请先选中一个文件或工作表。')
            table = read_table(p['sources'][indices[0]])
            self.notify('showinfo', '第一行列名', '\n'.join(str(v) for v in table[0]) if table else '空表')
        except Exception as exc:
            self.notify('showerror', '读取失败', self.error(exc))

    def import_rates(self):
        chosen = filedialog.askopenfilename(initialdir=BASE, title=self.tr('导入费率 JSON'), filetypes=[(self.tr('费率 JSON'), '*.json')])
        if not chosen:
            return
        try:
            currency, tariffs = load_tariffs(chosen)
            self.currency.set(currency)
            for category, tariff in tariffs.items():
                for key, value in tariff.as_dict().items():
                    self.panels[category][key].set(str(value))
            self.invalidate()
            self.set_summary('费率已导入；如需下次保留，请点击保存设置。')
        except Exception as exc:
            self.notify('showerror', '费率导入失败', self.error(exc))

    def save(self):
        data = {'currency': self.currency.get(), 'categories': {}}
        for category, p in self.panels.items():
            values = {key: p[key].get() for key in ('folder', 'usage_rate', 'daily_rate', 'data_unit', 'rate_unit', 'interval_minutes')}
            values['mapping'] = {k: v.get() for k, v in p['mapping'].items()}
            data['categories'][category] = values
        try:
            SETTINGS.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
            self.notify('showinfo', '设置', '设置已保存到 settings.local.json。')
        except Exception as exc:
            self.notify('showerror', '保存失败', self.error(exc))

    def restore(self):
        if not SETTINGS.exists():
            return
        try:
            data = json.loads(SETTINGS.read_text(encoding='utf-8'))
            # Validate structure before mutating widgets.
            if not isinstance(data.get('currency'), str) or not isinstance(data.get('categories'), dict):
                raise LocalizedError('设置格式不正确')
            for category, values in data['categories'].items():
                if category not in self.panels or not isinstance(values, dict) or not isinstance(values.get('mapping', {}), dict):
                    raise LocalizedError('分类设置格式不正确')
            self.currency.set(data['currency'])
            for category, values in data['categories'].items():
                for key in ('folder', 'usage_rate', 'daily_rate', 'data_unit', 'rate_unit', 'interval_minutes'):
                    if key in values:
                        self.panels[category][key].set(values[key])
                for key, value in values.get('mapping', {}).items():
                    if key in FIELDS:
                        self.panels[category]['mapping'][key].set(value)
        except Exception as exc:
            self.notify('showwarning', '设置读取失败', self.error(exc))

    def run(self):
        self.invalidate()
        try:
            currency = self.currency.get().strip().upper()
            if len(currency) != 3 or not currency.isascii() or not currency.isalpha():
                raise LocalizedError('币种应为三个英文字母，如 AUD。')
            start = date.fromisoformat(self.start.get().strip()) if self.start.get().strip() else None
            end = date.fromisoformat(self.end.get().strip()) if self.end.get().strip() else None
            if start and end and start > end:
                raise LocalizedError('开始日期不能晚于结束日期。')
            results, notes, used = [], [], set()
            for category, p in self.panels.items():
                selected = [p['sources'][i] for i in p['list'].curselection()]
                if not selected:
                    continue
                keys = {(s.path.resolve(), s.sheet) for s in selected}
                if used & keys:
                    raise LocalizedError('同一文件 / 工作表被分到多个类别，请取消重复选择。')
                used |= keys
                tariff = Tariff.parse({key: p[key].get() for key in ('usage_rate', 'daily_rate', 'data_unit', 'rate_unit', 'interval_minutes')})
                rows, duplicates = calculate(category, selected, tariff, {k: v.get() for k, v in p['mapping'].items()}, start, end)
                results.extend(rows)
                if rows:
                    missing = ((end or rows[-1].day) - (start or rows[0].day)).days + 1 - len(rows)
                    notes.append(Message('{category} {days}天，缺{missing}天，去重{duplicates}条', category=Message(CATEGORIES[category]), days=len(rows), missing=missing, duplicates=duplicates))
                else:
                    notes.append(Message('{category}：所选范围无记录', category=Message(CATEGORIES[category])))
            if not results:
                raise LocalizedError('没有可计算记录，请检查文件选择与日期。')
            self.results = sorted(results, key=lambda r: (r.day, r.category))
            self.result_currency = currency
            self.notes = notes
            self.render_results()
        except Exception as exc:
            self.notify('showerror', '无法计算', self.error(exc))

    def export(self):
        if not self.results:
            self.notify('showinfo', '导出', '请先计算。'); return
        path = filedialog.asksaveasfilename(initialdir=BASE / 'exports', title=self.tr('导出结果 CSV'), initialfile='utility-costs.csv', defaultextension='.csv', filetypes=[('CSV', '*.csv')])
        if path:
            try:
                with open(path, 'w', encoding='utf-8-sig', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow((*[self.tr(c) for c in self.columns], self.tr('币种')))
                    for r in self.results:
                        writer.writerow((self.tr(CATEGORIES[r.category]), r.day, r.count, r.usage, r.unit, r.usage_cost,
                                         r.daily_cost, r.total, self.status(r), self.result_currency))
                    writer.writerow((self.tr('合计'), '', '', '', '', '', '', money(sum(r.total for r in self.results)),
                                     self.tr('总额按未舍入金额求和后四舍五入至分'), self.result_currency))
            except Exception as exc:
                self.notify('showerror', '导出失败', self.error(exc))


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
