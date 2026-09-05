import json
import tempfile
import tkinter as tk
import unittest
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

import app
from billing import sources


class AppTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError:
            self.skipTest('GUI tests require a display')
        self.root.withdraw()
        self.addCleanup(self.root.destroy)
        self.tmp = tempfile.TemporaryDirectory(); self.addCleanup(self.tmp.cleanup)
        self.settings = Path(self.tmp.name) / 'settings.json'
        patcher = patch.object(app, 'SETTINGS', self.settings)
        patcher.start(); self.addCleanup(patcher.stop)
        self.gui = app.App(self.root)
        self.root.update_idletasks()
        self.errors = patch.object(app.messagebox, 'showerror').start()
        self.addCleanup(patch.stopall)

    def rates(self):
        with patch.object(app.filedialog, 'askopenfilename', return_value=str(app.BASE / 'examples/rates.example.json')):
            self.gui.import_rates()

    def test_import_calculate_export_and_invalidation(self):
        self.rates()
        self.gui.set_sources('electricity', sources(app.BASE / 'examples/usage.example.csv'))
        self.gui.select_all('electricity')
        self.gui.run()
        self.errors.assert_not_called()
        self.assertEqual(sum(r.total for r in self.gui.results), Decimal('2.21'))
        dest = Path(self.tmp.name) / 'export.csv'
        with patch.object(app.filedialog, 'asksaveasfilename', return_value=str(dest)):
            self.gui.export()
        self.assertIn('2.21', dest.read_text(encoding='utf-8-sig'))
        self.gui.panels['electricity']['usage_rate'].set('.5')
        self.assertEqual(self.gui.results, [])

    def test_settings_mapping_restored(self):
        self.rates()
        self.gui.panels['water']['mapping']['ReadValue'].set('用量')
        with patch.object(app.messagebox, 'showinfo'):
            self.gui.save()
        self.gui.panels['water']['mapping']['ReadValue'].set('Other')
        self.gui.restore()
        self.assertEqual(self.gui.panels['water']['mapping']['ReadValue'].get(), '用量')

    def test_invalid_import_is_atomic(self):
        self.rates()
        before = self.gui.panels['electricity']['usage_rate'].get()
        data = json.loads((app.BASE / 'examples/rates.example.json').read_text())
        data['categories']['electricity']['usage_rate'] = '999'
        data['categories']['gas']['daily_rate'] = '-1'
        path = Path(self.tmp.name) / 'invalid.json'; path.write_text(json.dumps(data))
        with patch.object(app.filedialog, 'askopenfilename', return_value=str(path)):
            self.gui.import_rates()
        self.errors.assert_called_once()
        self.assertEqual(self.gui.panels['electricity']['usage_rate'].get(), before)

    def test_language_switch_preserves_state_and_translates_export(self):
        self.assertEqual(self.gui.language, 'zh')
        self.rates()
        self.gui.set_sources('electricity', sources(app.BASE / 'examples/usage.example.csv'))
        self.gui.select_all('electricity')
        self.root.update()
        self.gui.run()
        before = list(self.gui.results)
        self.gui.toggle_language()
        self.root.update()
        self.assertEqual(self.gui.results, before)
        self.assertEqual(self.gui.panels['electricity']['list'].curselection(), (0,))
        self.assertEqual(self.gui.panels['electricity']['mapping']['ReadValue'].get(), 'ReadValue')
        self.assertEqual(self.gui.tabs.tab(0, 'text'), 'Electricity')
        self.assertIn('Total 2.21 AUD', self.gui.summary.get())
        first = self.gui.table.item(self.gui.table.get_children()[0], 'values')
        self.assertEqual(first[0], 'Electricity')
        self.assertIn('Incomplete coverage', first[-1])
        dest = Path(self.tmp.name) / 'english.csv'
        with patch.object(app.filedialog, 'asksaveasfilename', return_value=str(dest)):
            self.gui.export()
        output = dest.read_text(encoding='utf-8-sig')
        self.assertTrue(output.startswith('Category,Date,Records'))
        self.assertNotRegex(output, r'[\u4e00-\u9fff]')
        self.gui.toggle_language()
        self.root.update()
        self.assertEqual(self.gui.results, before)
        self.assertEqual(self.gui.tabs.tab(0, 'text'), '电费')
        self.assertIn('总计 2.21 AUD', self.gui.summary.get())

    def test_english_errors_and_custom_column_names(self):
        self.gui.toggle_language()
        self.rates()
        self.gui.set_sources('electricity', sources(app.BASE / 'examples/usage.example.csv'))
        self.gui.select_all('electricity')
        self.gui.panels['electricity']['usage_rate'].set('bad')
        self.gui.run()
        title, body = self.errors.call_args.args
        self.assertEqual(title, 'Calculation failed')
        self.assertEqual(body, 'Usage rate must be a number.')
        from billing import read_records
        from i18n import LocalizedError
        source = sources(app.BASE / 'examples/custom-columns.example.csv')[0]
        with self.assertRaises(LocalizedError) as error:
            list(read_records(source, {'ReadDatetime': '日期', 'ReadValue': '用量', 'MeterSerial': '表号'}))
        message = self.gui.error(error.exception)
        self.assertIn('missing column 日期', message)
        self.assertNotIn('缺少列', message)

    def test_static_labels_all_have_english_translations(self):
        from i18n import EN
        import re
        for _, key in self.gui.static_text:
            if re.search(r'[\u4e00-\u9fff]', key):
                self.assertIn(key, EN)
        self.gui.toggle_language()
        for widget, _ in self.gui.static_text:
            self.assertNotRegex(widget.cget('text'), r'[\u4e00-\u9fff]')


if __name__ == '__main__':
    unittest.main()
