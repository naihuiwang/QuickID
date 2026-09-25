"""Cross-platform QuickID desktop application."""
import argparse
from decimal import Decimal, InvalidOperation
import json
import os
from pathlib import Path
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from quickid_core import default_settings, process

SETTINGS = Path.home() / '.quickid' / 'settings.json'
MARKERS = (('P1', 'P1'), ('A', 'A'), ("A'", 'A1'), ('B', 'B'), ('C', 'C'),
           ('P2', 'P2'), ('D', 'D'), ('E', 'E'), ('F', 'F'), ("F'", 'F1'),
           ('G', 'G'), ("G'", 'G1'))
SETTINGS_VERSION = 4


def load_settings():
    values = default_settings()
    if SETTINGS.exists():
        saved = json.loads(SETTINGS.read_text(encoding='utf-8'))
        # Earlier releases saved on every Run; only explicit opt-in is restored.
        if saved.get('_remember_settings') == '1':
            values.update({key: value for key, value in saved.items() if key in values})
            values['remember_settings'] = '1'
    return values


def save_settings(values):
    SETTINGS.parent.mkdir(parents=True, exist_ok=True)
    SETTINGS.write_text(json.dumps({**values, 'remember_settings': '1',
                                    '_defaults_version': SETTINGS_VERSION,
                                    '_remember_settings': '1'}, indent=2), encoding='utf-8')


def stop_remembering_settings():
    if SETTINGS.exists():
        saved = json.loads(SETTINGS.read_text(encoding='utf-8'))
        if saved.get('_remember_settings') == '1':
            saved['_remember_settings'] = '0'
            SETTINGS.write_text(json.dumps(saved, indent=2), encoding='utf-8')


def checked_settings(values):
    """Do not silently round a user-entered mass tolerance."""
    values = dict(values)
    try:
        minimum = int(values['tbminrows'])
        if minimum < 0:
            raise ValueError
    except ValueError:
        raise ValueError('Minimum matched markers must be a nonnegative whole number') from None
    values['tbminrows'] = str(minimum)

    enabled = values.get('peak_warning_enabled', '1').strip()
    if enabled not in ('0', '1'):
        raise ValueError('Peak-count warning must be enabled or disabled')
    values['peak_warning_enabled'] = enabled
    if enabled == '1':
        raw_limit = values.get('peak_warning_threshold', '').strip()
        if not raw_limit.isdecimal() or int(raw_limit) < 1:
            raise ValueError('Enter a positive whole-number peak-count warning threshold')
        values['peak_warning_threshold'] = str(int(raw_limit))

    names = []
    for _, key in MARKERS:
        field = 'marker_' + key + '_name'
        name = values[field].strip()
        values[field] = name
        names.append(name.casefold())
    if any(names) and not all(names):
        raise ValueError('Fill all twelve built-in marker names or leave all twelve blank')
    if names[0] and len(names) != len(set(names)):
        raise ValueError('Built-in marker names must be unique')

    def check_pair(label, low_key, high_key):
        try:
            low, high = Decimal(values[low_key]), Decimal(values[high_key])
            if not low.is_finite() or not high.is_finite():
                raise ValueError
            if (low != low.quantize(Decimal('0.1')) or
                    high != high.quantize(Decimal('0.1')) or low > high):
                raise ValueError
        except (InvalidOperation, ValueError):
            raise ValueError(f'{label}: enter a valid range with no more than one decimal place') from None
        values[low_key], values[high_key] = f'{low:.1f}', f'{high:+.1f}'

    if names[0]:
        for label, key in MARKERS:
            check_pair(values['marker_' + key + '_name'], 'tb' + key + '1', 'tb' + key + '2')
    for index in range(1, 9):
        prefix = f'custom{index}_'
        if all(values.get(prefix + part, '').strip() for part in ('name', 'low', 'high')):
            check_pair(values[prefix + 'name'], prefix + 'low', prefix + 'high')
    return values


class App(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=14)
        self.pack(fill='both', expand=True)
        self.vars = {key: tk.StringVar(value=value) for key, value in load_settings().items()}
        self.columnconfigure(1, weight=1)
        ttk.Label(self, text='QuickID', font=('', 18, 'bold')).grid(row=0, column=0, columnspan=3, sticky='w', pady=(0, 12))
        ttk.Label(self, text='Reference file (.xlsx)').grid(row=1, column=0, sticky='w')
        ttk.Entry(self, textvariable=self.vars['reference']).grid(row=1, column=1, sticky='ew', padx=8)
        ttk.Button(self, text='Browse', command=self.pick_reference).grid(row=1, column=2)
        ttk.Label(self, text='Peak list folder (.txt/.tab)').grid(row=2, column=0, sticky='w')
        ttk.Entry(self, textvariable=self.vars['indir']).grid(row=2, column=1, sticky='ew', padx=8)
        ttk.Button(self, text='Browse', command=self.pick_folder).grid(row=2, column=2)
        ttk.Label(self, text='Save results as .csv').grid(row=3, column=0, sticky='w', pady=6)
        ttk.Entry(self, textvariable=self.vars['outdir']).grid(row=3, column=1, sticky='ew', padx=8)
        ttk.Button(self, text='Browse', command=self.pick_output).grid(row=3, column=2)
        minimum_row = ttk.Frame(self)
        minimum_row.grid(row=4, column=0, columnspan=3, sticky='w')
        ttk.Label(minimum_row, text='Minimum matched markers (results shown from this number)').pack(side='left')
        ttk.Entry(minimum_row, textvariable=self.vars['tbminrows'], width=8).pack(side='left', padx=(8, 0))
        warning = ttk.Frame(self)
        warning.grid(row=5, column=0, columnspan=3, sticky='w', pady=(6, 0))
        ttk.Checkbutton(warning, text='Warn if peak count exceeds',
                        variable=self.vars['peak_warning_enabled'], onvalue='1', offvalue='0',
                        command=self.update_peak_warning_field).pack(side='left')
        self.peak_warning_entry = ttk.Entry(
            warning, textvariable=self.vars['peak_warning_threshold'], width=8)
        self.peak_warning_entry.pack(side='left', padx=5)
        ttk.Label(warning, text='(possible false positive)').pack(side='left')
        self.update_peak_warning_field()
        unit_label = ttk.Frame(self)
        unit_label.grid(row=6, column=0, columnspan=3, sticky='w', pady=(10, 0))
        ttk.Label(unit_label, text='Mass tolerance by marker (').pack(side='left')
        ttk.Label(unit_label, text='m/z', font=('', 10, 'italic')).pack(side='left')
        ttk.Label(unit_label, text=')').pack(side='left')
        panel = ttk.LabelFrame(self, padding=8)
        panel.grid(row=7, column=0, columnspan=3, sticky='ew', pady=(0, 10))
        for index, (label, key) in enumerate(MARKERS):
            row, column = divmod(index, 4)
            cell = ttk.Frame(panel)
            cell.grid(row=row, column=column, sticky='w', padx=8, pady=5)
            ttk.Entry(cell, textvariable=self.vars['marker_' + key + '_name'], width=8).pack(side='left')
            for suffix in ('1', '2'):
                ttk.Entry(cell, textvariable=self.vars['tb' + key + suffix], width=7).pack(side='left', padx=2)
        ttk.Button(panel, text='Clear all', command=self.clear_builtin_markers).grid(
            row=3, column=3, sticky='e', padx=8, pady=(4, 0))
        ttk.Separator(panel, orient='horizontal').grid(row=4, column=0, columnspan=4, sticky='ew', pady=8)
        ttk.Label(panel, text='Optional markers: name / lower / upper').grid(
            row=5, column=0, columnspan=4, sticky='w', padx=8)
        for index in range(8):
            row, column = divmod(index, 4)
            prefix = f'custom{index + 1}_'
            cell = ttk.Frame(panel)
            cell.grid(row=row + 6, column=column, sticky='w', padx=8, pady=5)
            ttk.Entry(cell, textvariable=self.vars[prefix + 'name'], width=8).pack(side='left')
            for suffix in ('low', 'high'):
                ttk.Entry(cell, textvariable=self.vars[prefix + suffix], width=7).pack(side='left', padx=2)
        buttons = ttk.Frame(self)
        buttons.grid(row=8, column=0, columnspan=3, sticky='w', pady=4)
        ttk.Button(buttons, text='Run', command=self.run).pack(side='left', padx=(0, 7))
        ttk.Button(buttons, text='Save settings', command=self.save).pack(side='left', padx=(0, 7))
        ttk.Button(buttons, text='Open results', command=self.open_results).pack(side='left')
        ttk.Checkbutton(buttons, text='Keep settings after Run',
                        variable=self.vars['remember_settings'], onvalue='1', offvalue='0',
                        command=self.remember_changed).pack(side='left', padx=(12, 0))
        self.log = tk.Text(self, height=11, wrap='word')
        self.log.grid(row=9, column=0, columnspan=3, sticky='nsew', pady=6)
        self.rowconfigure(9, weight=1)

    def update_peak_warning_field(self):
        self.peak_warning_entry.configure(
            state='normal' if self.vars['peak_warning_enabled'].get() == '1' else 'disabled')

    def clear_builtin_markers(self):
        for _, key in MARKERS:
            for field in ('marker_' + key + '_name', 'tb' + key + '1', 'tb' + key + '2'):
                self.vars[field].set('')

    def pick_reference(self):
        path = filedialog.askopenfilename(filetypes=[('Excel workbook', '*.xlsx')])
        if path:
            self.vars['reference'].set(path)

    def pick_folder(self):
        path = filedialog.askdirectory(initialdir=self.vars['indir'].get() or str(Path.home()))
        if path:
            self.vars['indir'].set(path)

    def pick_output(self):
        path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV', '*.csv')])
        if path:
            self.vars['outdir'].set(path)

    def values(self):
        return {key: value.get().strip() for key, value in self.vars.items()}

    def remember_changed(self):
        if self.vars['remember_settings'].get() == '0':
            stop_remembering_settings()

    def save(self):
        try:
            settings = checked_settings(self.values())
            settings['remember_settings'] = '1'
            save_settings(settings)
            self.vars['remember_settings'].set('1')
            messagebox.showinfo('QuickID', 'Settings saved')
        except Exception as exc:
            messagebox.showerror('QuickID', str(exc))

    def run(self):
        self.log.delete('1.0', 'end')
        try:
            settings = checked_settings(self.values())
            if not settings['outdir']:
                raise ValueError('Select an output .csv file')
            count = process(settings['indir'], settings['outdir'], settings, self.report)
            if settings['remember_settings'] == '1':
                save_settings(settings)
            messagebox.showinfo('QuickID', f'Done: {count} peak list(s). Check assignments against spectra.')
        except Exception as exc:
            messagebox.showerror('QuickID', str(exc))

    def report(self, message):
        self.log.insert('end', message + '\n')
        self.log.see('end')
        self.update_idletasks()

    def open_results(self):
        path = self.vars['outdir'].get().strip()
        if not Path(path).is_file():
            messagebox.showerror('QuickID', 'No output file found')
            return
        if sys.platform == 'darwin':
            subprocess.Popen(['open', path])
        elif os.name == 'nt':
            os.startfile(path)
        else:
            subprocess.Popen(['xdg-open', path])


def main():
    parser = argparse.ArgumentParser(description='QuickID ZooMS marker matching')
    parser.add_argument('--input-dir', help='Run without GUI using this folder')
    parser.add_argument('--reference', help='Reference .xlsx file (with --input-dir)')
    parser.add_argument('--output', help='CSV output path (with --input-dir)')
    args = parser.parse_args()
    if args.input_dir:
        if not args.output:
            parser.error('--output is required with --input-dir')
        process(args.input_dir, args.output, load_settings(), print, args.reference)
    else:
        root = tk.Tk()
        root.title('QuickID')
        root.minsize(850, 680)
        App(root)
        root.mainloop()


if __name__ == '__main__':
    main()
