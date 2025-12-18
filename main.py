import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from DrissionPage import WebPage
from tablib import Dataset

DEFAULT_URL = 'https://sieportal.siemens.com/en-cn/search?scope=knowledgebase&Type=siePortal&SearchTerm=&SortingOption=CreationDateDesc&EntryTypes=all&Page=0&PageSize=20&ProductId=424821'


def run(input_path: str, output_path: str, url: str = DEFAULT_URL, log=print):
    data = Dataset().load(open(input_path, 'r', encoding='utf-8').read(), format='csv', headers=False)

    page = WebPage()
    page.get(url)
    output = Dataset()
    for i in data:
        item = i[0]
        log(f'Processing: {item}')
        try:
            page.ele('#search-field__input').input(item)
            # data-cy="header-search-field"
            page.ele('css:[type="submit"]').click()

            href = page.ele('css:[title="Technical Data"]').parent().parent().attr('href')
            log(href)
            output.append([item, href])
        except Exception as e:
            log(f"Error processing {item}: {e}")
            output.append([item, 'Error'])
            page.get(url)

    # export csv
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(output.export('csv'))
    log(f'Done. Saved to {output_path}')


def start_gui():
    root = tk.Tk()
    root.title('Siemens Technical Data Search')

    input_var = tk.StringVar(value='data.csv')
    output_var = tk.StringVar(value='output.csv')
    url_var = tk.StringVar(value=DEFAULT_URL)

    frm = ttk.Frame(root, padding=12)
    frm.grid(sticky='nsew')
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    ttk.Label(frm, text='Input CSV:').grid(row=0, column=0, sticky='w')
    in_entry = ttk.Entry(frm, textvariable=input_var, width=40)
    in_entry.grid(row=0, column=1, sticky='ew')

    def browse_in():
        path = filedialog.askopenfilename(title='Select input CSV', filetypes=[('CSV files', '*.csv'), ('All files', '*.*')])
        if path:
            input_var.set(path)

    ttk.Button(frm, text='Browse...', command=browse_in).grid(row=0, column=2)

    ttk.Label(frm, text='Output CSV:').grid(row=1, column=0, sticky='w')
    out_entry = ttk.Entry(frm, textvariable=output_var, width=40)
    out_entry.grid(row=1, column=1, sticky='ew')

    def browse_out():
        path = filedialog.asksaveasfilename(title='Save output CSV', defaultextension='.csv', filetypes=[('CSV files', '*.csv')])
        if path:
            output_var.set(path)

    ttk.Button(frm, text='Browse...', command=browse_out).grid(row=1, column=2)

    ttk.Label(frm, text='Search URL:').grid(row=2, column=0, sticky='w')
    url_entry = ttk.Entry(frm, textvariable=url_var, width=60)
    url_entry.grid(row=2, column=1, columnspan=2, sticky='ew')

    log_text = tk.Text(frm, height=12, width=60, wrap='word')
    log_text.grid(row=3, column=0, columnspan=3, sticky='nsew', pady=(8, 0))

    frm.columnconfigure(1, weight=1)
    frm.rowconfigure(3, weight=1)

    def append_log(msg: str):
        log_text.insert('end', msg + '\n')
        log_text.see('end')
        root.update_idletasks()

    start_btn = ttk.Button(frm, text='Start')
    start_btn.grid(row=4, column=0, columnspan=3, pady=8)

    def on_start():
        input_path = input_var.get()
        output_path = output_var.get()
        url = url_var.get()
        if not input_path:
            messagebox.showerror('Error', 'Please select input CSV')
            return
        start_btn.config(state='disabled')

        def worker():
            try:
                run(input_path, output_path, url, log=append_log)
                messagebox.showinfo('Completed', f'Finished. Output saved to:\n{output_path}')
            except Exception as e:
                messagebox.showerror('Error', str(e))
            finally:
                start_btn.config(state='normal')

        threading.Thread(target=worker, daemon=True).start()

    start_btn.config(command=on_start)

    root.mainloop()


if __name__ == '__main__':
    start_gui()
