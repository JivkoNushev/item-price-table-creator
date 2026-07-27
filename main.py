#!/usr/bin/env python3
import tkinter as tk
from tkinter import ttk, messagebox
import os

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
except ImportError:
    print("openpyxl is not installed. Run: pip install openpyxl")
    exit(1)

XLSX_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "entries.xlsx")


class AutocompleteEntry(tk.Frame):
    def __init__(self, parent, suggestions=None, **kwargs):
        entry_kwargs = {k: v for k, v in kwargs.items() if k in ("font", "width")}
        frame_kwargs = {k: v for k, v in kwargs.items() if k not in ("font", "width")}
        super().__init__(parent, **frame_kwargs)
        self._suggestions = suggestions if suggestions is not None else []
        self._filtered = []
        self._visible = False
        self._hide_after_id = None

        self.entry = tk.Entry(self, **entry_kwargs)
        self.entry.pack(fill=tk.X, expand=True)
        self.entry.bind("<KeyRelease>", self._on_key_release)
        self.entry.bind("<FocusOut>", self._on_focus_out)
        self.entry.bind("<Down>", self._on_arrow_down)

        root = self.winfo_toplevel()
        self.listbox = tk.Listbox(root, height=6, bd=1, relief=tk.SOLID, font=entry_kwargs.get("font", ("Segoe UI", 10)))
        self.listbox.bind("<ButtonRelease-1>", self._on_select)
        self.listbox.bind("<Down>", self._on_listbox_down)
        self.listbox.bind("<Up>", self._on_listbox_up)
        self.listbox.bind("<Return>", self._on_select)
        self.listbox.bind("<Escape>", lambda e: self._hide_listbox())
        self.listbox.bind("<Enter>", self._on_listbox_enter)
        self.listbox.bind("<Leave>", self._on_listbox_leave)

    def set_suggestions(self, suggestions):
        self._suggestions = list(suggestions)

    def _on_key_release(self, event):
        if event.keysym in ("Up", "Down", "Return", "Escape", "Tab"):
            return
        text = self.entry.get().strip()
        if not text:
            self._hide_listbox()
            return
        lower_text = text.lower()
        self._filtered = [s for s in self._suggestions if lower_text in s.lower()]
        if not self._filtered:
            self._hide_listbox()
            return
        self.listbox.delete(0, tk.END)
        for item in self._filtered:
            self.listbox.insert(tk.END, item)
        self._show_listbox()

    def _on_arrow_down(self, event):
        if self._visible and self.listbox.size() > 0:
            self.listbox.focus_set()
            self.listbox.select_set(0)

    def _on_listbox_down(self, event):
        sel = self.listbox.curselection()
        if sel and sel[0] < self.listbox.size() - 1:
            self.listbox.select_clear(0, tk.END)
            self.listbox.select_set(sel[0] + 1)
            self.listbox.see(sel[0] + 1)

    def _on_listbox_up(self, event):
        sel = self.listbox.curselection()
        if sel and sel[0] > 0:
            self.listbox.select_clear(0, tk.END)
            self.listbox.select_set(sel[0] - 1)
            self.listbox.see(sel[0] - 1)

    def _on_select(self, event):
        sel = self.listbox.curselection()
        if sel:
            self.entry.delete(0, tk.END)
            self.entry.insert(0, self.listbox.get(sel[0]))
            self._hide_listbox()
            self.entry.focus_set()

    def _on_focus_out(self, event):
        self._hide_after_id = self.after(200, self._hide_listbox)

    def _on_listbox_enter(self, event):
        if self._hide_after_id:
            self.after_cancel(self._hide_after_id)
            self._hide_after_id = None

    def _on_listbox_leave(self, event):
        self._hide_after_id = self.after(200, self._hide_listbox)

    def _show_listbox(self):
        if self._hide_after_id:
            self.after_cancel(self._hide_after_id)
            self._hide_after_id = None
        root = self.winfo_toplevel()
        root_x = root.winfo_rootx()
        root_y = root.winfo_rooty()
        x = self.entry.winfo_rootx() - root_x
        y = self.entry.winfo_rooty() + self.entry.winfo_height() - root_y
        w = self.entry.winfo_width()
        self.listbox.place(x=x, y=y, width=w)
        self.listbox.lift()
        self._visible = True

    def _hide_listbox(self):
        self._hide_after_id = None
        if self._visible:
            self.listbox.place_forget()
            self._visible = False

    def get(self):
        return self.entry.get()

    def delete(self, *args):
        self.entry.delete(*args)

    def insert(self, *args):
        self.entry.insert(*args)


class EntryRow(tk.Frame):
    def __init__(self, parent, model_sku, price, price_sample, delete_cb, save_cb, entries_ref, original_index, **kwargs):
        super().__init__(parent, bd=1, relief=tk.RIDGE, **kwargs)
        self.model_sku = model_sku
        self.price = price
        self.price_sample = price_sample
        self.delete_cb = delete_cb
        self.save_cb = save_cb
        self.entries_ref = entries_ref
        self.original_index = original_index
        self._editing = False
        self._build_view()

    def _build_view(self):
        for w in self.winfo_children():
            w.destroy()

        self.columnconfigure(0, weight=3)
        self.columnconfigure(1, weight=1)
        self.columnconfigure(2, weight=1)
        self.columnconfigure(3, weight=0)
        self.columnconfigure(4, weight=0)

        price_display = f"{self.price} €" if self.price else "—"
        price_sample_display = f"{self.price_sample} €" if self.price_sample else "—"

        self.lbl_sku = tk.Label(self, text=self.model_sku, anchor="w", padx=5, pady=4, font=("Segoe UI", 10))
        self.lbl_sku.grid(row=0, column=0, sticky="ew")

        self.lbl_price = tk.Label(self, text=price_display, anchor="w", padx=5, pady=4, font=("Segoe UI", 10))
        self.lbl_price.grid(row=0, column=1, sticky="ew")

        self.lbl_price_sample = tk.Label(self, text=price_sample_display, anchor="w", padx=5, pady=4, font=("Segoe UI", 10))
        self.lbl_price_sample.grid(row=0, column=2, sticky="ew")

        btn_frame = tk.Frame(self)
        btn_frame.grid(row=0, column=3, padx=2)

        self.btn_edit = tk.Button(btn_frame, text="Edit", width=5, command=self._start_edit)
        self.btn_edit.pack(side=tk.LEFT, padx=1)

        self.btn_delete = tk.Button(btn_frame, text="Del", width=4, fg="red",
                                     command=lambda: self.delete_cb(self))
        self.btn_delete.pack(side=tk.LEFT, padx=1)

    def _start_edit(self):
        if self._editing:
            return
        self._editing = True

        for w in self.winfo_children():
            w.destroy()

        self.columnconfigure(0, weight=3)
        self.columnconfigure(1, weight=1)
        self.columnconfigure(2, weight=1)
        self.columnconfigure(3, weight=0)
        self.columnconfigure(4, weight=0)

        self.ent_sku = tk.Entry(self, font=("Segoe UI", 10))
        self.ent_sku.insert(0, self.model_sku)
        self.ent_sku.grid(row=0, column=0, sticky="ew", padx=1, pady=2)

        self.ent_price = tk.Entry(self, font=("Segoe UI", 10), width=10)
        self.ent_price.insert(0, self.price if self.price else "")
        self.ent_price.grid(row=0, column=1, sticky="ew", padx=1, pady=2)

        self.ent_price_sample = tk.Entry(self, font=("Segoe UI", 10), width=10)
        self.ent_price_sample.insert(0, self.price_sample if self.price_sample else "")
        self.ent_price_sample.grid(row=0, column=2, sticky="ew", padx=1, pady=2)

        btn_frame = tk.Frame(self)
        btn_frame.grid(row=0, column=3, padx=2)

        self.btn_save = tk.Button(btn_frame, text="Save", width=5, fg="green", command=self._save_edit)
        self.btn_save.pack(side=tk.LEFT, padx=1)

        self.btn_cancel = tk.Button(btn_frame, text="X", width=4, command=self._cancel_edit)
        self.btn_cancel.pack(side=tk.LEFT, padx=1)

        self.ent_sku.focus_set()

    def _save_edit(self):
        new_sku = self.ent_sku.get().strip()
        new_price = self.ent_price.get().strip()
        new_price_sample = self.ent_price_sample.get().strip()

        if not new_sku:
            messagebox.showwarning("Warning", "Модел SKU cannot be empty.")
            return

        if not is_valid_price(new_price):
            messagebox.showwarning("Warning", "Цена must be a valid positive number.")
            return

        if new_price_sample and not is_valid_price(new_price_sample):
            messagebox.showwarning("Warning", "Цена на мостра must be a valid positive number if provided.")
            return

        new_entry = (new_sku, format_price(new_price), format_price(new_price_sample) if new_price_sample else "")
        for i, e in enumerate(self.entries_ref):
            if i != self.original_index and e == new_entry:
                messagebox.showwarning("Duplicate", "This entry already exists.")
                return

        old_entry = (self.model_sku, self.price, self.price_sample)
        for i, e in enumerate(self.entries_ref):
            if i == self.original_index:
                self.entries_ref[i] = new_entry
                break

        self.model_sku = new_entry[0]
        self.price = new_entry[1]
        self.price_sample = new_entry[2]
        self._editing = False
        self._build_view()
        self.save_cb()

    def _cancel_edit(self):
        self._editing = False
        self._build_view()

    def get_data(self):
        return (self.model_sku, self.price, self.price_sample)


def is_valid_price(value):
    if not value:
        return False
    try:
        v = float(value.replace(",", "."))
        return v >= 0
    except ValueError:
        return False


def format_price(value):
    if not value:
        return ""
    try:
        v = float(value.replace(",", "."))
        if v == int(v):
            return str(int(v))
        return f"{v:.2f}"
    except ValueError:
        return value


def load_entries_from_xlsx():
    entries = []
    if not os.path.exists(XLSX_FILE):
        return entries
    try:
        wb = load_workbook(XLSX_FILE)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                continue
            sku = str(row[0]).strip()
            price_raw = str(row[1]).replace("€", "").strip() if row[1] else ""
            price_sample_raw = str(row[2]).replace("€", "").strip() if len(row) > 2 and row[2] else ""
            if price_sample_raw in ("—", "-", ""):
                price_sample_raw = ""

            price = format_price(price_raw) if price_raw else ""
            price_sample = format_price(price_sample_raw) if price_sample_raw else ""

            if sku:
                entries.append((sku, price, price_sample))
        wb.close()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load entries:\n{e}")
    return entries


def save_entries_to_xlsx(entries):
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Entries"

        header_font = Font(bold=True, size=11)
        header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

        headers = ["Модел SKU", "Цена", "Цена на мостра"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center")

        for row_idx, (sku, price, price_sample) in enumerate(entries, 2):
            price_display = f"{price} €" if price else "—"
            price_sample_display = f"{price_sample} €" if price_sample else "—"

            ws.cell(row=row_idx, column=1, value=sku).border = thin_border
            ws.cell(row=row_idx, column=2, value=price_display).border = thin_border
            ws.cell(row=row_idx, column=3, value=price_sample_display).border = thin_border

        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 15
        ws.column_dimensions["C"].width = 20

        wb.save(XLSX_FILE)
        wb.close()
        return True
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save entries:\n{e}")
        return False


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Table Generator")
        self.geometry("750x550")
        self.minsize(650, 450)
        self.configure(bg="#f0f0f0")

        self.entries = load_entries_from_xlsx()
        self._build_ui()
        self._refresh_suggestions()

    def _build_ui(self):
        style = ttk.Style()
        style.configure("TButton", font=("Segoe UI", 10), padding=4)
        style.configure("TLabel", font=("Segoe UI", 10))

        top_frame = tk.Frame(self, bg="#f0f0f0", pady=10, padx=10)
        top_frame.pack(fill=tk.X)

        tk.Label(top_frame, text="Модел SKU", bg="#f0f0f0", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.autocomplete = AutocompleteEntry(top_frame, font=("Segoe UI", 10), width=35)
        self.autocomplete.grid(row=1, column=0, sticky="ew", padx=(0, 10))
        self.autocomplete.entry.bind("<Return>", lambda e: self._focus_price())

        tk.Label(top_frame, text="Цена", bg="#f0f0f0", font=("Segoe UI", 10, "bold")).grid(row=0, column=1, sticky="w")
        self.ent_price = tk.Entry(top_frame, font=("Segoe UI", 10), width=12)
        self.ent_price.grid(row=1, column=1, sticky="ew", padx=(0, 10))
        self.ent_price.bind("<Return>", lambda e: self._focus_price_sample())

        tk.Label(top_frame, text="Цена на мостра", bg="#f0f0f0", font=("Segoe UI", 10, "bold")).grid(row=0, column=2, sticky="w")
        self.ent_price_sample = tk.Entry(top_frame, font=("Segoe UI", 10), width=12)
        self.ent_price_sample.grid(row=1, column=2, sticky="ew", padx=(0, 10))
        self.ent_price_sample.bind("<Return>", lambda e: self._add_entry())

        btn_frame_top = tk.Frame(top_frame, bg="#f0f0f0")
        btn_frame_top.grid(row=2, column=0, columnspan=3, pady=(8, 0), sticky="w")

        self.btn_add = tk.Button(btn_frame_top, text="Add Entry", font=("Segoe UI", 10, "bold"),
                                  bg="#4CAF50", fg="white", padx=12, pady=2, command=self._add_entry)
        self.btn_add.pack(side=tk.LEFT, padx=(0, 5))

        sep = ttk.Separator(self, orient="horizontal")
        sep.pack(fill=tk.X, padx=10)

        list_frame = tk.Frame(self, bg="#f0f0f0")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        header_frame = tk.Frame(list_frame, bg="#D9E1F2")
        header_frame.pack(fill=tk.X)

        tk.Label(header_frame, text="Модел SKU", bg="#D9E1F2", font=("Segoe UI", 10, "bold"),
                 anchor="w", padx=5, pady=4).pack(side=tk.LEFT, expand=True, fill=tk.X)
        tk.Label(header_frame, text="Цена", bg="#D9E1F2", font=("Segoe UI", 10, "bold"),
                 anchor="w", padx=5, pady=4, width=12).pack(side=tk.LEFT)
        tk.Label(header_frame, text="Цена на мостра", bg="#D9E1F2", font=("Segoe UI", 10, "bold"),
                 anchor="w", padx=5, pady=4, width=14).pack(side=tk.LEFT)
        tk.Label(header_frame, text="", bg="#D9E1F2", width=12).pack(side=tk.LEFT)

        canvas_frame = tk.Frame(list_frame, bg="#f0f0f0")
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(canvas_frame, bg="white", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg="white")

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)

        bottom_frame = tk.Frame(self, bg="#f0f0f0", pady=8, padx=10)
        bottom_frame.pack(fill=tk.X)

        self.btn_finish = tk.Button(bottom_frame, text="Finish", font=("Segoe UI", 10, "bold"),
                                     bg="#1565C0", fg="white", padx=16, pady=2, command=self._finish)
        self.btn_finish.pack(side=tk.RIGHT)

        self._render_entries()

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event):
        if event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _focus_price(self):
        self.ent_price.focus_set()

    def _focus_price_sample(self):
        self.ent_price_sample.focus_set()

    def _refresh_suggestions(self):
        names = list({e[0] for e in self.entries})
        names.sort()
        self.autocomplete.set_suggestions(names)

    def _add_entry(self):
        sku = self.autocomplete.get().strip()
        price_raw = self.ent_price.get().strip()
        price_sample_raw = self.ent_price_sample.get().strip()

        if not sku:
            messagebox.showwarning("Warning", "Please enter a Модел SKU.")
            return

        if not is_valid_price(price_raw):
            messagebox.showwarning("Warning", "Цена must be a valid positive number.")
            return

        if price_sample_raw and not is_valid_price(price_sample_raw):
            messagebox.showwarning("Warning", "Цена на мостра must be a valid positive number if provided.")
            return

        price = format_price(price_raw)
        price_sample = format_price(price_sample_raw) if price_sample_raw else ""

        new_entry = (sku, price, price_sample)
        if new_entry in self.entries:
            messagebox.showwarning("Duplicate", "This entry already exists.")
            return

        self.entries.append(new_entry)

        self.autocomplete.delete(0, tk.END)
        self.ent_price.delete(0, tk.END)
        self.ent_price_sample.delete(0, tk.END)
        self.autocomplete.focus_set()

        self._refresh_suggestions()
        self._render_entries()
        save_entries_to_xlsx(self.entries)

    def _finish(self):
        if not self.entries:
            messagebox.showwarning("Warning", "No entries to save.")
            return
        if save_entries_to_xlsx(self.entries):
            messagebox.showinfo("Success", f"Saved {len(self.entries)} entries to:\n{XLSX_FILE}")
            self.destroy()

    def _auto_save(self):
        save_entries_to_xlsx(self.entries)

    def _render_entries(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        if not self.entries:
            tk.Label(self.scrollable_frame, text="No entries yet.", bg="white",
                      font=("Segoe UI", 10, "italic"), fg="gray", pady=20).pack()
            return

        for i, (sku, price, price_sample) in enumerate(self.entries):
            row = EntryRow(self.scrollable_frame, sku, price, price_sample,
                           delete_cb=self._delete_entry, save_cb=self._auto_save,
                           entries_ref=self.entries, original_index=i,
                           bg="white" if i % 2 == 0 else "#f9f9f9")
            row.pack(fill=tk.X)

    def _delete_entry(self, row_widget):
        data = row_widget.get_data()
        self.entries = [e for e in self.entries if not (e[0] == data[0] and e[1] == data[1] and e[2] == data[2])]
        self._refresh_suggestions()
        self._render_entries()
        save_entries_to_xlsx(self.entries)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
