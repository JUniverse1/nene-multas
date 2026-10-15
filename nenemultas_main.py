"""Interfaz de escritorio para buscar registros en el CSV de fotodetección."""

import csv
import logging
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from lxml.etree import XMLSyntaxError

from compafechas import comparar_fechas, diferencia_en_anos
from nenemul_brain import (
    ADDRESS_COLUMN,
    DEFAULT_CSV,
    find_matches,
    read_csv,
)
from xls_to_csv_transformer import convert_excel_to_csv

logger = logging.getLogger(__name__)


class AddressSearchApp:
    THEMES = {
        "Claro": {
            "background": "#f3f5f7",
            "surface": "#ffffff",
            "foreground": "#17202a",
            "muted": "#586575",
            "field": "#ffffff",
            "border": "#cbd2d9",
            "heading": "#e7ecf1",
            "selected": "#c9e2ff",
            "button_active": "#dbe5ef",
        },
        "Oscuro": {
            "background": "#070c13",
            "surface": "#2b323a",
            "foreground": "#f2f8ff",
            "muted": "#b4bec9",
            "field": "#343d47",
            "border": "#48535f",
            "heading": "#3a444f",
            "selected": "#315b80",
            "button_active": "#465361",
        },
        "Vaporwave": {
            "background": "#f4eff8",
            "surface": "#fbf8fc",
            "foreground": "#3f3548",
            "muted": "#6d6176",
            "field": "#f8f4fb",
            "border": "#d9cfe2",
            "heading": "#eee7f3",
            "selected": "#ded2ea",
            "button_active": "#e8deef",
        },
        "Solarizado": {
            "background": "#f6f3e8",
            "surface": "#fbf9f2",
            "foreground": "#46575a",
            "muted": "#5e7073",
            "field": "#fdfbf5",
            "border": "#d7d2c2",
            "heading": "#eee9d9",
            "selected": "#d9e5e3",
            "button_active": "#e4e9dc",
        },
        "Deep ocean": {
            "background": "#0b1e2d",
            "surface": "#1a3a5f",
            "foreground": "#a9cfe0",
            "muted": "#5a7b9d",
            "field": "#1a3a5f",
            "border": "#2c4a7a",
            "heading": "#1a3a5f",
            "selected": "#3a6aa5",
            "button_active": "#2c4a7a",
        }
    }

    def __init__(self, root):
        self.root = root
        self.root.title("Búsqueda de fotodetección")
        self.root.geometry("1450x780")
        self.style = ttk.Style(root)
        self.style.theme_use("clam")
        self.theme_var = tk.StringVar(value="Claro")
        self.csv_path = DEFAULT_CSV
        self.records = []
        self.fields = []
        self.last_matches = []
        self.criteria_rows = []
        self.multa_count = 0
        self.page_count = 0

        self._apply_theme()
        self._build_ui()
        self._load_csv(self.csv_path)

    def _apply_theme(self, *_):
        colors = self.THEMES[self.theme_var.get()]
        self.root.configure(background=colors["background"])
        self.style.configure(
            ".",
            background=colors["background"],
            foreground=colors["foreground"],
            font=("Segoe UI", 10),
        )
        self.style.configure(
            "TFrame",
            background=colors["background"],
        )
        self.style.configure(
            "TLabel",
            background=colors["background"],
            foreground=colors["foreground"],
        )
        self.style.configure(
            "Title.TLabel",
            background=colors["background"],
            foreground=colors["foreground"],
            font=("Segoe UI", 16, "bold"),
        )
        self.style.configure(
            "TLabelframe",
            background=colors["background"],
            foreground=colors["foreground"],
            bordercolor=colors["border"],
        )
        self.style.configure(
            "TLabelframe.Label",
            background=colors["background"],
            foreground=colors["foreground"],
        )
        self.style.configure(
            "TButton",
            background=colors["surface"],
            foreground=colors["foreground"],
            bordercolor=colors["border"],
            padding=(10, 5),
        )
        self.style.map(
            "TButton",
            background=[
                ("active", colors["button_active"]),
                ("pressed", colors["selected"]),
            ],
            foreground=[("disabled", colors["muted"])],
        )
        self.style.configure(
            "TEntry",
            fieldbackground=colors["field"],
            foreground=colors["foreground"],
            insertcolor=colors["foreground"],
            bordercolor=colors["border"],
        )
        self.style.configure(
            "TCombobox",
            fieldbackground=colors["field"],
            background=colors["surface"],
            foreground=colors["foreground"],
            arrowcolor=colors["foreground"],
            bordercolor=colors["border"],
        )
        self.style.map(
            "TCombobox",
            fieldbackground=[("readonly", colors["field"])],
            foreground=[("readonly", colors["foreground"])],
            selectbackground=[("readonly", colors["field"])],
            selectforeground=[("readonly", colors["foreground"])],
        )
        self.style.configure(
            "Treeview",
            background=colors["field"],
            fieldbackground=colors["field"],
            foreground=colors["foreground"],
            bordercolor=colors["border"],
            rowheight=26,
        )
        self.style.map(
            "Treeview",
            background=[("selected", colors["selected"])],
            foreground=[("selected", colors["foreground"])],
        )
        self.style.configure(
            "Treeview.Heading",
            background=colors["heading"],
            foreground=colors["foreground"],
            bordercolor=colors["border"],
            padding=(7, 5),
            font=("Segoe UI", 9, "bold"),
        )
        self.style.map(
            "Treeview.Heading",
            background=[("active", colors["button_active"])],
        )
        self.style.configure(
            "TScrollbar",
            background=colors["heading"],
            troughcolor=colors["background"],
            bordercolor=colors["border"],
            arrowcolor=colors["foreground"],
        )
        if hasattr(self, "table"):
            self.table.tag_configure(
                "even", background=colors["field"], foreground=colors["foreground"]
            )
            self.table.tag_configure(
                "odd", background=colors["surface"], foreground=colors["foreground"]
            )

    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        header = ttk.Frame(self.root, padding=(12, 12, 12, 6))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)
        ttk.Label(
            header,
            text="Buscar registros",
            style="Title.TLabel",
        ).grid(row=0, column=0, sticky="w")
        self.csv_label = ttk.Label(header, text="")
        self.csv_label.grid(row=0, column=1, sticky="e", padx=10)
        ttk.Label(header, text="Apariencia:").grid(
            row=0, column=2, sticky="e", padx=(4, 6)
        )
        theme_box = ttk.Combobox(
            header,
            textvariable=self.theme_var,
            values=tuple(self.THEMES),
            state="readonly",
            width=10,
        )
        theme_box.grid(row=0, column=3, sticky="e", padx=(0, 10))
        theme_box.bind("<<ComboboxSelected>>", self._apply_theme)
        ttk.Button(
            header, text="Abrir CSV...", command=self._choose_csv
        ).grid(row=0, column=4, sticky="e")
        ttk.Button(
            header, text="Convertir XLS...", command=self._convert_xls
        ).grid(row=0, column=5, sticky="e", padx=(8, 0))

        criteria_panel = ttk.LabelFrame(
            self.root,
            text="Criterios de búsqueda (se deben cumplir todos)",
            padding=10,
        )
        criteria_panel.grid(row=1, column=0, sticky="ew", padx=12, pady=6)
        criteria_panel.columnconfigure(0, weight=1)
        self.criteria_frame = ttk.Frame(criteria_panel)
        self.criteria_frame.grid(row=0, column=0, sticky="ew")
        self.criteria_frame.columnconfigure(1, weight=1)

        controls = ttk.Frame(criteria_panel)
        controls.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(
            controls, text="＋ Agregar criterio", command=self._add_criterion
        ).pack(side="left")
        ttk.Button(
            controls, text="Buscar", command=self._search
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            controls, text="Comparar fechas", command=self._open_date_comparison
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            controls, text="Limpiar criterios", command=self._clear_criteria
        ).pack(side="left", padx=(8, 0))
        self.export_button = ttk.Button(
            controls, text="Exportar resultados...",
            command=self._export_results, state="disabled",
        )
        self.export_button.pack(side="left", padx=(8, 0))
        counter_controls = ttk.Frame(controls)
        counter_controls.pack(side="right", padx=(12, 0))
        self.multa_counter_button = ttk.Button(
            counter_controls,
            text="Multa #: 0",
            command=self._increment_multa_count,
        )
        self.multa_counter_button.pack(side="left")
        self.page_counter_button = ttk.Button(
            counter_controls,
            text="Página: 0",
            command=self._increment_page_count,
        )
        self.page_counter_button.pack(side="left", padx=(6, 0))
        ttk.Button(
            counter_controls,
            text="Reiniciar conteo",
            command=self._reset_counts,
        ).pack(side="left", padx=(6, 0))
        self.status_label = ttk.Label(controls, text="")
        self.status_label.pack(side="right")

        table_frame = ttk.Frame(self.root, padding=(12, 4, 12, 12))
        table_frame.grid(row=2, column=0, sticky="nsew")
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        self.table = ttk.Treeview(table_frame, show="headings")
        vertical_scroll = ttk.Scrollbar(
            table_frame, orient="vertical", command=self.table.yview
        )
        horizontal_scroll = ttk.Scrollbar(
            table_frame, orient="horizontal", command=self.table.xview
        )
        self.table.configure(
            yscrollcommand=vertical_scroll.set,
            xscrollcommand=horizontal_scroll.set,
        )
        self.table.grid(row=0, column=0, sticky="nsew")
        vertical_scroll.grid(row=0, column=1, sticky="ns")
        horizontal_scroll.grid(row=1, column=0, sticky="ew")

    def _increment_multa_count(self):
        self.multa_count += 1
        self.multa_counter_button.configure(text=f"Multa #: {self.multa_count}")

    def _increment_page_count(self):
        self.page_count += 1
        self.page_counter_button.configure(text=f"Página: {self.page_count}")

    def _reset_counts(self):
        self.multa_count = 0
        self.page_count = 0
        self.multa_counter_button.configure(text="Multa #: 0")
        self.page_counter_button.configure(text="Página: 0")

    def _load_csv(self, path):
        try:
            fields, records = read_csv(path)
        except (OSError, UnicodeError, csv.Error, ValueError) as error:
            logger.exception("No se pudo cargar %s", path)
            messagebox.showerror(
                "No se pudo abrir el CSV",
                f"No se pudo cargar el archivo:\n{error}",
                parent=self.root,
            )
            return

        self.csv_path = Path(path)
        self.records = records
        self.fields = fields
        self.last_matches = []
        self.export_button.configure(state="disabled")
        self.csv_label.configure(text=self.csv_path.name)
        self._clear_criteria()
        self._configure_table()
        self.status_label.configure(
            text=f"{len(records):,} registros cargados"
        )

    def _configure_table(self):
        self.table.delete(*self.table.get_children())
        columns = ("_coincidencia", *self.fields)
        self.table.configure(columns=columns)
        self.table.heading("_coincidencia", text="Coincidencia")
        self.table.column(
            "_coincidencia", width=105, minwidth=105, stretch=False,
            anchor="center",
        )
        for field in self.fields:
            self.table.heading(field, text=field)
            self.table.column(
                field,
                width=310 if field == ADDRESS_COLUMN else 155,
                minwidth=100,
                stretch=False,
            )

    def _choose_csv(self):
        selected = filedialog.askopenfilename(
            title="Seleccionar archivo CSV",
            filetypes=(("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")),
            initialdir=str(self.csv_path.parent),
        )
        if selected:
            self._load_csv(Path(selected))

    def _convert_xls(self):
        selected = filedialog.askopenfilename(
            title="Seleccionar archivo XLSX",
            filetypes=(("Archivos XLSX", "*.xlsx"), ("Todos los archivos", "*.*")),
            initialdir=str(self.csv_path.parent),
        )
        if not selected:
            return

        input_path = Path(selected)
        output = filedialog.asksaveasfilename(
            title="Guardar CSV convertido",
            initialdir=str(input_path.parent),
            initialfile=input_path.with_suffix(".csv").name,
            defaultextension=".csv",
            filetypes=(("Archivos CSV", "*.csv"),),
        )
        if not output:
            return

        try:
            output_path = convert_excel_to_csv(input_path, output)
        except (OSError, ImportError, ValueError) as error:
            logger.exception("No se pudo convertir %s", input_path)
            messagebox.showerror(
                "No se pudo convertir el archivo Excel",
                f"No se pudo convertir el archivo:\n{error}",
                parent=self.root,
            )
            return

        if messagebox.askyesno(
            "Conversión completa",
            f"Se guardó el archivo:\n{output_path}\n\n"
            "¿Quieres cargarlo en la interfaz?",
            parent=self.root,
        ):
            self._load_csv(output_path)

    def _add_criterion(self, field=ADDRESS_COLUMN):
        if not self.fields:
            return
        row = ttk.Frame(self.criteria_frame)
        row.pack(fill="x", pady=2)
        field_var = tk.StringVar(value=field if field in self.fields else self.fields[0])
        field_box = ttk.Combobox(
            row, textvariable=field_var, values=self.fields,
            state="readonly", width=32,
        )
        field_box.grid(row=0, column=0, sticky="w", padx=(0, 8))
        query_entry = ttk.Entry(row)
        query_entry.grid(row=0, column=1, sticky="ew")
        remove_button = ttk.Button(
            row, text="Quitar", command=lambda: self._remove_criterion(row)
        )
        remove_button.grid(row=0, column=2, padx=(8, 0))
        self.criteria_rows.append((row, field_var, query_entry))
        query_entry.focus_set()

    def _remove_criterion(self, row):
        self.criteria_rows = [
            criterion for criterion in self.criteria_rows
            if criterion[0] is not row
        ]
        row.destroy()

    def _clear_criteria(self):
        for row, _, _ in self.criteria_rows:
            row.destroy()
        self.criteria_rows.clear()
        if self.fields:
            self._add_criterion()

    def _open_date_comparison(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Comparar fechas")
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.columnconfigure(1, weight=1)

        ttk.Label(
            dialog,
            text="Ingresa las dos fechas en formato dd/mm/aaaa:",
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=14, pady=(14, 10))

        ttk.Label(dialog, text="Primera fecha:").grid(
            row=1, column=0, sticky="w", padx=14, pady=4
        )
        first_date = ttk.Entry(dialog, width=18)
        first_date.grid(row=1, column=1, sticky="w", padx=(8, 14), pady=4)

        ttk.Label(dialog, text="Segunda fecha:").grid(
            row=2, column=0, sticky="w", padx=14, pady=4
        )
        second_date = ttk.Entry(dialog, width=18)
        second_date.grid(row=2, column=1, sticky="w", padx=(8, 14), pady=4)

        result_label = ttk.Label(dialog, text="", wraplength=390)
        result_label.grid(
            row=3, column=0, columnspan=2, sticky="w", padx=14, pady=(8, 4)
        )

        def compare():
            try:
                first = datetime.strptime(first_date.get().strip(), "%d/%m/%Y").date()
                second = datetime.strptime(
                    second_date.get().strip(), "%d/%m/%Y"
                ).date()
            except ValueError:
                result_label.configure(
                    text="Ingresa dos fechas válidas con el formato dd/mm/aaaa."
                )
                return

            difference = abs((first - second).days)
            years = diferencia_en_anos(first, second)
            result_label.configure(
                text=(
                    f"Válido: hay {difference} días de diferencia "
                    f"({years:.2f} años)."
                )
            )

        ttk.Button(
            dialog, text="Comparar", command=compare
        ).grid(row=4, column=0, sticky="w", padx=14, pady=(8, 14))
        ttk.Button(
            dialog, text="Cerrar", command=dialog.destroy
        ).grid(row=4, column=1, sticky="e", padx=14, pady=(8, 14))

        dialog.bind(
            "<Return>",
            lambda event: compare() if event.keysym == "Return" else None,
        )
        dialog.bind(
            "<Escape>",
            lambda event: dialog.destroy() if event.keysym == "Escape" else None,
        )
        dialog.grab_set()
        first_date.focus_set()

    def _search(self):
        criteria = []
        for _, field_var, query_entry in self.criteria_rows:
            query = query_entry.get().strip()
            if query:
                criteria.append((field_var.get(), query))

        if not criteria:
            messagebox.showwarning(
                "Falta un criterio",
                "Escribe al menos un valor para buscar.",
                parent=self.root,
            )
            return

        self.root.configure(cursor="watch")
        self.status_label.configure(text="Buscando...")
        self.root.update_idletasks()
        try:
            matches = find_matches(self.records, criteria)
        finally:
            self.root.configure(cursor="")
        logger.info("Criterios %s -> %d coincidencias", criteria, len(matches))
        self.last_matches = matches
        self.export_button.configure(state="normal" if matches else "disabled")
        self.table.delete(*self.table.get_children())
        colors = self.THEMES[self.theme_var.get()]
        self.table.tag_configure(
            "even", background=colors["field"], foreground=colors["foreground"]
        )
        self.table.tag_configure(
            "odd", background=colors["surface"], foreground=colors["foreground"]
        )
        for index, (score, _, record) in enumerate(matches):
            percentage = f"{score:.1f}%" if score else "—"
            values = (percentage, *(record.get(field, "") or "" for field in self.fields))
            self.table.insert(
                "", "end", values=values, tags=("even" if index % 2 == 0 else "odd",)
            )
        self.status_label.configure(
            text=f"{len(matches):,} coincidencias"
        )

    def _export_results(self):
        if not self.last_matches:
            return
        output = filedialog.asksaveasfilename(
            title="Exportar resultados a CSV",
            initialdir=str(self.csv_path.parent),
            initialfile=f"{self.csv_path.stem}-resultados.csv",
            defaultextension=".csv",
            filetypes=(("Archivos CSV", "*.csv"),),
        )
        if not output:
            return
        try:
            # utf-8-sig para que Excel muestre bien las tildes.
            with Path(output).open("w", encoding="utf-8-sig", newline="") as file:
                writer = csv.writer(file)
                writer.writerow(("Coincidencia", *self.fields))
                for score, _, record in self.last_matches:
                    writer.writerow((
                        f"{score:.1f}" if score else "",
                        *(record.get(field, "") or "" for field in self.fields),
                    ))
        except OSError as error:
            logger.exception("No se pudo exportar a %s", output)
            messagebox.showerror(
                "No se pudo exportar",
                f"No se pudo guardar el archivo:\n{error}",
                parent=self.root,
            )
            return
        messagebox.showinfo(
            "Exportación completa",
            f"Se guardaron {len(self.last_matches):,} filas en:\n{output}",
            parent=self.root,
        )


def main():
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s"
    )
    root = tk.Tk()
    AddressSearchApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()