from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from input_scanner import FilterResults, WorkbookStructure, inspect_workbook, scan_filters


APP_TITLE = "Generic SF Weekly Report"
WINDOW_SIZE = "980x720"
SUPPORTED_FILE_TYPES = [
    ("Excel workbooks", "*.xlsx *.xlsm"),
    ("Excel workbook", "*.xlsx"),
    ("Excel macro-enabled workbook", "*.xlsm"),
]


class GSFWeeklyApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(WINDOW_SIZE)
        self.minsize(860, 620)

        self.structure: WorkbookStructure | None = None
        self.results: FilterResults | None = None
        self.organization_values: list[str] = []
        self.selected_organizations: set[str] = set()

        self.file_var = tk.StringVar()
        self.park_var = tk.StringVar()
        self.department_var = tk.StringVar()
        self.search_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Select a Weekly Training workbook to begin.")
        self.count_var = tk.StringVar(value="Unique employees selected: 0")
        self.org_count_var = tk.StringVar(value="Organizations selected: 0")

        self._configure_style()
        self._build_ui()
        self.search_var.trace_add("write", self._on_search_changed)

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        available = style.theme_names()
        if "vista" in available:
            style.theme_use("vista")

        style.configure("Title.TLabel", font=("Segoe UI", 20, "bold"))
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10))
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
        style.configure("Count.TLabel", font=("Segoe UI", 10, "bold"))
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 8))

    def _build_ui(self) -> None:
        container = ttk.Frame(self, padding=18)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text=APP_TITLE, style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            container,
            text="Select a Weekly Training report, then filter by Park, Department, and Organization / Venue.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(2, 16))

        file_frame = ttk.LabelFrame(
            container,
            text="1. Weekly Training Report",
            padding=12,
            style="Section.TLabelframe",
        )
        file_frame.pack(fill="x")
        file_frame.columnconfigure(0, weight=1)

        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_var, state="readonly")
        self.file_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.browse_button = ttk.Button(file_frame, text="Browse...", command=self._browse_file)
        self.browse_button.grid(row=0, column=1, padx=(0, 8))
        self.load_button = ttk.Button(file_frame, text="Load Report", command=self._start_scan)
        self.load_button.grid(row=0, column=2)

        filter_frame = ttk.LabelFrame(
            container,
            text="2. Filters",
            padding=12,
            style="Section.TLabelframe",
        )
        filter_frame.pack(fill="x", pady=(14, 0))
        filter_frame.columnconfigure(0, weight=1)
        filter_frame.columnconfigure(1, weight=1)

        ttk.Label(filter_frame, text="Park").grid(row=0, column=0, sticky="w")
        ttk.Label(filter_frame, text="Department").grid(row=0, column=1, sticky="w", padx=(12, 0))

        self.park_combo = ttk.Combobox(
            filter_frame,
            textvariable=self.park_var,
            state="disabled",
        )
        self.park_combo.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self.park_combo.bind("<<ComboboxSelected>>", self._on_park_selected)

        self.department_combo = ttk.Combobox(
            filter_frame,
            textvariable=self.department_var,
            state="disabled",
        )
        self.department_combo.grid(row=1, column=1, sticky="ew", padx=(12, 0), pady=(4, 0))
        self.department_combo.bind("<<ComboboxSelected>>", self._on_department_selected)

        org_frame = ttk.LabelFrame(
            container,
            text="3. Organization / Venue",
            padding=12,
            style="Section.TLabelframe",
        )
        org_frame.pack(fill="both", expand=True, pady=(14, 0))
        org_frame.columnconfigure(0, weight=1)
        org_frame.rowconfigure(2, weight=1)

        search_row = ttk.Frame(org_frame)
        search_row.grid(row=0, column=0, columnspan=2, sticky="ew")
        search_row.columnconfigure(0, weight=1)

        self.search_entry = ttk.Entry(
            search_row,
            textvariable=self.search_var,
            state="disabled",
        )
        self.search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Label(search_row, text="Search organizations").grid(row=1, column=0, sticky="w")

        self.select_all_button = ttk.Button(
            search_row,
            text="Select All Displayed",
            command=self._select_all_displayed,
            state="disabled",
        )
        self.select_all_button.grid(row=0, column=1, rowspan=2, padx=(0, 8))

        self.clear_button = ttk.Button(
            search_row,
            text="Clear All",
            command=self._clear_all,
            state="disabled",
        )
        self.clear_button.grid(row=0, column=2, rowspan=2)

        list_frame = ttk.Frame(org_frame)
        list_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(10, 8))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.org_listbox = tk.Listbox(
            list_frame,
            selectmode=tk.MULTIPLE,
            exportselection=False,
            font=("Segoe UI", 10),
            activestyle="none",
            borderwidth=1,
            relief="solid",
        )
        self.org_listbox.grid(row=0, column=0, sticky="nsew")
        self.org_listbox.bind("<<ListboxSelect>>", self._on_listbox_selection)

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.org_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.org_listbox.configure(yscrollcommand=scrollbar.set)

        counts = ttk.Frame(org_frame)
        counts.grid(row=3, column=0, columnspan=2, sticky="ew")
        ttk.Label(counts, textvariable=self.org_count_var, style="Count.TLabel").pack(side="left")
        ttk.Label(counts, textvariable=self.count_var, style="Count.TLabel").pack(side="right")

        footer = ttk.Frame(container)
        footer.pack(fill="x", pady=(14, 0))
        footer.columnconfigure(0, weight=1)

        ttk.Label(footer, textvariable=self.status_var).grid(row=0, column=0, sticky="w")
        self.generate_button = ttk.Button(
            footer,
            text="Generate Report",
            style="Primary.TButton",
            command=self._show_selection_summary,
            state="disabled",
        )
        self.generate_button.grid(row=0, column=1, sticky="e")

        progress_frame = ttk.Frame(container)
        progress_frame.pack(fill="x", pady=(8, 0))
        self.progress = ttk.Progressbar(progress_frame, mode="indeterminate")
        self.progress.pack(fill="x")

    def _browse_file(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select Weekly Training Report",
            filetypes=SUPPORTED_FILE_TYPES,
        )
        if selected:
            self.file_var.set(selected)
            self._start_scan()

    def _start_scan(self) -> None:
        file_text = self.file_var.get().strip()
        if not file_text:
            messagebox.showwarning(APP_TITLE, "Select a Weekly Training workbook first.")
            return

        file_path = Path(file_text)
        if not file_path.exists():
            messagebox.showerror(APP_TITLE, "The selected workbook could not be found.")
            return

        self._set_busy(True)
        self.status_var.set("Reading Weekly Training report...")
        thread = threading.Thread(target=self._scan_worker, args=(file_path,), daemon=True)
        thread.start()

    def _scan_worker(self, file_path: Path) -> None:
        try:
            structure = inspect_workbook(file_path)
            results = scan_filters(structure)
        except Exception as exc:
            self.after(0, self._scan_failed, str(exc))
            return

        self.after(0, self._scan_complete, structure, results)

    def _scan_failed(self, error_text: str) -> None:
        self._set_busy(False)
        self.status_var.set("Unable to load the Weekly Training report.")
        messagebox.showerror(APP_TITLE, error_text)

    def _scan_complete(self, structure: WorkbookStructure, results: FilterResults) -> None:
        self.structure = structure
        self.results = results
        self.selected_organizations.clear()
        self.organization_values = []

        parks = sorted(results.parks)
        self.park_combo.configure(values=parks, state="readonly" if parks else "disabled")
        self.department_combo.configure(values=[], state="disabled")
        self.search_entry.configure(state="disabled")
        self.select_all_button.configure(state="disabled")
        self.clear_button.configure(state="disabled")
        self.generate_button.configure(state="disabled")
        self.park_var.set("")
        self.department_var.set("")
        self.search_var.set("")
        self._refresh_org_listbox()
        self._update_counts()
        self._set_busy(False)

        self.status_var.set(
            f"Loaded {structure.file_path.name} | {results.total_data_rows:,} training rows | "
            f"{len(parks)} parks"
        )

    def _set_busy(self, busy: bool) -> None:
        if busy:
            self.progress.start(10)
            self.browse_button.configure(state="disabled")
            self.load_button.configure(state="disabled")
        else:
            self.progress.stop()
            self.browse_button.configure(state="normal")
            self.load_button.configure(state="normal")

    def _on_park_selected(self, _event=None) -> None:
        if not self.results:
            return

        park = self.park_var.get()
        departments = sorted(self.results.departments.get(park, set()))
        self.department_var.set("")
        self.department_combo.configure(
            values=departments,
            state="readonly" if departments else "disabled",
        )
        self.organization_values = []
        self.selected_organizations.clear()
        self.search_var.set("")
        self.search_entry.configure(state="disabled")
        self.select_all_button.configure(state="disabled")
        self.clear_button.configure(state="disabled")
        self.generate_button.configure(state="disabled")
        self._refresh_org_listbox()
        self._update_counts()
        self.status_var.set(f"Park selected: {park}")

    def _on_department_selected(self, _event=None) -> None:
        if not self.results:
            return

        park = self.park_var.get()
        department = self.department_var.get()
        self.organization_values = sorted(
            self.results.organizations.get((park, department), set())
        )
        self.selected_organizations.clear()
        self.search_var.set("")

        enabled_state = "normal" if self.organization_values else "disabled"
        self.search_entry.configure(state=enabled_state)
        self.select_all_button.configure(state=enabled_state)
        self.clear_button.configure(state=enabled_state)
        self.generate_button.configure(state="disabled")
        self._refresh_org_listbox()
        self._update_counts()
        self.status_var.set(
            f"{len(self.organization_values)} organizations available for {department}."
        )

    def _on_search_changed(self, *_args) -> None:
        self._capture_visible_selection()
        self._refresh_org_listbox()

    def _visible_organizations(self) -> list[str]:
        query = self.search_var.get().strip().casefold()
        if not query:
            return list(self.organization_values)
        return [value for value in self.organization_values if query in value.casefold()]

    def _refresh_org_listbox(self) -> None:
        self.org_listbox.delete(0, tk.END)
        visible = self._visible_organizations()

        for organization in visible:
            self.org_listbox.insert(tk.END, organization)

        for index, organization in enumerate(visible):
            if organization in self.selected_organizations:
                self.org_listbox.selection_set(index)

    def _capture_visible_selection(self) -> None:
        visible = self._visible_organizations()
        selected_indexes = set(self.org_listbox.curselection())

        for index, organization in enumerate(visible):
            if index in selected_indexes:
                self.selected_organizations.add(organization)
            else:
                self.selected_organizations.discard(organization)

    def _on_listbox_selection(self, _event=None) -> None:
        self._capture_visible_selection()
        self._update_counts()

    def _select_all_displayed(self) -> None:
        visible = self._visible_organizations()
        self.selected_organizations.update(visible)
        self.org_listbox.selection_set(0, tk.END)
        self._update_counts()

    def _clear_all(self) -> None:
        self.selected_organizations.clear()
        self.org_listbox.selection_clear(0, tk.END)
        self._update_counts()

    def _selected_employee_ids(self) -> set[str]:
        if not self.results:
            return set()

        park = self.park_var.get()
        department = self.department_var.get()
        employees: set[str] = set()

        for organization in self.selected_organizations:
            employees.update(
                self.results.employees_by_organization.get(
                    (park, department, organization), set()
                )
            )

        return employees

    def _update_counts(self) -> None:
        employee_count = len(self._selected_employee_ids())
        organization_count = len(self.selected_organizations)
        self.org_count_var.set(f"Organizations selected: {organization_count}")
        self.count_var.set(f"Unique employees selected: {employee_count}")
        self.generate_button.configure(
            state="normal" if organization_count else "disabled"
        )

    def _show_selection_summary(self) -> None:
        self._capture_visible_selection()
        self._update_counts()

        if not self.selected_organizations:
            messagebox.showwarning(APP_TITLE, "Select at least one organization.")
            return

        organizations = "\n".join(
            f"- {value}" for value in sorted(self.selected_organizations)
        )
        employee_count = len(self._selected_employee_ids())

        messagebox.showinfo(
            APP_TITLE,
            "Selections are ready for report generation.\n\n"
            f"Park: {self.park_var.get()}\n"
            f"Department: {self.department_var.get()}\n"
            f"Unique employees: {employee_count}\n\n"
            f"Organizations:\n{organizations}\n\n"
            "The Excel report generator is the next project checkpoint.",
        )


def main() -> None:
    app = GSFWeeklyApp()
    app.mainloop()


if __name__ == "__main__":
    main()
