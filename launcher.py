from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from roof_automation import generate_bid_form, generate_rfp, level_bids


class RoofTool(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Maverick Roof RFP & Bid Leveling")
        self.geometry("760x390")
        self.minsize(680, 360)
        self.intake = tk.StringVar()
        self.pdf = tk.BooleanVar(value=True)
        self.status = tk.StringVar(value="Select a project intake workbook to begin.")
        self._build()

    def _build(self):
        style = ttk.Style(self)
        style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"), foreground="#245B63")
        style.configure("Step.TLabel", font=("Segoe UI", 11, "bold"))
        frame = ttk.Frame(self, padding=24)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Maverick Roof Automation", style="Title.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
        ttk.Label(frame, text="Generate a locked roof RFP, issue the standard contractor bid form, and level returned Excel bids.", wraplength=690).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 22))

        ttk.Label(frame, text="1. Project intake", style="Step.TLabel").grid(row=2, column=0, columnspan=3, sticky="w")
        ttk.Entry(frame, textvariable=self.intake, width=78).grid(row=3, column=0, columnspan=2, sticky="ew", pady=7)
        ttk.Button(frame, text="Browse...", command=self.choose_intake).grid(row=3, column=2, padx=(8, 0))

        ttk.Label(frame, text="2. Prepare procurement package", style="Step.TLabel").grid(row=4, column=0, columnspan=3, sticky="w", pady=(18, 4))
        ttk.Button(frame, text="Generate RFP + Contractor Bid Form", command=self.prepare).grid(row=5, column=0, sticky="w")
        ttk.Checkbutton(frame, text="Also create PDF when Microsoft Word is available", variable=self.pdf).grid(row=5, column=1, columnspan=2, sticky="w", padx=(14, 0))

        ttk.Label(frame, text="3. Level returned contractor bids", style="Step.TLabel").grid(row=6, column=0, columnspan=3, sticky="w", pady=(18, 4))
        ttk.Button(frame, text="Select Excel Bids + Create Leveling", command=self.level).grid(row=7, column=0, columnspan=2, sticky="w")

        ttk.Separator(frame).grid(row=8, column=0, columnspan=3, sticky="ew", pady=(24, 12))
        ttk.Label(frame, textvariable=self.status, wraplength=690).grid(row=9, column=0, columnspan=3, sticky="w")
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)

    def choose_intake(self):
        path = filedialog.askopenfilename(title="Select Roof Project Intake", filetypes=[("Excel workbooks", "*.xlsx")])
        if path:
            self.intake.set(path)
            self.status.set(f"Selected: {Path(path).name}")

    def require_intake(self) -> Path | None:
        path = Path(self.intake.get()).expanduser()
        if not path.exists():
            messagebox.showerror("Missing intake", "Select a valid Roof_Project_Intake.xlsx file first.")
            return None
        return path

    def prepare(self):
        intake = self.require_intake()
        if not intake:
            return
        directory = filedialog.askdirectory(title="Choose output folder")
        if not directory:
            return
        try:
            output_dir = Path(directory)
            rfp = generate_rfp(intake, output_dir, self.pdf.get())
            bid_form = generate_bid_form(intake, output_dir)
            self.status.set(f"Created {rfp.name} and {bid_form.name}")
            messagebox.showinfo("Package created", f"Created:\n\n{rfp}\n{bid_form}")
        except Exception as exc:
            self.status.set("Generation stopped. Review the message and correct the intake workbook.")
            messagebox.showerror("Could not generate package", str(exc))

    def level(self):
        intake = self.require_intake()
        if not intake:
            return
        bid_files = filedialog.askopenfilenames(title="Select 1-4 completed contractor bid workbooks", filetypes=[("Excel workbooks", "*.xlsx")])
        if not bid_files:
            return
        output = filedialog.asksaveasfilename(title="Save bid leveling workbook", defaultextension=".xlsx", initialfile="Roof_Bid_Leveling.xlsx", filetypes=[("Excel workbook", "*.xlsx")])
        if not output:
            return
        try:
            result = level_bids(intake, [Path(p) for p in bid_files], Path(output))
            self.status.set(f"Created {result.name}")
            messagebox.showinfo("Bid leveling created", f"Created:\n\n{result}")
        except Exception as exc:
            self.status.set("Bid import stopped. Review the message and contractor files.")
            messagebox.showerror("Could not level bids", str(exc))


if __name__ == "__main__":
    RoofTool().mainloop()
