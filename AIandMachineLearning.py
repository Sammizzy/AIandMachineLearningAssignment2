
from __future__ import annotations

import json
import os
import sqlite3
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from pathlib import Path
from datetime import datetime, timezone

import pandas as pd


APP_TITLE = "Mortgage Application Review Desk"


if os.name == "nt":
    _data_root = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "MortgageReviewDesk"
else:
    _data_root = Path.home() / ".local" / "share" / "MortgageReviewDesk"
_data_root.mkdir(parents=True, exist_ok=True)
DB_PATH = _data_root / "mortgage_review_decisions.sqlite3"


EXCLUDED_REVIEW_FIELDS = {
    "repayment_difficulty_12m",
    "risk_tier",
    "risk_flag_count",
    "high_risk_loan",
    "F_LTV_90plus",
    "F_DTI_over_0.45",
    "F_Score_under_600",
    "F_Util_70plus",
    "F_Prev_default",
    "F_Missed_2plus",
    "missed_payment_group",
    "has_missed_payments",
    "risk_indicator_count",
    "risk_indicator_score_10",
    "risk_indicator_reasons",
}

FIELD_LABELS = {
    "annual_income_gbp": "Annual income",
    "property_value_gbp": "Property value",
    "deposit_gbp": "Deposit",
    "requested_loan_gbp": "Requested loan",
    "loan_to_value_pct": "Loan-to-value (LTV)",
    "monthly_debt_commitments_gbp": "Monthly debt commitments",
    "debt_to_income_ratio": "Debt-to-income ratio (dataset field)",
    "credit_score": "Credit score (source field)",
    "employment_type": "Employment type",
    "employment_years": "Years in employment",
    "savings_gbp": "Savings",
    "missed_payments_24m": "Missed payments (24 months)",
    "credit_utilisation_pct": "Credit utilisation",
    "previous_default_flag": "Previous default flag (source field)",
    "term_years": "Requested term",
    "interest_rate_pct": "Illustrative interest rate in source",
    "application_channel": "Application channel",
}

CURRENCY_FIELDS = {
    "annual_income_gbp", "property_value_gbp", "deposit_gbp",
    "requested_loan_gbp", "monthly_debt_commitments_gbp", "savings_gbp"
}
PERCENT_FIELDS = {"loan_to_value_pct", "credit_utilisation_pct", "interest_rate_pct"}


def safe_value(value):
    if pd.isna(value):
        return "Not supplied"
    return value


def display_value(field, value):
    value = safe_value(value)
    if value == "Not supplied":
        return value
    try:
        if field in CURRENCY_FIELDS:
            return f"£{float(value):,.2f}"
        if field in PERCENT_FIELDS:
            return f"{float(value):,.2f}%"
        if field == "debt_to_income_ratio":
            return f"{float(value):.3f} (as supplied)"
        if field == "previous_default_flag":
            return "Yes (flagged)" if int(float(value)) == 1 else "No (not flagged)"
        if field == "term_years":
            return f"{float(value):g} years"
        if field == "employment_years":
            return f"{float(value):g} years"
        if field == "credit_score":
            return f"{float(value):g}"
        return str(value)
    except (TypeError, ValueError):
        return str(value)


class ReviewDatabase:
    def __init__(self, path: Path):
        self.conn = sqlite3.connect(path, timeout=30)
        self.conn.execute("PRAGMA busy_timeout=30000")
        self.conn.execute("PRAGMA journal_mode=WAL")
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS decisions (
                application_id TEXT PRIMARY KEY,
                decision TEXT NOT NULL,
                reviewer TEXT NOT NULL,
                reason TEXT NOT NULL,
                decided_at TEXT NOT NULL
            )
        """)
        self.conn.commit()

    def all(self):
        return {
            row[0]: {
                "decision": row[1], "reviewer": row[2],
                "reason": row[3], "decided_at": row[4]
            }
            for row in self.conn.execute(
                "SELECT application_id, decision, reviewer, reason, decided_at FROM decisions"
            )
        }

    def save(self, app_id, decision, reviewer, reason):
        stamp = datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")
        self.conn.execute("""
            INSERT INTO decisions(application_id, decision, reviewer, reason, decided_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(application_id) DO UPDATE SET
                decision=excluded.decision,
                reviewer=excluded.reviewer,
                reason=excluded.reason,
                decided_at=excluded.decided_at
        """, (app_id, decision, reviewer, reason, stamp))
        self.conn.commit()


class MortgageReviewApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1180x760")
        self.root.minsize(980, 640)
        self.db = ReviewDatabase(DB_PATH)
        self.decisions = self.db.all()
        self.raw_df = pd.DataFrame()
        self.cleaned_df = pd.DataFrame()
        self.norm_df = pd.DataFrame()
        self.workbook_sheets = {}
        self.all_segments_df = pd.DataFrame()
        self.applications = pd.DataFrame()
        self.current_id = None
        self._build_ui()
        self._refresh_counts()

    def _build_ui(self):
        header = ttk.Frame(self.root, padding=(14, 12))
        header.pack(fill="x")
        ttk.Label(header, text=APP_TITLE, font=("Segoe UI", 18, "bold")).pack(anchor="w")
        ttk.Label(
            header,
            text=("Rule-based risk indicator score (1–10) for review prioritisation only. "
                  "It is not an approval recommendation; the reviewer makes the final decision."),
            wraplength=1080
        ).pack(anchor="w", pady=(4, 8))

        import_bar = ttk.Frame(self.root, padding=(14, 0, 14, 10))
        import_bar.pack(fill="x")
        ttk.Button(import_bar, text="1. Import raw CSV", command=self.import_csv).pack(side="left")
        ttk.Button(import_bar, text="2. Import analysis workbook", command=self.import_workbook).pack(side="left", padx=8)
        ttk.Button(import_bar, text="Load both supplied files", command=self.load_both).pack(side="left")
        self.status = ttk.Label(import_bar, text="Import the CSV and workbook to begin.")
        self.status.pack(side="left", padx=14)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        self.review_tab = ttk.Frame(self.notebook, padding=10)
        self.processed_tab = ttk.Frame(self.notebook, padding=10)
        self.data_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.review_tab, text="Applications to review")
        self.notebook.add(self.processed_tab, text="Accepted / declined history")
        self.notebook.add(self.data_tab, text="Imported data & notes")

        self._build_review_tab()
        self._build_processed_tab()
        self._build_data_tab()

        footer = ttk.Frame(self.root, padding=(14, 0, 14, 10))
        footer.pack(fill="x")
        ttk.Label(
            footer,
            text=f"Decision history is stored locally in: {DB_PATH.name}",
            foreground="#555"
        ).pack(anchor="w")

    def _build_review_tab(self):
        top = ttk.Frame(self.review_tab)
        top.pack(fill="x", pady=(0, 8))
        self.pending_count = ttk.Label(top, text="Pending: 0", font=("Segoe UI", 10, "bold"))
        self.pending_count.pack(side="left")
        self.total_count = ttk.Label(top, text="Imported: 0")
        self.total_count.pack(side="left", padx=18)

        body = ttk.Panedwindow(self.review_tab, orient="horizontal")
        body.pack(fill="both", expand=True)

        left = ttk.Frame(body)
        right = ttk.Frame(body, padding=(10, 0, 0, 0))
        body.add(left, weight=1)
        body.add(right, weight=3)

        ttk.Label(left, text="Pending applications", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.app_list = ttk.Treeview(left, columns=("id", "score"), show="headings", height=20)
        self.app_list.heading("id", text="Application ID")
        self.app_list.heading("score", text="Risk /10")
        self.app_list.column("id", width=135, anchor="w")
        self.app_list.column("score", width=70, anchor="center")
        self.app_list.pack(fill="both", expand=True, pady=6)
        self.app_list.bind("<<TreeviewSelect>>", self.on_select)

        ttk.Label(right, text="Application information", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.detail_text = tk.Text(right, wrap="word", height=22, font=("Segoe UI", 10), state="disabled")
        self.detail_text.pack(fill="both", expand=True, pady=6)
        self.detail_text.tag_configure("heading", font=("Segoe UI", 10, "bold"))
        self.detail_text.tag_configure("notice", foreground="#7a3e00")

        actions = ttk.Frame(right)
        actions.pack(fill="x", pady=(4, 0))
        ttk.Button(actions, text="Accept", command=lambda: self.record_decision("Accepted")).pack(side="left")
        ttk.Button(actions, text="Decline", command=lambda: self.record_decision("Declined")).pack(side="left", padx=8)
        ttk.Button(actions, text="Clear selection", command=self.clear_selection).pack(side="left")

    def _build_processed_tab(self):
        bar = ttk.Frame(self.processed_tab)
        bar.pack(fill="x", pady=(0, 8))
        ttk.Label(bar, text="Recorded decisions", font=("Segoe UI", 11, "bold")).pack(side="left")
        ttk.Button(bar, text="Refresh", command=self.refresh_processed).pack(side="right")
        self.history = ttk.Treeview(
            self.processed_tab,
            columns=("id", "score", "decision", "reviewer", "date", "reason"),
            show="headings"
        )
        for col, title, width in [
            ("id", "Application ID", 120), ("score", "Risk /10", 75),
            ("decision", "Decision", 100), ("reviewer", "Reviewer", 130),
            ("date", "Date/time", 180), ("reason", "Reason recorded", 390)
        ]:
            self.history.heading(col, text=title)
            self.history.column(col, width=width, anchor="w")
        self.history.pack(fill="both", expand=True)

    def _build_data_tab(self):
        ttk.Label(self.data_tab, text="Import overview", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.data_info = tk.Text(self.data_tab, wrap="word", height=15, state="disabled")
        self.data_info.pack(fill="both", expand=True, pady=8)
        ttk.Button(
            self.data_tab,
            text="Run Fresh Risk Validation",
            command=self.generate_validation_report
        ).pack(anchor="w", pady=(0, 6))
        ttk.Button(
            self.data_tab,
            text="Export All Application Scores (CSV)",
            command=self.export_application_scores
        ).pack(anchor="w", pady=(0, 10))
        note = (
            "Review safeguards\n"
            "• The 1–10 risk-indicator score is calculated from six explicit thresholds: "
            "LTV >= 90%, DTI > 0.45, credit score < 600, utilisation >= 70%, previous "
            "default, and 2+ missed payments in 24 months. Each triggered indicator reduces "
            "the score by 1.5 points (minimum 1).\n"
            "• 10 means none of these six indicators were triggered; 1 means all six were "
            "triggered. This is not 'definitely accept/decline', not a probability, and not "
            "a validated underwriting score. It must not be used as an automatic decision.\n"
            "• If a required source value is missing or was imputed, the score is shown as "
            "Insufficient data rather than assuming the applicant has no risk indicator.\n"
            "• The retrospective repayment_difficulty_12m outcome is never used in an "
            "individual score.\n"
            "• Normalised columns are retained for data inspection only. Min-max/z-score "
            "values are not interpreted as credit scores or approval recommendations.\n"
            "• Employment type, application channel and other potentially sensitive or "
            "proxy variables should be handled under an approved fair-lending policy.\n"
            "• The reviewer must record an individual reason for every decision. "
            "A reason is not generated by this application.\n"
            "• The dataset appears to contain synthetic/example records. Do not use "
            "these records to make real lending decisions."
        )
        ttk.Label(self.data_tab, text=note, justify="left", wraplength=1000).pack(anchor="w", pady=8)

    def load_both(self):
        csv_path = filedialog.askopenfilename(
            title="Select raw mortgage CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not csv_path:
            return
        xlsx_path = filedialog.askopenfilename(
            title="Select mortgage analysis workbook",
            filetypes=[("Excel workbooks", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not xlsx_path:
            return
        self._load_csv_path(csv_path)
        self._load_workbook_path(xlsx_path)

    def import_csv(self):
        path = filedialog.askopenfilename(
            title="Select raw mortgage CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if path:
            self._load_csv_path(path)

    def import_workbook(self):
        path = filedialog.askopenfilename(
            title="Select mortgage analysis workbook",
            filetypes=[("Excel workbooks", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self._load_workbook_path(path)

    def _load_csv_path(self, path):
        try:
            df = pd.read_csv(path)
            if "application_id" not in df.columns:
                raise ValueError("CSV must contain an application_id column.")
            self.raw_df = df.drop_duplicates("application_id", keep="last")
            self._merge_data()
            self.status.config(text=f"Raw CSV loaded: {len(self.raw_df)} applications")
            self._update_data_info()
        except Exception as exc:
            messagebox.showerror("CSV import failed", str(exc))

    def _load_workbook_path(self, path):
        try:
            sheets = pd.read_excel(path, sheet_name=None)
            cleaned = sheets.get("Cleaned_Data")
            normalised = sheets.get("Normalised_Data")
            segments = sheets.get("All_Segments")
            if cleaned is None:
                raise ValueError("Workbook must contain a 'Cleaned_Data' sheet.")
            if "application_id" not in cleaned.columns:
                raise ValueError("Cleaned_Data must contain application_id.")
            self.cleaned_df = cleaned.drop_duplicates("application_id", keep="last")
            self.norm_df = (
                normalised.drop_duplicates("application_id", keep="last")
                if normalised is not None and "application_id" in normalised.columns
                else pd.DataFrame()
            )
            self.all_segments_df = (
                segments.drop_duplicates("application_id", keep="last")
                if segments is not None and "application_id" in segments.columns
                else pd.DataFrame()
            )
            self.workbook_sheets = {name: len(frame) for name, frame in sheets.items()}
            # Prefer the raw CSV for source values where present; workbook cleaned
            # fields fill gaps. Normalised data is joined separately by ID.
            self._merge_data()
            self.status.config(text=f"Workbook loaded: {len(sheets)} sheets")
            self._update_data_info()
        except Exception as exc:
            messagebox.showerror("Workbook import failed", str(exc))

    def _merge_data(self):
        cleaned = getattr(self, "cleaned_df", pd.DataFrame())
        if not self.raw_df.empty and not cleaned.empty:
            base = cleaned.set_index("application_id").combine_first(
                self.raw_df.set_index("application_id")
            ).reset_index()
        elif not self.raw_df.empty:
            base = self.raw_df.copy()
        elif not cleaned.empty:
            base = cleaned.copy()
        else:
            return
        if "application_id" not in base:
            return
        self.applications = base.drop_duplicates("application_id", keep="last").copy()
        self._calculate_risk_scores()
        self._refresh_pending()
        self._refresh_counts()
        self.refresh_processed()
        self.clear_selection()

    def _calculate_risk_scores(self):

        if self.applications.empty:
            return
        rules = {
            "F_LTV_90plus": ("loan_to_value_pct", lambda x: x >= 90),
            "F_DTI_over_0.45": ("debt_to_income_ratio", lambda x: x > 0.45),
            "F_Score_under_600": ("credit_score", lambda x: x < 600),
            "F_Util_70plus": ("credit_utilisation_pct", lambda x: x >= 70),
            "F_Prev_default": ("previous_default_flag", lambda x: x == 1),
            "F_Missed_2plus": ("missed_payments_24m", lambda x: x >= 2),
        }
        scores, counts, reasons = [], [], []
        for _, row in self.applications.iterrows():
            triggered = []
            missing = []
            for flag_name, (column, test) in rules.items():
                missing_flag = f"{column}_was_missing"
                was_imputed = False
                if missing_flag in row.index and not pd.isna(row[missing_flag]):
                    try:
                        was_imputed = int(float(row[missing_flag])) == 1
                    except (TypeError, ValueError):
                        was_imputed = False
                value = pd.to_numeric(pd.Series([row.get(column)]), errors="coerce").iloc[0]
                if pd.isna(value) or was_imputed:
                    missing.append(column)
                elif bool(test(value)):
                    triggered.append(flag_name)
            counts.append(len(triggered))
            reasons.append("; ".join(triggered) if triggered else "None of the six indicators triggered")
            if missing:
                scores.append(pd.NA)
            else:
                scores.append(max(1.0, round(10.0 - 1.5 * len(triggered), 1)))
        self.applications["risk_indicator_count"] = counts
        self.applications["risk_indicator_score_10"] = scores
        self.applications["risk_indicator_reasons"] = reasons

    def _refresh_pending(self):
        self.app_list.delete(*self.app_list.get_children())
        if self.applications.empty:
            return
        for _, row in self.applications.iterrows():
            app_id = str(row["application_id"])
            if app_id not in self.decisions:
                score = row.get("risk_indicator_score_10", pd.NA)
                score_text = "N/A" if pd.isna(score) else f"{float(score):.1f}"
                self.app_list.insert("", "end", iid=app_id, values=(app_id, score_text))

    def _refresh_counts(self):
        pending = 0 if self.applications.empty else sum(
            str(x) not in self.decisions for x in self.applications["application_id"]
        )
        total = len(self.applications)
        if hasattr(self, "pending_count"):
            self.pending_count.config(text=f"Pending: {pending}")
            self.total_count.config(text=f"Imported: {total}")

    def generate_validation_report(self):

        if self.applications.empty:
            messagebox.showinfo("No data loaded", "Import the raw CSV and analysis workbook first.")
            return
        if self.all_segments_df.empty:
            messagebox.showwarning(
                "Validation reference unavailable",
                "The workbook's All_Segments sheet has not been loaded. Import the analysis workbook first."
            )
            return

        df = self.applications.copy()
        required = {
            "loan_to_value_pct", "debt_to_income_ratio", "credit_score",
            "credit_utilisation_pct", "previous_default_flag", "missed_payments_24m"
        }
        missing = sorted(required - set(df.columns))
        if missing:
            messagebox.showerror(
                "Required fields missing",
                "Cannot recalculate the flags because these columns are missing:\n" + "\n".join(missing)
            )
            return


        numeric = {}
        for col in required:
            numeric[col] = pd.to_numeric(df[col], errors="coerce")

        fresh_flags = pd.DataFrame({
            "F_LTV_90plus": numeric["loan_to_value_pct"] >= 90,
            "F_DTI_over_0.45": numeric["debt_to_income_ratio"] > 0.45,
            "F_Score_under_600": numeric["credit_score"] < 600,
            "F_Util_70plus": numeric["credit_utilisation_pct"] >= 70,
            "F_Prev_default": numeric["previous_default_flag"] == 1,
            "F_Missed_2plus": numeric["missed_payments_24m"] >= 2,
        }).fillna(False).astype(int)
        fresh = pd.DataFrame({"application_id": df["application_id"].astype(str)})
        fresh = pd.concat([fresh.reset_index(drop=True), fresh_flags.reset_index(drop=True)], axis=1)
        fresh["risk_flag_count_fresh"] = fresh_flags.sum(axis=1)
        fresh["risk_tier_fresh"] = fresh["risk_flag_count_fresh"].map(
            lambda n: "Low" if n == 0 else ("Moderate" if n == 1 else "High")
        )

        ref = self.all_segments_df.copy()
        ref["application_id"] = ref["application_id"].astype(str)
        merged = fresh.merge(ref, on="application_id", how="inner", suffixes=("_fresh", "_workbook"))
        if merged.empty:
            messagebox.showwarning(
                "No matching applications",
                "No application IDs in the imported data matched the workbook's All_Segments sheet."
            )
            return

        lines = [
            "FRESH RISK FLAG VALIDATION, DATASET CHECK ONLY",
            "",
            f"Fresh records: {len(fresh)}",
            f"Matched to All_Segments: {len(merged)}",
            "",
            "Rules recalculated:",
            "• LTV >= 90%",
            "• Debt-to-income ratio > 0.45",
            "• Credit score < 600",
            "• Credit utilisation >= 70%",
            "• Previous default flag = 1",
            "• Missed payments in 24 months >= 2",
            "Tier: 0 flags = Low; 1 flag = Moderate; 2+ flags = High.",
            "",
        ]

        count_col = "risk_flag_count"
        tier_col = "risk_tier"
        if count_col in merged.columns:
            wb_count = pd.to_numeric(merged[count_col], errors="coerce")
            count_match = wb_count.eq(merged["risk_flag_count_fresh"])
            valid_count = wb_count.notna()
            lines.append(
                f"Risk flag count agreement: {int(count_match[valid_count].sum())} / "
                f"{int(valid_count.sum())} ({(count_match[valid_count].mean() * 100):.1f}%)"
                if valid_count.any() else "Risk flag count agreement: workbook count values unavailable."
            )
        else:
            lines.append("Risk flag count agreement: no risk_flag_count column in All_Segments.")

        if tier_col in merged.columns:
            tier_match = merged[tier_col].astype(str).str.strip().str.casefold().eq(
                merged["risk_tier_fresh"].astype(str).str.casefold()
            )
            lines.append(
                f"Risk tier agreement: {int(tier_match.sum())} / {len(merged)} "
                f"({tier_match.mean() * 100:.1f}%)"
            )
        else:
            lines.append("Risk tier agreement: no risk_tier column in All_Segments.")

        flag_cols = list(fresh_flags.columns)
        available_flags = [c for c in flag_cols if c in ref.columns]
        if available_flags:
            mismatch_details = []
            for col in available_flags:

                workbook_col = col + "_workbook" if col + "_workbook" in merged.columns else col
                fresh_col = col + "_fresh" if col + "_fresh" in merged.columns else col
                if workbook_col in merged.columns and fresh_col in merged.columns:
                    wb = pd.to_numeric(merged[workbook_col], errors="coerce")
                    fr = pd.to_numeric(merged[fresh_col], errors="coerce")
                    ok = wb.notna() & fr.notna()
                    mismatches = int((wb[ok] != fr[ok]).sum())
                    mismatch_details.append(f"{col}: {mismatches} mismatch(es)")
            if mismatch_details:
                lines.extend(["", "Individual flag comparison:", *mismatch_details])


        report = fresh.merge(ref, on="application_id", how="left", suffixes=("_fresh", "_workbook"))
        save_path = filedialog.asksaveasfilename(
            title="Save fresh risk validation report",
            defaultextension=".csv",
            initialfile="fresh_risk_validation_report.csv",
            filetypes=[("CSV files", "*.csv")]
        )
        if save_path:
            try:
                report.to_csv(save_path, index=False)
                lines.extend(["", f"Detailed comparison report saved to: {save_path}"])
            except Exception as exc:
                messagebox.showerror("Could not save report", str(exc))
                return
        lines.extend([

        ])
        messagebox.showinfo("Validation results", "\n".join(lines))

    def export_application_scores(self):

        if self.applications.empty:
            messagebox.showinfo("No data loaded", "Import the CSV and analysis workbook first.")
            return
        path = filedialog.asksaveasfilename(
            title="Export application risk-indicator scores",
            defaultextension=".csv",
            initialfile="mortgage_application_risk_indicator_scores.csv",
            filetypes=[("CSV files", "*.csv")]
        )
        if not path:
            return
        columns = [c for c in [
            "application_id", "risk_indicator_score_10", "risk_indicator_count",
            "risk_indicator_reasons"
        ] if c in self.applications.columns]
        try:
            self.applications[columns].to_csv(path, index=False)
            messagebox.showinfo(
                "Export complete",
                f"Exported scores for {len(self.applications)} applications.\n\n{path}\n\n"
                "A blank score means one or more required source values were missing or imputed."
            )
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))

    def _update_data_info(self):
        lines = []
        if not self.raw_df.empty:
            lines.append(f"Raw CSV: {len(self.raw_df)} records; {len(self.raw_df.columns)} columns.")
        if not self.norm_df.empty:
            lines.append(f"Normalised_Data: {len(self.norm_df)} records; {len(self.norm_df.columns)} columns.")
        if hasattr(self, "cleaned_df"):
            lines.append(f"Cleaned_Data: {len(self.cleaned_df)} records; {len(self.cleaned_df.columns)} columns.")
        if hasattr(self, "workbook_sheets"):
            lines.append("Workbook sheets: " + ", ".join(
                f"{name} ({count} rows)" for name, count in self.workbook_sheets.items()
            ))
        lines.append("")
        lines.append("Imported application fields:")
        if not self.applications.empty:
            lines.extend(f"• {col}" for col in self.applications.columns)
        self.data_info.config(state="normal")
        self.data_info.delete("1.0", "end")
        self.data_info.insert("1.0", "\n".join(lines))
        self.data_info.config(state="disabled")

    def on_select(self, _event=None):
        selected = self.app_list.selection()
        if not selected:
            return
        self.current_id = selected[0]
        row = self.applications[self.applications["application_id"].astype(str) == self.current_id]
        if row.empty:
            return
        record = row.iloc[0]
        self.detail_text.config(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.insert("end", f"Application {self.current_id}\n", "heading")
        score = record.get("risk_indicator_score_10", pd.NA)
        flag_count = record.get("risk_indicator_count", pd.NA)
        self.detail_text.insert("end", "RISK INDICATOR SCORE, SCREENING ONLY\n", "heading")
        if pd.isna(score):
            self.detail_text.insert(
                "end", "Score: Insufficient data. One or more required source values are missing or were imputed.\n"
            )
        else:
            self.detail_text.insert(
                "end", f"Score: {float(score):.1f} / 10  |  Indicators triggered: {int(flag_count)} of 6\n"
            )
        self.detail_text.insert(
            "end", "Higher scores mean fewer of these six indicators were triggered. "
                  "This is not a probability or an accept/decline recommendation.\n\n"
        )
        self.detail_text.insert("end", "Source information\n\n", "heading")

        preferred = [
            "annual_income_gbp", "property_value_gbp", "deposit_gbp",
            "requested_loan_gbp", "loan_to_value_pct",
            "monthly_debt_commitments_gbp", "debt_to_income_ratio",
            "credit_score", "employment_type", "employment_years",
            "savings_gbp", "missed_payments_24m", "credit_utilisation_pct",
            "previous_default_flag", "term_years", "interest_rate_pct",
            "application_channel"
        ]
        fields = [f for f in preferred if f in record.index]
        fields += [f for f in record.index if f not in fields
                   and f not in EXCLUDED_REVIEW_FIELDS
                   and not f.endswith("_was_missing")
                   and f != "application_id"
                   and not f.endswith("_minmax") and not f.endswith("_z")]
        for field in fields:
            label = FIELD_LABELS.get(field, field.replace("_", " ").title())
            self.detail_text.insert("end", f"{label}: ", "heading")
            self.detail_text.insert("end", f"{display_value(field, record[field])}\n")

        self.detail_text.insert("end", "\nRisk indicator breakdown\n", "heading")
        indicator_columns = [
            ("F_LTV_90plus", "LTV >= 90%"),
            ("F_DTI_over_0.45", "Debt-to-income ratio > 0.45"),
            ("F_Score_under_600", "Credit score < 600"),
            ("F_Util_70plus", "Credit utilisation >= 70%"),
            ("F_Prev_default", "Previous default flagged"),
            ("F_Missed_2plus", "2+ missed payments in 24 months"),
        ]
        for col, label in indicator_columns:
            source_col = {
                "F_LTV_90plus": "loan_to_value_pct",
                "F_DTI_over_0.45": "debt_to_income_ratio",
                "F_Score_under_600": "credit_score",
                "F_Util_70plus": "credit_utilisation_pct",
                "F_Prev_default": "previous_default_flag",
                "F_Missed_2plus": "missed_payments_24m",
            }[col]
            missing_flag = f"{source_col}_was_missing"
            missing_value = (source_col not in record.index or pd.isna(record.get(source_col)) or
                             (missing_flag in record.index and not pd.isna(record.get(missing_flag)) and
                              str(record.get(missing_flag)).strip() in {"1", "1.0", "True", "true"}))
            if missing_value:
                status = "Not assessed, missing/imputed source data"
            else:
                value = pd.to_numeric(pd.Series([record.get(source_col)]), errors="coerce").iloc[0]
                tests = {
                    "F_LTV_90plus": value >= 90,
                    "F_DTI_over_0.45": value > 0.45,
                    "F_Score_under_600": value < 600,
                    "F_Util_70plus": value >= 70,
                    "F_Prev_default": value == 1,
                    "F_Missed_2plus": value >= 2,
                }
                status = "Triggered" if tests[col] else "Not triggered"
            self.detail_text.insert("end", f"• {label}: {status}\n")


        missing_flags = [f for f in record.index if f.endswith("_was_missing") and
                         not pd.isna(record[f]) and int(record[f]) == 1]
        missing_fields = [f for f in preferred if f in record.index and pd.isna(record[f])]
        self.detail_text.insert("end", "\nData quality checks\n", "heading")
        if missing_flags:
            self.detail_text.insert(
                "end", "Source cleaning flags: " + ", ".join(x.replace("_was_missing", "") for x in missing_flags) + "\n"
            )
        if missing_fields:
            self.detail_text.insert("end", "Not supplied: " + ", ".join(missing_fields) + "\n")
        if not missing_flags and not missing_fields:
            self.detail_text.insert("end", "No missingness flags found in the imported fields.\n")

        if not self.norm_df.empty:
            norm = self.norm_df[self.norm_df["application_id"].astype(str) == self.current_id]
            if not norm.empty:
                norm_cols = [c for c in norm.columns if c.endswith("_minmax") or c.endswith("_z")]
                self.detail_text.insert("end", "\nNormalised data available\n", "heading")
                self.detail_text.insert(
                    "end",
                    f"{len(norm_cols)} normalised feature columns are linked to this application. "
                    "They are not used to generate the risk-indicator score. See the workbook's "
                    "Normalised_Data sheet for the underlying values.\n"
                )

        self.detail_text.insert(
            "end",
            "\nReviewer action required\n",
            "heading"
        )
        self.detail_text.insert(
            "end",
            "Use the score only as a prompt for further review. Check the source evidence "
            "against your organisation's current, approved underwriting policy and record "
            "your own decision and rationale using the buttons below.\n",
            "notice"
        )
        self.detail_text.config(state="disabled")

    def record_decision(self, decision):
        if not self.current_id:
            messagebox.showinfo("Select an application", "Please select a pending application first.")
            return
        reviewer = simpledialog.askstring("Reviewer", "Enter your name or staff ID:")
        if not reviewer or not reviewer.strip():
            return
        reason = simpledialog.askstring(
            f"{decision} rationale",
            "Record the specific reason for this decision:"
        )
        if not reason or not reason.strip():
            messagebox.showwarning("Reason required", "A reason is required to record a decision.")
            return
        confirm = messagebox.askyesno(
            "Confirm decision",
            f"Record {decision.lower()} for {self.current_id}?\n\n"
            "This records your decision; the application does not make the decision for you."
        )
        if not confirm:
            return
        self.db.save(self.current_id, decision, reviewer.strip(), reason.strip())
        self.decisions = self.db.all()
        self._refresh_pending()
        self._refresh_counts()
        self.refresh_processed()
        self.clear_selection()
        messagebox.showinfo("Decision recorded", f"{decision} decision recorded.")

    def refresh_processed(self):
        if not hasattr(self, "history"):
            return
        self.history.delete(*self.history.get_children())
        self.decisions = self.db.all()
        for app_id, item in sorted(self.decisions.items(), key=lambda x: x[1]["decided_at"], reverse=True):
            score_text = "N/A"
            if not self.applications.empty:
                app_row = self.applications[self.applications["application_id"].astype(str) == str(app_id)]
                if not app_row.empty:
                    score = app_row.iloc[0].get("risk_indicator_score_10", pd.NA)
                    if not pd.isna(score):
                        score_text = f"{float(score):.1f}"
            self.history.insert("", "end", values=(
                app_id, score_text, item["decision"], item["reviewer"],
                item["decided_at"], item["reason"]
            ))

    def clear_selection(self):
        self.current_id = None
        if hasattr(self, "app_list"):
            for item in self.app_list.selection():
                self.app_list.selection_remove(item)
        if hasattr(self, "detail_text"):
            self.detail_text.config(state="normal")
            self.detail_text.delete("1.0", "end")
            self.detail_text.insert(
                "1.0", "Select an application from the list to inspect its source information."
            )
            self.detail_text.config(state="disabled")


def main():
    root = tk.Tk()
    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    MortgageReviewApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
