"""
═══════════════════════════════════════════════════════════════════════════════
 Corporate Financial Variance Analysis Engine — variance_engine.py
 Asset Management Division | Financial Reporting
 Version: 1.0.0
═══════════════════════════════════════════════════════════════════════════════

USAGE:
    python variance_engine.py --financial financials.csv --kpi kpi.csv
    python variance_engine.py --financial financials.xlsx --kpi kpi.xlsx --output reports/

DEPLOYMENT:
    1. pip install -r requirements.txt
    2. Set env: DATABASE_URL, SLACK_WEBHOOK, SMTP_HOST
    3. Schedule: 0 6 1 * * python variance_engine.py --financial ... --kpi ...

requirements.txt:
    pandas>=2.0
    numpy>=1.24
    openpyxl>=3.1
    sqlalchemy>=2.0
    python-dotenv>=1.0
═══════════════════════════════════════════════════════════════════════════════
"""

import pandas as pd
import numpy as np
import sqlite3
import logging
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("logs/variance.log", mode="a", encoding="utf-8")
    ]
)
logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
CONFIG = {
    "variance_thresholds": {
        "favorable":    2.0,
        "unfavorable": -2.0,
        "critical":   -10.0,
    },
    "kpi_tolerance": 0.05,
    "required_financial_cols": [
        "MONTH", "REGION", "SEGMENT",
        "REVENUE_BUDGET", "REVENUE_ACTUAL",
        "EXPENSE_BUDGET", "EXPENSE_ACTUAL"
    ],
    "required_kpi_cols": ["KPI", "MONTH", "TARGET", "ACTUAL"],
}


# ─────────────────────────────────────────────────────────────────────────────
# CLASS: VarianceEngine
# ─────────────────────────────────────────────────────────────────────────────
class VarianceEngine:
    """
    Corporate Financial Variance Analysis Engine.
    Automates Budget vs Actual comparison, KPI tracking,
    regional breakdown, client reporting, and monthly report generation.
    """

    def __init__(self, db_path="variance.db"):
        self.db_path  = db_path
        self.conn     = sqlite3.connect(db_path)
        self.run_date = datetime.today().strftime("%Y-%m-%d")
        self.run_ts   = datetime.now().isoformat()
        self.results  = {}
        self._init_db()

    def _init_db(self):
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS run_log (
                run_id      INTEGER PRIMARY KEY AUTOINCREMENT,
                run_date    TEXT,
                run_ts      TEXT,
                total_records INT,
                critical_items INT,
                total_variance REAL,
                status      TEXT
            )
        """)
        self.conn.commit()

    # ── Data Loading ──────────────────────────────────────────────────────────
    def load(self, filepath: str) -> pd.DataFrame:
        ext = Path(filepath).suffix.lower()
        df  = pd.read_excel(filepath) if ext in (".xlsx", ".xls") else pd.read_csv(filepath)
        df.columns = [c.strip().upper().replace(" ", "_") for c in df.columns]
        for col in ["REVENUE_BUDGET","REVENUE_ACTUAL","EXPENSE_BUDGET","EXPENSE_ACTUAL",
                    "TARGET","ACTUAL","BUDGET"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        logger.info(f"Loaded {filepath}: {len(df)} rows")
        return df

    def validate(self, df: pd.DataFrame, required: list, name: str) -> bool:
        missing = [c for c in required if c not in df.columns]
        if missing:
            logger.error(f"{name}: Missing columns: {missing}")
            return False
        logger.info(f"{name}: Validation passed — {len(df)} records")
        return True

    # ── Core Variance Analysis ────────────────────────────────────────────────
    def compute_variance(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Budget vs Actual variance with 4-tier classification.
        FAVORABLE > ON TARGET > UNFAVORABLE > CRITICAL
        """
        logger.info("Computing variance analysis...")
        df = df.copy()
        df["REV_VARIANCE"]    = (df["REVENUE_ACTUAL"] - df["REVENUE_BUDGET"]).round(2)
        df["REV_VAR_PCT"]     = np.where(
            df["REVENUE_BUDGET"] != 0,
            (df["REV_VARIANCE"] / df["REVENUE_BUDGET"] * 100).round(2), 0)
        df["EXP_VARIANCE"]    = (df["EXPENSE_ACTUAL"] - df["EXPENSE_BUDGET"]).round(2)
        df["EXP_VAR_PCT"]     = np.where(
            df["EXPENSE_BUDGET"] != 0,
            (df["EXP_VARIANCE"] / df["EXPENSE_BUDGET"] * 100).round(2), 0)
        df["NI_ACTUAL"]       = (df["REVENUE_ACTUAL"] - df["EXPENSE_ACTUAL"]).round(2)
        df["NI_BUDGET"]       = (df["REVENUE_BUDGET"] - df["EXPENSE_BUDGET"]).round(2)
        df["NI_VARIANCE"]     = (df["NI_ACTUAL"] - df["NI_BUDGET"]).round(2)
        df["NET_MARGIN_PCT"]  = np.where(
            df["REVENUE_ACTUAL"] != 0,
            (df["NI_ACTUAL"] / df["REVENUE_ACTUAL"] * 100).round(2), 0)
        df["VARIANCE_STATUS"] = df["REV_VAR_PCT"].apply(self._classify)
        df["RISK_FLAG"]       = df["REV_VAR_PCT"].apply(
            lambda x: "CRITICAL" if x <= CONFIG["variance_thresholds"]["critical"]
            else ("WATCH" if x < CONFIG["variance_thresholds"]["unfavorable"] else "OK"))
        df["RUN_DATE"]        = self.run_date

        self.results["variance"] = df
        critical = int((df["RISK_FLAG"] == "CRITICAL").sum())
        logger.info(f"Variance complete: {len(df)} records | Critical: {critical}")
        return df

    def _classify(self, pct: float) -> str:
        t = CONFIG["variance_thresholds"]
        if pct >= t["favorable"]:   return "FAVORABLE"
        if pct >= t["unfavorable"]: return "ON TARGET"
        if pct >= t["critical"]:    return "UNFAVORABLE"
        return "CRITICAL"

    # ── Regional Breakdown ────────────────────────────────────────────────────
    def regional_breakdown(self) -> pd.DataFrame:
        df  = self.results.get("variance", pd.DataFrame())
        grp = df.groupby("REGION").agg(
            Rev_Budget  =("REVENUE_BUDGET",  "sum"),
            Rev_Actual  =("REVENUE_ACTUAL",  "sum"),
            Exp_Budget  =("EXPENSE_BUDGET",  "sum"),
            Exp_Actual  =("EXPENSE_ACTUAL",  "sum"),
            NI_Actual   =("NI_ACTUAL",       "sum"),
        ).reset_index()
        grp["Rev_Variance"] = (grp["Rev_Actual"] - grp["Rev_Budget"]).round(2)
        grp["Rev_Var_Pct"]  = (grp["Rev_Variance"] / grp["Rev_Budget"] * 100).round(2)
        grp["Net_Margin"]   = (grp["NI_Actual"]   / grp["Rev_Actual"] * 100).round(2)
        grp["Rank"]         = grp["Rev_Actual"].rank(ascending=False).astype(int)
        grp["Status"]       = grp["Rev_Var_Pct"].apply(self._classify)
        grp["RUN_DATE"]     = self.run_date
        self.results["regional"] = grp
        logger.info(f"Regional breakdown: {len(grp)} regions")
        return grp

    # ── Segment Analysis ──────────────────────────────────────────────────────
    def segment_analysis(self) -> pd.DataFrame:
        df  = self.results.get("variance", pd.DataFrame())
        grp = df.groupby("SEGMENT").agg(
            Rev_Budget=("REVENUE_BUDGET","sum"),
            Rev_Actual=("REVENUE_ACTUAL","sum"),
            NI_Actual =("NI_ACTUAL","sum"),
        ).reset_index()
        grp["Rev_Variance"] = (grp["Rev_Actual"] - grp["Rev_Budget"]).round(2)
        grp["Rev_Var_Pct"]  = (grp["Rev_Variance"] / grp["Rev_Budget"] * 100).round(2)
        grp["NI_Margin"]    = (grp["NI_Actual"] / grp["Rev_Actual"] * 100).round(2)
        grp["RUN_DATE"]     = self.run_date
        self.results["segment"] = grp
        return grp

    # ── Trend Analysis ────────────────────────────────────────────────────────
    def trend_analysis(self) -> pd.DataFrame:
        df    = self.results.get("variance", pd.DataFrame())
        gcols = ["MONTH"] + (["MONTH_NUM"] if "MONTH_NUM" in df.columns else [])
        trend = df.groupby(gcols).agg(
            Rev_Budget=("REVENUE_BUDGET","sum"),
            Rev_Actual=("REVENUE_ACTUAL","sum"),
            Exp_Actual=("EXPENSE_ACTUAL","sum"),
        ).reset_index()
        if "MONTH_NUM" in trend.columns:
            trend = trend.sort_values("MONTH_NUM")
        trend["NI_Actual"]  = (trend["Rev_Actual"] - trend["Exp_Actual"]).round(2)
        trend["Rev_Var"]    = (trend["Rev_Actual"] - trend["Rev_Budget"]).round(2)
        trend["Rev_Var_Pct"]= (trend["Rev_Var"] / trend["Rev_Budget"] * 100).round(2)
        trend["MoM_Growth"] = (trend["Rev_Actual"].pct_change() * 100).round(2)
        trend["Rolling_3M"] = trend["Rev_Actual"].rolling(3, min_periods=1).mean().round(2)
        trend["NI_Margin"]  = (trend["NI_Actual"] / trend["Rev_Actual"] * 100).round(2)
        trend["RUN_DATE"]   = self.run_date
        self.results["trend"] = trend
        logger.info(f"Trend analysis: {len(trend)} periods")
        return trend

    # ── KPI Scorecard ─────────────────────────────────────────────────────────
    def kpi_scorecard(self, kpi_df: pd.DataFrame) -> pd.DataFrame:
        df = kpi_df.copy()
        df["VARIANCE"]      = (df["ACTUAL"] - df["TARGET"]).round(4)
        df["VAR_PCT"]       = np.where(
            df["TARGET"] != 0,
            (df["VARIANCE"] / df["TARGET"] * 100).round(2), 0)
        tol = CONFIG["kpi_tolerance"] * 100
        df["STATUS"]        = df["VAR_PCT"].apply(
            lambda x: "ABOVE" if x > tol else ("ON TARGET" if x >= -tol else "BELOW"))
        df["TRAFFIC_LIGHT"] = df["STATUS"].map(
            {"ABOVE":"GREEN","ON TARGET":"AMBER","BELOW":"RED"})
        df["RUN_DATE"]      = self.run_date
        self.results["kpi"] = df
        below = int((df["STATUS"] == "BELOW").sum())
        logger.info(f"KPI scorecard: {len(df)} KPIs | Below target: {below}")
        return df

    # ── Summary Stats ─────────────────────────────────────────────────────────
    def summary_stats(self) -> dict:
        df = self.results.get("variance", pd.DataFrame())
        if df.empty: return {}
        return {
            "run_date":              self.run_date,
            "total_records":         len(df),
            "total_revenue_actual":  round(df["REVENUE_ACTUAL"].sum(), 2),
            "total_revenue_budget":  round(df["REVENUE_BUDGET"].sum(), 2),
            "revenue_variance":      round(df["REV_VARIANCE"].sum(), 2),
            "revenue_var_pct":       round(df["REV_VAR_PCT"].mean(), 2),
            "total_ni_actual":       round(df["NI_ACTUAL"].sum(), 2),
            "avg_net_margin_pct":    round(df["NET_MARGIN_PCT"].mean(), 2),
            "critical_items":        int((df["RISK_FLAG"] == "CRITICAL").sum()),
            "watch_items":           int((df["RISK_FLAG"] == "WATCH").sum()),
            "status_breakdown":      df["VARIANCE_STATUS"].value_counts().to_dict(),
            "risk_breakdown":        df["RISK_FLAG"].value_counts().to_dict(),
        }

    # ── Persistence ───────────────────────────────────────────────────────────
    def save_to_db(self):
        for name, df in self.results.items():
            df.to_sql(name, self.conn, if_exists="replace", index=False)
        stats = self.summary_stats()
        self.conn.execute("""
            INSERT INTO run_log (run_date,run_ts,total_records,critical_items,total_variance,status)
            VALUES (?,?,?,?,?,?)
        """, (self.run_date, self.run_ts,
              stats.get("total_records",0),
              stats.get("critical_items",0),
              stats.get("revenue_variance",0),
              "COMPLETE"))
        self.conn.commit()
        logger.info(f"Saved to {self.db_path}")

    # ── Export ────────────────────────────────────────────────────────────────
    def export_all(self, output_dir="reports/"):
        out = Path(output_dir)
        out.mkdir(parents=True, exist_ok=True)
        d   = self.run_date

        for name, df in self.results.items():
            df.to_csv(out / f"{name}_{d}.csv", index=False)
            logger.info(f"Exported {name} -> {out / f'{name}_{d}.csv'}")

        # Multi-sheet Excel (Power BI ready)
        with pd.ExcelWriter(out / f"variance_report_{d}.xlsx", engine="openpyxl") as writer:
            for name, df in self.results.items():
                df.to_excel(writer, sheet_name=name[:31], index=False)
        logger.info(f"Excel report -> {out / f'variance_report_{d}.xlsx'}")

        # Summary JSON
        with open(out / f"summary_{d}.json", "w") as f:
            json.dump(self.summary_stats(), f, indent=2)
        logger.info(f"Summary -> {out / f'summary_{d}.json'}")

    # ── Alerts ────────────────────────────────────────────────────────────────
    def send_alerts(self):
        stats = self.summary_stats()
        if stats.get("critical_items", 0) > 0:
            msg = (
                f"CRITICAL VARIANCE ALERT - {self.run_date}\n"
                f"Critical items: {stats['critical_items']}\n"
                f"Revenue variance: ${stats['revenue_variance']:,.0f}\n"
                f"Avg variance %: {stats['revenue_var_pct']}%\n"
                f"Immediate CFO review required."
            )
            logger.warning(msg)
            # Plug in SMTP/Slack here
        else:
            logger.info("No critical items - no alerts sent")


# ─────────────────────────────────────────────────────────────────────────────
# PIPELINE: Monthly automated run
# ─────────────────────────────────────────────────────────────────────────────
def monthly_run(financial_file: str, kpi_file: str,
                output_dir: str = "reports/",
                db_path: str = "variance.db") -> dict:
    logger.info("=" * 60)
    logger.info(" CORPORATE FINANCIAL VARIANCE ANALYSIS - MONTHLY RUN")
    logger.info("=" * 60)

    engine  = VarianceEngine(db_path=db_path)
    fin_df  = engine.load(financial_file)
    kpi_df  = engine.load(kpi_file)

    if not engine.validate(fin_df, CONFIG["required_financial_cols"], "Financials"):
        raise ValueError("Financial data validation failed")
    if not engine.validate(kpi_df, CONFIG["required_kpi_cols"], "KPI"):
        raise ValueError("KPI data validation failed")

    engine.compute_variance(fin_df)
    engine.regional_breakdown()
    engine.segment_analysis()
    engine.trend_analysis()
    engine.kpi_scorecard(kpi_df)

    engine.save_to_db()
    engine.export_all(output_dir)
    engine.send_alerts()

    stats = engine.summary_stats()
    logger.info(f"DONE - Revenue variance: ${stats['revenue_variance']:,.0f} | "
                f"Critical: {stats['critical_items']}")
    logger.info("=" * 60)
    return stats


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Corporate Financial Variance Engine")
    parser.add_argument("--financial", required=True, help="Financial data file (CSV/XLSX)")
    parser.add_argument("--kpi",       required=True, help="KPI data file (CSV/XLSX)")
    parser.add_argument("--output",    default="reports/", help="Output directory")
    parser.add_argument("--db",        default="variance.db", help="SQLite DB path")
    args = parser.parse_args()

    Path("logs").mkdir(exist_ok=True)
    stats = monthly_run(args.financial, args.kpi, args.output, args.db)
    print(json.dumps(stats, indent=2))
