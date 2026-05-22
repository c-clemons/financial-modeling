"""
Data persistence layer for the CNS dashboard.
Wraps baseline_data.py + financial_calcs.py with a DataStore singleton.
"""

import copy
import json
from datetime import datetime
from pathlib import Path
import sys
from typing import Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from baseline_data import (
    DEFAULT_ASSUMPTIONS, ACTUALS_2025, ACTUALS_2025_TOTALS,
    ACTUALS_2026_QBO, NUM_2026_ACTUALS, TEAM_ROSTER,
    HISTORICAL_AR_BOBA, HISTORICAL_AR_GAP, HISTORICAL_AR_TOTAL,
    FORECAST_MONTH_LABELS, NUM_FORECAST_MONTHS,
    BALANCE_SHEET_2025,
    BOBA_VOLUME_2024, GAP_VOLUME_2024,
    BOBA_VOLUME_2025, GAP_VOLUME_2025,
)

_FORECAST_LABEL_TO_IDX = {label: i for i, label in enumerate(FORECAST_MONTH_LABELS)}
from financial_calcs import (
    generate_monthly_pl_forecast,
    generate_pl_by_location,
    generate_cash_flow_forecast,
    calculate_dashboard_metrics,
    forecast_expansion_costs,
    forecast_payroll,
)
from baseline_data import LOCATIONS

DATA_DIR = Path(__file__).parent / "data"
OVERRIDES_PATH = DATA_DIR / "user_overrides.json"
COMMITTED_PATH = DATA_DIR / "committed_actuals.json"
COMMITTED_KEYS = ("actuals_uploads", "cash_balance_uploads", "account_mapping_extras")
N = NUM_FORECAST_MONTHS


class DataStore:
    """Singleton data layer for CNS dashboard.

    Two persistence files:

    - ``user_overrides.json`` — soft overrides (assumptions, volumes, team
      roster, scenario state). Frequently rewritten.
    - ``committed_actuals.json`` — locked uploads (P&L actuals, balance-sheet
      cash actuals, learned account mappings). Only the explicit upload-page
      Commit/Revert buttons touch this file; assumption/volume/scenario writes
      never do. This guarantees a committed actuals file survives any reset of
      the soft-override layer.
    """

    _instance = None

    @classmethod
    def get(cls) -> "DataStore":
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    HISTORICAL_DEFAULTS = {
        'boba_2024': list(BOBA_VOLUME_2024),
        'gap_2024': list(GAP_VOLUME_2024),
        'boba_2025': list(BOBA_VOLUME_2025),
        'gap_2025': list(GAP_VOLUME_2025),
    }

    def __init__(self):
        self.defaults = copy.deepcopy(DEFAULT_ASSUMPTIONS)
        self.overrides = {}
        self.committed = {}
        self.merged = {}
        self._baseline_actuals_2025 = copy.deepcopy(ACTUALS_2025)
        self._baseline_actuals_2025_totals = copy.deepcopy(ACTUALS_2025_TOTALS)
        self._baseline_actuals_2026 = copy.deepcopy(ACTUALS_2026_QBO)
        self.team_roster = copy.deepcopy(TEAM_ROSTER)
        self.balance_sheet_2025 = BALANCE_SHEET_2025
        self.historical_ar = HISTORICAL_AR_TOTAL
        self._loaded = False
        # Auto-load so a freshly constructed instance never starts with empty
        # state. Without this, a Streamlit module reload could rebuild the
        # singleton with empty overrides and the next save would wipe disk.
        self.load()

    def load(self):
        self.overrides = self._load_json(OVERRIDES_PATH)
        self.committed = self._load_json(COMMITTED_PATH)
        self._migrate_committed_from_overrides()
        self.merged = self._deep_merge(self.defaults, self.overrides)
        self._loaded = True

    def _migrate_committed_from_overrides(self):
        """One-time migration: pull committed-actuals keys out of the soft
        overrides file into the locked file. Safe to run repeatedly."""
        moved = False
        for key in COMMITTED_KEYS:
            if key in self.overrides:
                # Disk-locked file wins if both exist (shouldn't, but be safe).
                self.committed.setdefault(key, self.overrides[key])
                del self.overrides[key]
                moved = True
        if moved:
            self._save_json(COMMITTED_PATH, self.committed)
            self._save_json(OVERRIDES_PATH, self.overrides)

    # ------------------------------------------------------------------
    # Actuals (with upload support) — backed by the locked committed file
    # ------------------------------------------------------------------
    @property
    def actuals_2025(self) -> dict:
        uploads = self.committed.get("actuals_uploads", {}).get("2025")
        if not uploads:
            return self._baseline_actuals_2025
        merged = copy.deepcopy(self._baseline_actuals_2025)
        for k, v in uploads.get("data", {}).items():
            merged[k] = list(v)
        return merged

    @property
    def actuals_2025_totals(self) -> dict:
        uploads = self.committed.get("actuals_uploads", {}).get("2025")
        if not uploads:
            return self._baseline_actuals_2025_totals
        merged = copy.deepcopy(self._baseline_actuals_2025_totals)
        merged.update(uploads.get("totals", {}))
        return merged

    @property
    def actuals_2026(self) -> dict:
        uploads = self.committed.get("actuals_uploads", {}).get("2026")
        if not uploads:
            return self._baseline_actuals_2026
        merged = copy.deepcopy(self._baseline_actuals_2026)
        # Replace months and per-key arrays for the uploaded range
        merged["months"] = list(uploads.get("months", merged.get("months", [])))
        for k, v in uploads.get("data", {}).items():
            merged[k] = list(v)
        for k, v in uploads.get("totals", {}).items():
            merged[k] = list(v) if isinstance(v, list) else v
        return merged

    @property
    def n_actuals_2026(self) -> int:
        uploads = self.committed.get("actuals_uploads", {}).get("2026")
        if uploads and uploads.get("months"):
            return len(uploads["months"])
        return NUM_2026_ACTUALS

    def set_uploaded_actuals(self, year: int, payload: dict) -> dict:
        """Persist a parsed P&L upload as actuals for the given year.

        Writes to the locked ``committed_actuals.json`` — independent from
        the soft-override layer. payload shape:
            {'months': [...], 'data': {key: [vals]}, 'totals': {key: [vals]},
             'source_filename': str, 'uploaded_at': iso-str}

        Returns the github_sync status dict.
        """
        uploads = self.committed.setdefault("actuals_uploads", {})
        uploads[str(year)] = payload
        return self.save_committed(f"CNS dashboard: commit {year} P&L actuals")

    def clear_uploaded_actuals(self, year: int) -> dict:
        uploads = self.committed.get("actuals_uploads", {})
        if str(year) in uploads:
            del uploads[str(year)]
            return self.save_committed(f"CNS dashboard: revert {year} P&L actuals")
        return {"ok": True, "message": "nothing to clear", "sha": None, "url": None}

    def get_uploaded_actuals_meta(self, year: int) -> dict:
        return self.committed.get("actuals_uploads", {}).get(str(year), {})

    # ------------------------------------------------------------------
    # Cash balance actuals (from QBO Balance Sheet upload) — locked file
    # ------------------------------------------------------------------
    def set_uploaded_cash_balance(self, year: int, payload: dict) -> dict:
        """Persist a parsed Balance Sheet upload as monthly cash actuals.

        Writes to the locked ``committed_actuals.json``. payload shape:
            {'months': [...], 'cash_total': [...],
             'cash_by_account': {label: [...]},
             'source_filename': str, 'uploaded_at': iso-str}

        Returns the github_sync status dict.
        """
        uploads = self.committed.setdefault("cash_balance_uploads", {})
        uploads[str(year)] = payload
        return self.save_committed(f"CNS dashboard: commit {year} cash actuals")

    def clear_uploaded_cash_balance(self, year: int) -> dict:
        uploads = self.committed.get("cash_balance_uploads", {})
        if str(year) in uploads:
            del uploads[str(year)]
            return self.save_committed(f"CNS dashboard: revert {year} cash actuals")
        return {"ok": True, "message": "nothing to clear", "sha": None, "url": None}

    def get_uploaded_cash_balance_meta(self, year: int) -> dict:
        return self.committed.get("cash_balance_uploads", {}).get(str(year), {})

    def get_cash_balance_actuals_by_index(self) -> dict:
        """Map forecast-index → actual closing cash across all uploaded years.

        Used by generate_cash_flow_forecast to anchor ending_cash for months
        where we have a QBO balance sheet.
        """
        out: dict[int, float] = {}
        for payload in self.committed.get("cash_balance_uploads", {}).values():
            for m, v in zip(payload.get("months", []), payload.get("cash_total", [])):
                idx = _FORECAST_LABEL_TO_IDX.get(m)
                if idx is not None:
                    out[idx] = float(v)
        return out

    def get_cash_actuals_count(self, year: int) -> int:
        return len(self.committed.get("cash_balance_uploads", {})
                   .get(str(year), {}).get("months", []))

    # ------------------------------------------------------------------
    # Account mapping (for unknown QBO accounts on upload) — locked file
    # ------------------------------------------------------------------
    def get_account_mapping_extras(self) -> dict:
        return dict(self.committed.get("account_mapping_extras", {}))

    def add_account_mapping(self, qbo_label: str, target_key: str) -> dict:
        extras = self.committed.setdefault("account_mapping_extras", {})
        extras[qbo_label] = target_key
        return self.save_committed("CNS dashboard: learn account mapping")

    def add_account_mappings(self, mappings: dict) -> dict:
        """Bulk-add account mappings with a single committed-file write."""
        if not mappings:
            return {"ok": True, "message": "no mappings", "sha": None, "url": None}
        extras = self.committed.setdefault("account_mapping_extras", {})
        for label, target in mappings.items():
            extras[label] = target
        return self.save_committed(
            f"CNS dashboard: learn {len(mappings)} account mapping(s)"
        )

    # ------------------------------------------------------------------
    # Assumptions
    # ------------------------------------------------------------------
    def get_assumptions(self) -> dict:
        a = copy.deepcopy(self.merged)
        a['cash_balance_actuals_by_index'] = self.get_cash_balance_actuals_by_index()
        return a

    def set_assumption(self, key: str, value):
        self.overrides[key] = value
        self.merged = self._deep_merge(self.defaults, self.overrides)
        self.save_overrides()

    def set_assumptions_bulk(self, updates: dict):
        for k, v in updates.items():
            self.overrides[k] = v
        self.merged = self._deep_merge(self.defaults, self.overrides)
        self.save_overrides()

    # ------------------------------------------------------------------
    # Surgery volumes
    # ------------------------------------------------------------------
    def get_surgery_volumes(self) -> tuple:
        a = self.merged
        return (list(a['bobas_volume']), list(a['gap_volume']))

    def set_surgery_volumes(self, bobas: list, gap: list):
        self.overrides['bobas_volume'] = bobas
        self.overrides['gap_volume'] = gap
        self.merged = self._deep_merge(self.defaults, self.overrides)
        self.save_overrides()

    # ------------------------------------------------------------------
    # Historical surgery volumes (display context for prior years)
    # ------------------------------------------------------------------
    def get_historical_volumes(self, key: str) -> list:
        store = self.overrides.get('historical_volumes', {})
        if key in store:
            return list(store[key])
        return list(self.HISTORICAL_DEFAULTS[key])

    def set_historical_volumes(self, **updates):
        store = self.overrides.setdefault('historical_volumes', {})
        for k, v in updates.items():
            if k not in self.HISTORICAL_DEFAULTS:
                raise KeyError(f"Unknown historical volume key: {k}")
            expected_len = len(self.HISTORICAL_DEFAULTS[k])
            arr = [int(x or 0) for x in (v or [])]
            if len(arr) != expected_len:
                raise ValueError(
                    f"Historical {k} must have {expected_len} values, got {len(arr)}"
                )
            store[k] = arr
        self.save_overrides()

    # ------------------------------------------------------------------
    # Team roster
    # ------------------------------------------------------------------
    def get_team_roster(self) -> list:
        return self.merged.get('team_roster', copy.deepcopy(self.team_roster))

    def set_team_roster(self, roster: list):
        self.overrides['team_roster'] = roster
        self.merged = self._deep_merge(self.defaults, self.overrides)
        self.save_overrides()

    # ------------------------------------------------------------------
    # Expansions
    # ------------------------------------------------------------------
    def get_expansions(self) -> list:
        return self.merged.get('expansions', [])

    def set_expansions(self, expansions: list):
        self.overrides['expansions'] = expansions
        self.merged = self._deep_merge(self.defaults, self.overrides)
        self.save_overrides()

    # ------------------------------------------------------------------
    # Locations
    # ------------------------------------------------------------------
    def get_locations(self) -> list:
        return self.merged.get('locations', LOCATIONS)

    def get_volumes_by_location(self) -> dict:
        return self.merged.get('volumes_by_location', {})

    def set_volumes_by_location(self, volumes: dict):
        self.overrides['volumes_by_location'] = volumes
        # Also update consolidated for backwards compat
        bobas_cons = [0] * N
        gap_cons = [0] * N
        for loc_data in volumes.values():
            for i in range(N):
                bobas_cons[i] += loc_data.get('bobas', [0]*N)[i]
                gap_cons[i] += loc_data.get('gap', [0]*N)[i]
        self.overrides['bobas_volume'] = bobas_cons
        self.overrides['gap_volume'] = gap_cons
        self.merged = self._deep_merge(self.defaults, self.overrides)
        self.save_overrides()

    # ------------------------------------------------------------------
    # Forecast (runs calculations)
    # ------------------------------------------------------------------
    def run_forecast(self) -> dict:
        a = self.get_assumptions()
        pl = generate_monthly_pl_forecast(a)
        cf = generate_cash_flow_forecast(a)
        return {'pl': pl, 'cf': cf}

    def run_forecast_by_location(self) -> dict:
        a = self.get_assumptions()
        pl_by_loc = generate_pl_by_location(a)
        cf = generate_cash_flow_forecast(a)
        return {'pl_by_location': pl_by_loc, 'cf': cf}

    def run_dashboard_metrics(self) -> dict:
        return calculate_dashboard_metrics(self.get_assumptions())

    def run_expansion_detail(self) -> dict:
        return forecast_expansion_costs(self.get_assumptions())

    def run_payroll_detail(self) -> dict:
        a = self.get_assumptions()
        team = a.get('team_roster', self.team_roster)
        return forecast_payroll(
            team,
            payroll_tax_rate=a.get('payroll_tax_rate', 8.6),
            salary_annual_increase=a.get('salary_annual_increase', 5.0),
        )

    # ------------------------------------------------------------------
    # Scenarios
    # ------------------------------------------------------------------
    def save_scenario(self, name: str):
        scenario_dir = DATA_DIR / "scenarios"
        scenario_dir.mkdir(parents=True, exist_ok=True)
        data = copy.deepcopy(self.overrides)
        data["_scenario_name"] = name
        data["_saved_at"] = datetime.now().isoformat()
        self._save_json(scenario_dir / f"{name}.json", data)

    def load_scenario(self, name: str):
        path = DATA_DIR / "scenarios" / f"{name}.json"
        if not path.exists():
            raise FileNotFoundError(f"Scenario '{name}' not found")
        self.overrides = self._load_json(path)
        self.merged = self._deep_merge(self.defaults, self.overrides)
        self.save_overrides()

    def list_scenarios(self) -> list:
        scenario_dir = DATA_DIR / "scenarios"
        if not scenario_dir.exists():
            return []
        return [
            {"name": f.stem, "saved_at": self._load_json(f).get("_saved_at", "?")}
            for f in sorted(scenario_dir.glob("*.json"))
        ]

    def delete_scenario(self, name: str):
        path = DATA_DIR / "scenarios" / f"{name}.json"
        if path.exists():
            path.unlink()

    # ------------------------------------------------------------------
    # Persistence
    # ------------------------------------------------------------------
    def save_overrides(self):
        self.overrides["_last_updated"] = datetime.now().isoformat()
        self._save_json(OVERRIDES_PATH, self.overrides)

    def save_committed(self, commit_message: Optional[str] = None) -> dict:
        """Write the locked committed-actuals file and mirror it to GitHub.

        Returns the github_sync status dict so callers can surface success/
        failure in the UI. GitHub sync is a best-effort no-op when no token is
        configured (e.g. local development).
        """
        self.committed["_last_updated"] = datetime.now().isoformat()
        self._save_json(COMMITTED_PATH, self.committed)
        try:
            from dashboard import github_sync
        except Exception:
            return {"ok": False, "message": "github_sync import failed",
                    "sha": None, "url": None}
        if not github_sync.sync_enabled():
            return {"ok": False, "message": "no token configured",
                    "sha": None, "url": None}
        return github_sync.push_committed_file(
            COMMITTED_PATH,
            commit_message or f"CNS dashboard: update committed actuals ({self.committed['_last_updated']})",
        )

    # ------------------------------------------------------------------
    # Internals
    # ------------------------------------------------------------------
    @staticmethod
    def _load_json(path: Path) -> dict:
        if path.exists():
            with open(path) as f:
                return json.load(f)
        return {}

    @staticmethod
    def _save_json(path: Path, data: dict):
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w") as f:
            json.dump(data, f, indent=2, default=str)

    @staticmethod
    def _deep_merge(base: dict, override: dict) -> dict:
        result = copy.deepcopy(base)
        for key, val in override.items():
            if key.startswith("_"):
                continue
            if key in result and isinstance(result[key], dict) and isinstance(val, dict):
                result[key] = DataStore._deep_merge(result[key], val)
            else:
                result[key] = copy.deepcopy(val)
        return result
