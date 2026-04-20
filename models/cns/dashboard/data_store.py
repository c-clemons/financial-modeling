"""
Data persistence layer for the CNS dashboard.
Wraps baseline_data.py + financial_calcs.py with a DataStore singleton.
"""

import copy
import json
from datetime import datetime
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parent.parent))

from baseline_data import (
    DEFAULT_ASSUMPTIONS, ACTUALS_2025, ACTUALS_2025_TOTALS,
    ACTUALS_2026_QBO, NUM_2026_ACTUALS, TEAM_ROSTER,
    HISTORICAL_AR_BOBA, HISTORICAL_AR_GAP, HISTORICAL_AR_TOTAL,
    FORECAST_MONTH_LABELS, NUM_FORECAST_MONTHS,
    BALANCE_SHEET_2025,
)
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
N = NUM_FORECAST_MONTHS


class DataStore:
    """Singleton data layer for CNS dashboard."""

    _instance = None

    @classmethod
    def get(cls) -> "DataStore":
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def __init__(self):
        self.defaults = copy.deepcopy(DEFAULT_ASSUMPTIONS)
        self.overrides = {}
        self.merged = {}
        self.actuals_2025 = ACTUALS_2025
        self.actuals_2025_totals = ACTUALS_2025_TOTALS
        self.actuals_2026 = ACTUALS_2026_QBO
        self.n_actuals_2026 = NUM_2026_ACTUALS
        self.team_roster = copy.deepcopy(TEAM_ROSTER)
        self.balance_sheet_2025 = BALANCE_SHEET_2025
        self.historical_ar = HISTORICAL_AR_TOTAL
        self._loaded = False

    def load(self):
        self.overrides = self._load_json(DATA_DIR / "user_overrides.json")
        self.merged = self._deep_merge(self.defaults, self.overrides)
        self._loaded = True

    # ------------------------------------------------------------------
    # Assumptions
    # ------------------------------------------------------------------
    def get_assumptions(self) -> dict:
        return copy.deepcopy(self.merged)

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
        self._save_json(DATA_DIR / "user_overrides.json", self.overrides)

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
