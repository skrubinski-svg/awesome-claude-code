import os
from pathlib import Path

import pytest
from fsc import eia, schedules

ROOT = Path(__file__).resolve().parent.parent


@pytest.fixture(scope="session")
def params():
    return schedules.load_params(ROOT / "config.json")


@pytest.fixture(scope="session")
def eia_prices():
    """Real price data, only when requested by `fsc_report.py --check` (or the env vars)."""
    csv_path = os.environ.get("FSC_EIA_CSV")
    if csv_path:
        return eia.load_csv(csv_path)
    if os.environ.get("FSC_EIA_LIVE") != "1":
        pytest.skip("live EIA data not requested (use `python fsc_report.py --check`)")
    try:
        return eia.parse_eia_xls(eia.download())
    except eia.FscError as e:
        pytest.skip(str(e))
