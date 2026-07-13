from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd
import yfinance as yf
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

from covariance_matrix.core import (
    _cell,
    _configure_formats,
    _extract_adjusted_close,
    _quote_sheet,
    _range,
    _weighted_covariance_formula,
    _write_formula,
    _write_formula_matrix,