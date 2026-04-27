from datetime import date, datetime
from pathlib import Path
import os

from dotenv import load_dotenv
from fredapi import Fred
import pandas as pd


def get_fred_client() -> Fred:
    load_dotenv()
    api_key = os.getenv("FRED_API_KEY")
    if not api_key:
        raise RuntimeError("FRED_API_KEY not found in environment or .env file")
    return Fred(api_key=api_key.strip())


def pull_series(
    fred: Fred,
    series_map: dict[str, str],
    start: datetime | date,
    end: datetime | date,
) -> pd.DataFrame:
    frames = {
        name: fred.get_series(sid, observation_start=start, observation_end=end)
        for name, sid in series_map.items()
    }
    df = pd.DataFrame(frames)
    df.index.name = "Date"
    return df.reset_index()


def resolve_output_dir(script_file: str | Path, topic: str) -> Path:
    out_dir = Path(script_file).resolve().parent / "output" / topic
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir


def get_recession_periods(
    fred: Fred,
    start: datetime | date,
    end: datetime | date,
) -> list[tuple[pd.Timestamp, pd.Timestamp]]:
    """Return NBER recession periods from FRED's USREC series as (start, end) tuples."""
    rec = fred.get_series("USREC", observation_start=start, observation_end=end)
    rec = rec.dropna().astype(int)
    periods: list[tuple[pd.Timestamp, pd.Timestamp]] = []
    in_recession = False
    cur_start: pd.Timestamp | None = None
    prev_idx: pd.Timestamp | None = None
    for idx, val in rec.items():
        if val == 1 and not in_recession:
            cur_start = idx
            in_recession = True
        elif val == 0 and in_recession:
            periods.append((cur_start, prev_idx))
            in_recession = False
            cur_start = None
        prev_idx = idx
    if in_recession and cur_start is not None and prev_idx is not None:
        periods.append((cur_start, prev_idx))
    return periods


def style_macro_chart(
    ax,
    *,
    title: str,
    ylabel: str,
    ylim: tuple[float, float] | None = None,
    recessions: list[tuple[pd.Timestamp, pd.Timestamp]] | None = None,
    hlines: list[dict] | None = None,
) -> None:
    """Apply shared macro-chart styling.

    hlines: list of dicts, each with keys:
      - "y" (float, required): y-value of the line
      - "label" (str | None): legend label; None to omit from legend
      - "color" (str, optional): line color, default "0.4"
      - "linestyle" (str, optional): default "--"
    """
    if recessions:
        for r_start, r_end in recessions:
            ax.axvspan(r_start, r_end, color="0.85", alpha=0.5, zorder=0)
    for h in hlines or []:
        ax.axhline(
            h["y"],
            color=h.get("color", "0.4"),
            linestyle=h.get("linestyle", "--"),
            linewidth=1,
            label=h.get("label"),
            zorder=1,
        )
    if ylim is not None:
        ax.set_ylim(*ylim)
    ax.set_title(title)
    ax.set_ylabel(ylabel)
    ax.set_xlabel("")
    ax.grid(True, alpha=0.3)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    ax.legend(loc="best", frameon=False)
