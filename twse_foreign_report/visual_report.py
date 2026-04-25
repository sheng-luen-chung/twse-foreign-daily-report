from __future__ import annotations

import html
import json
import shutil
from pathlib import Path
from typing import TYPE_CHECKING, Iterable

import pandas as pd

if TYPE_CHECKING:
    from .daily_report import BuildResult
    from .shareholder_distribution import ShareholderBuildResult


BG = "#050607"
PANEL = "#11161b"
PANEL_2 = "#1b2024"
GRID = "#20282f"
TEXT = "#f2f5f7"
MUTED = "#929aa3"
GREEN = "#62f05a"
RED = "#ff4b47"
BLUE = "#2f78a8"
ORANGE = "#dba03e"


def _e(value: object) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def _num(value: object, default: float = 0.0) -> float:
    if value is None or pd.isna(value):
        return default
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _fmt_int(value: object) -> str:
    return f"{int(round(_num(value))):,}"


def _fmt_float(value: object, digits: int = 2) -> str:
    return f"{_num(value):,.{digits}f}"


def _fmt_optional_float(value: object, digits: int = 2) -> str:
    if value is None or pd.isna(value):
        return "-"
    return _fmt_float(value, digits)


def _fmt_optional_int(value: object) -> str:
    if value is None or pd.isna(value):
        return "-"
    return _fmt_int(value)


def _fmt_optional_lots(value: object) -> str:
    if value is None or pd.isna(value):
        return "-"
    return _fmt_int(_num(value) / 1000.0)


def _fmt_signed(value: object, suffix: str = "", digits: int = 0) -> str:
    n = _num(value)
    if digits:
        body = f"{n:+,.{digits}f}"
    else:
        body = f"{int(round(n)):+,}"
    return f"{body}{suffix}"


def _short_date(date: str) -> str:
    return f"{date[4:6]}/{date[6:8]}" if len(str(date)) == 8 else str(date)


def _range_label(level: int) -> str:
    labels = {
        15: "1000張以上",
        14: "800~1000",
        13: "600~800",
        12: "400~600",
        11: "200~400",
        10: "100~200",
        9: "50~100",
        8: "40~50",
        7: "30~40",
        6: "20~30",
        5: "15~20",
        4: "10~15",
        3: "5~10",
        2: "1~5",
        1: "1張以下",
    }
    return labels.get(level, str(level))


def _target_quote(result: BuildResult, code: str) -> dict[str, float]:
    row = result.merged[result.merged["code"].astype(str) == str(code)]
    if row.empty:
        return {"close": 0.0, "change": 0.0, "pct": 0.0}
    item = row.iloc[0]
    return {
        "close": _num(item.get("close")),
        "change": _num(item.get("change")),
        "pct": _num(item.get("pct")),
    }


def _price_history(outdir: Path, dates: Iterable[str], code: str, latest_quote: dict[str, float]) -> dict[str, float]:
    prices: dict[str, float] = {}
    history_path = outdir / "history" / "price_history.csv"
    if history_path.exists():
        try:
            history = pd.read_csv(history_path, dtype={"date": str, "code": str})
            matched = history[history["code"].astype(str) == str(code)]
            for _, row in matched.iterrows():
                close = _num(row.get("close"))
                if close:
                    prices[str(row["date"])] = close
        except Exception:
            pass

    for date in dates:
        csv_path = outdir / "archive" / date / f"twse_foreign_full_{date}.csv"
        if csv_path.exists():
            try:
                df = pd.read_csv(csv_path, dtype={"code": str})
                row = df[df["code"].astype(str) == str(code)]
                if not row.empty:
                    close = _num(row.iloc[0].get("close"))
                    if close:
                        prices[date] = close
            except Exception:
                pass
    return prices


def _market_history(outdir: Path, dates: Iterable[str], code: str) -> dict[str, dict[str, float | None]]:
    history: dict[str, dict[str, float | None]] = {}
    history_path = outdir / "history" / "price_history.csv"
    if history_path.exists():
        try:
            df = pd.read_csv(history_path, dtype={"date": str, "code": str})
            matched = df[df["code"].astype(str) == str(code)]
            for _, row in matched.iterrows():
                history[str(row["date"])] = {
                    "close": row.get("close"),
                    "volume": row.get("volume"),
                }
        except Exception:
            pass
    for date, close in _price_history(outdir, dates, code, {}).items():
        history.setdefault(date, {})["close"] = close
    return history


def _points(values: list[float], width: int, height: int, pad: int = 24) -> list[tuple[float, float]]:
    if not values:
        return []
    lo = min(values)
    hi = max(values)
    if hi == lo:
        hi = lo + 1
    step = 0 if len(values) == 1 else (width - pad * 2) / (len(values) - 1)
    return [
        (pad + i * step, pad + (hi - value) / (hi - lo) * (height - pad * 2))
        for i, value in enumerate(values)
    ]


def _polyline(points: list[tuple[float, float]]) -> str:
    return " ".join(f"{x:.1f},{y:.1f}" for x, y in points)


def _bar_chart_svg(
    dates: list[str],
    bars: list[float],
    line: list[float],
    bar_color: str,
    y_label: str,
    height: int = 310,
) -> str:
    width = 760
    pad_x = 48
    pad_y = 30
    plot_w = width - pad_x * 2
    plot_h = height - pad_y * 2
    max_bar = max(bars) if bars else 1
    min_bar = min(bars) if bars else 0
    if max_bar == min_bar:
        min_bar = 0
    gap = 9
    bar_w = max(8, (plot_w - gap * max(0, len(bars) - 1)) / max(1, len(bars)))
    line_points = _points(line, width, height, pad=pad_x) if line else []
    rects = []
    labels = []
    for i, value in enumerate(bars):
        x = pad_x + i * (bar_w + gap)
        h = (value - min_bar) / (max_bar - min_bar or 1) * (plot_h - 18) + 8
        y = pad_y + plot_h - h
        rects.append(f'<rect x="{x:.1f}" y="{y:.1f}" width="{bar_w:.1f}" height="{h:.1f}" rx="7" fill="{bar_color}" opacity="0.92"/>')
        if i % max(1, len(dates) // 6) == 0 or i == len(dates) - 1:
            labels.append(f'<text x="{x + bar_w / 2:.1f}" y="{height - 5}" text-anchor="middle" class="axis">{_short_date(dates[i])}</text>')
    poly = f'<polyline points="{_polyline(line_points)}" fill="none" stroke="{ORANGE}" stroke-width="3" opacity="0.95"/>' if line_points else ""
    return f"""
<svg class="chart-svg" viewBox="0 0 {width} {height}" role="img" aria-label="{_e(y_label)}">
  <g>{''.join(f'<line x1="{pad_x}" x2="{width - pad_x}" y1="{pad_y + j * plot_h / 5:.1f}" y2="{pad_y + j * plot_h / 5:.1f}" class="grid"/>' for j in range(6))}</g>
  <text x="8" y="26" class="axis">{_e(y_label)}</text>
  {''.join(rects)}
  {poly}
  {''.join(labels)}
</svg>"""


def _area_chart_svg(dates: list[str], area: list[float], line: list[float], y_label: str) -> str:
    width = 760
    height = 310
    points = _points(area, width, height, pad=48)
    baseline = height - 30
    area_path = ""
    if points:
        area_path = (
            f'M {points[0][0]:.1f},{baseline} L '
            + " L ".join(f"{x:.1f},{y:.1f}" for x, y in points)
            + f' L {points[-1][0]:.1f},{baseline} Z'
        )
    line_points = _points(line, width, height, pad=48) if line else []
    return f"""
<svg class="chart-svg" viewBox="0 0 {width} {height}" role="img" aria-label="{_e(y_label)}">
  <g>{''.join(f'<line x1="48" x2="{width - 48}" y1="{30 + j * 250 / 5:.1f}" y2="{30 + j * 250 / 5:.1f}" class="grid"/>' for j in range(6))}</g>
  <text x="8" y="26" class="axis">{_e(y_label)}</text>
  <path d="{area_path}" fill="{BLUE}" opacity="0.95"/>
  <polyline points="{_polyline(line_points)}" fill="none" stroke="{ORANGE}" stroke-width="3" opacity="0.95"/>
  {''.join(f'<text x="{x:.1f}" y="{height - 5}" text-anchor="middle" class="axis">{_short_date(dates[i])}</text>' for i, (x, _) in enumerate(points) if i % max(1, len(dates) // 6) == 0 or i == len(dates) - 1)}
</svg>"""


def _holder_rows(detail: pd.DataFrame, code: str, dates: list[str]) -> pd.DataFrame:
    return detail[(detail["code"].astype(str) == str(code)) & (detail["date"].astype(str).isin(dates))].copy()


def _history_rows(outdir: Path, current_detail: pd.DataFrame, code: str) -> pd.DataFrame:
    history_path = outdir / "history" / "shareholders_detail_history.csv"
    frames = [current_detail]
    if history_path.exists():
        try:
            frames.append(pd.read_csv(history_path, dtype={"code": str, "date": str}))
        except Exception:
            pass
    rows = pd.concat(frames, ignore_index=True)
    rows["code"] = rows["code"].astype(str)
    rows["date"] = rows["date"].astype(str)
    rows = rows[rows["code"] == str(code)].copy()
    if rows.empty:
        return rows
    rows = rows.drop_duplicates(subset=["code", "date", "holding_level"], keep="last")
    return rows.sort_values(["date", "holding_level"], ascending=[False, True]).reset_index(drop=True)


def _date_metric(rows: pd.DataFrame, dates: list[str], level: int, metric: str) -> list[float]:
    values = []
    for date in dates:
        item = rows[(rows["date"].astype(str) == date) & (rows["holding_level"] == level)]
        values.append(_num(item.iloc[0][metric]) if not item.empty else 0.0)
    return values


def _small_holder_ratio(rows: pd.DataFrame, date: str) -> float:
    item = rows[(rows["date"].astype(str) == date) & (rows["holding_level"].between(1, 8))]
    return float(item["ratio_pct"].sum()) if not item.empty else 0.0


def _total_holders(summary: pd.DataFrame, code: str, dates: list[str]) -> list[int]:
    row = summary[summary["code"].astype(str) == str(code)]
    if row.empty:
        return [0 for _ in dates]
    item = row.iloc[0]
    return [int(_num(item.get(f"total_holders_{date}"))) for date in dates]


def _total_holders_from_rows(rows: pd.DataFrame, dates: list[str]) -> list[int]:
    values: list[int] = []
    for date in dates:
        date_rows = rows[rows["date"].astype(str) == date]
        values.append(int(date_rows["holders"].fillna(0).sum()) if not date_rows.empty else 0)
    return values


def _total_shares_from_rows(rows: pd.DataFrame, dates: list[str]) -> list[int]:
    values: list[int] = []
    for date in dates:
        date_rows = rows[rows["date"].astype(str) == date]
        values.append(int(date_rows["shares"].fillna(0).sum()) if not date_rows.empty else 0)
    return values


def _week_turnovers(
    market_by_date: dict[str, dict[str, float | None]],
    dates_latest_first: list[str],
    total_shares_by_date: dict[str, int],
) -> dict[str, float | None]:
    out: dict[str, float | None] = {}
    ordered_market_dates = sorted(market_by_date)
    for index, date in enumerate(dates_latest_first):
        previous = dates_latest_first[index + 1] if index + 1 < len(dates_latest_first) else None
        if previous:
            period_dates = [item for item in ordered_market_dates if previous < item <= date]
        else:
            period_dates = [item for item in ordered_market_dates if item <= date]
        volume = sum(_num(market_by_date[item].get("volume")) for item in period_dates)
        shares = total_shares_by_date.get(date, 0)
        out[date] = volume / shares * 100.0 if shares else None
    return out


def _latest_first(values_by_chronological: list[object]) -> list[object]:
    return list(reversed(values_by_chronological))


def _changes(values: list[float]) -> list[float]:
    return [0.0] + [values[i] - values[i - 1] for i in range(1, len(values))]


def _streak_label(changes_latest_first: list[float], suffix: str = "") -> tuple[str, str]:
    nonzero = [change for change in changes_latest_first if abs(float(change)) > 1e-9]
    if not nonzero:
        return "持平", "flat"
    sign = 1 if nonzero[0] > 0 else -1
    count = 0
    for change in changes_latest_first:
        if abs(float(change)) <= 1e-9:
            break
        if (change > 0 and sign > 0) or (change < 0 and sign < 0):
            count += 1
        else:
            break
    word = "增" if sign > 0 else "減"
    tone = "up" if sign > 0 else "down"
    return f"連 {count} {word}{suffix}", tone


def _display_name(shareholder_result: ShareholderBuildResult, code: str, fallback: str) -> str:
    summary = shareholder_result.summary
    if "code" in summary.columns and "name" in summary.columns:
        matched = summary[summary["code"].astype(str) == str(code)]
        if not matched.empty:
            name = str(matched.iloc[0]["name"]).strip()
            if name and "?" not in name:
                return name
    return fallback


def _pyramid_dataset(rows: pd.DataFrame, dates: list[str]) -> dict[str, object]:
    dataset: dict[str, object] = {"dates": dates, "weeks": {}}
    for index, date in enumerate(dates):
        current_rows = rows[rows["date"].astype(str) == date].sort_values("holding_level", ascending=False)
        previous_date = dates[index + 1] if index + 1 < len(dates) else date
        previous_rows = rows[rows["date"].astype(str) == previous_date]
        previous_by_level = {int(row["holding_level"]): row for _, row in previous_rows.iterrows()}

        total_holders = int(current_rows["holders"].fillna(0).sum()) if not current_rows.empty else 0
        previous_total = int(previous_rows["holders"].fillna(0).sum()) if not previous_rows.empty else total_holders
        max_holders = max([_num(v) for v in current_rows["holders"]] or [1])
        max_ratio = max([_num(v) for v in current_rows["ratio_pct"]] or [1])

        row_items = []
        for _, row in current_rows.iterrows():
            level = int(row["holding_level"])
            previous = previous_by_level.get(level)
            holders = _num(row["holders"])
            ratio = _num(row["ratio_pct"])
            previous_holders = _num(previous["holders"]) if previous is not None else holders
            previous_ratio = _num(previous["ratio_pct"]) if previous is not None else ratio
            row_items.append(
                {
                    "level": level,
                    "range": _range_label(level),
                    "holders": int(round(holders)),
                    "holdersChange": int(round(holders - previous_holders)),
                    "holdersWidth": 0 if max_holders == 0 else holders / max_holders * 100,
                    "ratio": ratio,
                    "ratioChange": ratio - previous_ratio,
                    "ratioWidth": 0 if max_ratio == 0 else ratio / max_ratio * 100,
                }
            )

        dataset["weeks"][date] = {
            "date": date,
            "shortDate": _short_date(date),
            "previousDate": previous_date,
            "totalHolders": total_holders,
            "totalHoldersChange": total_holders - previous_total,
            "rows": row_items,
        }
    return dataset


def _render_target(
    result: BuildResult,
    shareholder_result: ShareholderBuildResult,
    outdir: Path,
    code: str,
    name: str,
) -> str:
    name = _display_name(shareholder_result, code, name)
    rows = _history_rows(outdir, shareholder_result.detail, code)
    dates = sorted(rows["date"].astype(str).unique(), reverse=True) if not rows.empty else list(shareholder_result.selected_dates)
    chronological = list(reversed(dates))
    latest = dates[0]
    previous = dates[1] if len(dates) > 1 else dates[0]
    quote = _target_quote(result, code)
    market_by_date = _market_history(outdir, chronological, code)
    prices_by_date = {date: data.get("close") for date, data in market_by_date.items()}
    prices = [prices_by_date.get(date) for date in chronological]
    price_line = prices if len(prices) > 1 and all(price is not None for price in prices) else []

    big_ratios = _date_metric(rows, chronological, 15, "ratio_pct")
    small_ratios = [_small_holder_ratio(rows, date) for date in chronological]
    total_holders = _total_holders_from_rows(rows, chronological)
    total_shares = _total_shares_from_rows(rows, chronological)
    total_changes = _changes([float(value) for value in total_holders])
    big_changes = _changes(big_ratios)
    small_changes = _changes(small_ratios)
    avg_lots = [
        total_shares[i] / total_holders[i] / 1000.0 if total_holders[i] else 0.0
        for i in range(len(chronological))
    ]
    avg_lot_changes = _changes(avg_lots)
    display_dates = list(reversed(chronological))
    display_big_ratios = _latest_first(big_ratios)
    display_big_changes = _latest_first(big_changes)
    display_small_ratios = _latest_first(small_ratios)
    display_small_changes = _latest_first(small_changes)
    display_total_holders = _latest_first(total_holders)
    display_total_changes = _latest_first(total_changes)
    display_avg_lots = _latest_first(avg_lots)
    display_avg_lot_changes = _latest_first(avg_lot_changes)
    display_prices = _latest_first(prices)
    display_volumes = [market_by_date.get(date, {}).get("volume") for date in display_dates]
    total_shares_by_date = {date: total_shares[i] for i, date in enumerate(chronological)}
    turnovers_by_date = _week_turnovers(market_by_date, dates, total_shares_by_date)
    display_turnovers = [turnovers_by_date.get(date) for date in display_dates]
    big_streak, big_streak_tone = _streak_label(display_big_changes, "%")
    small_streak, small_streak_tone = _streak_label(display_small_changes, "%")

    big_table = "".join(
        f"""
        <tr>
          <td>{_short_date(date)}</td>
          <td>{_fmt_float(display_big_ratios[i])}%</td>
          <td class="{'up' if display_big_changes[i] >= 0 else 'down'}">{_fmt_signed(display_big_changes[i], '%', 2)}</td>
          <td>{_fmt_float(display_small_ratios[i])}%</td>
          <td>{_fmt_optional_float(display_prices[i])}</td>
        </tr>"""
        for i, date in enumerate(display_dates)
    )
    holder_table = "".join(
        f"""
        <tr>
          <td>{_short_date(date)}</td>
          <td>{_fmt_int(display_total_holders[i])}</td>
          <td class="{'up' if display_total_changes[i] >= 0 else 'down'}">{_fmt_signed(display_total_changes[i])}</td>
          <td>{_fmt_optional_float(display_prices[i])}</td>
          <td>{_fmt_optional_lots(display_volumes[i])}</td>
        </tr>"""
        for i, date in enumerate(display_dates)
    )
    average_table = "".join(
        f"""
        <tr>
          <td>{_short_date(date)}</td>
          <td>{_fmt_float(display_avg_lots[i])}</td>
          <td class="{'up' if display_avg_lot_changes[i] >= 0 else 'down'}">{_fmt_signed(display_avg_lot_changes[i], '', 2)}</td>
          <td>{_fmt_optional_float(display_turnovers[i])}</td>
          <td>{_fmt_optional_float(display_prices[i])}</td>
        </tr>"""
        for i, date in enumerate(display_dates)
    )
    latest_total = display_total_holders[0] if display_total_holders else 0
    latest_total_change = display_total_changes[0] if display_total_changes else 0
    pyramid_data = _pyramid_dataset(rows, dates)
    pyramid_json = json.dumps(pyramid_data, ensure_ascii=False).replace("</", "<\\/")
    target_id = f"pyramid-{code}"
    quote_tone = "up" if quote["change"] >= 0 else "down"
    return f"""
    <section class="phone">
      <header class="quote">
        <div>
          <div class="price {quote_tone}">{_fmt_float(quote['close'])}</div>
          <div class="change {quote_tone}">{_fmt_signed(quote['change'], '', 2)} ({_fmt_signed(quote['pct'], '%', 2)})</div>
          <div class="stock">{_e(code)} {_e(name)}</div>
        </div>
        <div class="quote-source">
          <span>TWSE 收盤</span>
          <b>{_e(result.date)}</b>
          <small>TDCC 股權分散</small>
        </div>
      </header>
      <nav class="tabs" data-tabs>
        <button type="button" class="active" data-view="ratio">持股比</button>
        <button type="button" data-view="pyramid">金字塔</button>
        <button type="button" data-view="average">股東均張</button>
        <button type="button" data-view="holders">股東人數</button>
      </nav>

      <article class="panel" data-panel="ratio">
        <div class="panel-title"><b>大戶持股</b><span>{_short_date(latest)}</span></div>
        {_bar_chart_svg(chronological, big_ratios, price_line, "#d95749", "大戶%")}
        <table>
          <thead><tr><th>日期</th><th>大戶持股</th><th>大戶變動</th><th>散戶持股</th><th>股價</th></tr></thead>
          <tbody>{big_table}</tbody>
        </table>
      </article>

      <article class="panel" data-panel="holders" hidden>
        <div class="panel-title"><b>總股東人數</b><span>{_fmt_int(latest_total)} 人</span></div>
        {_area_chart_svg(chronological, total_holders, price_line, "人")}
        <table>
          <thead><tr><th>日期</th><th>總股東人數</th><th>人數增減</th><th>股價</th><th>成交量</th></tr></thead>
          <tbody>{holder_table}</tbody>
        </table>
      </article>

      <article class="panel" data-panel="average" hidden>
        <div class="panel-title"><b>股東均張</b><span>{_short_date(latest)}</span></div>
        {_area_chart_svg(chronological, avg_lots, price_line, "張")}
        <table>
          <thead><tr><th>日期</th><th>股東均張</th><th>均張增減</th><th>週轉率</th><th>股價</th></tr></thead>
          <tbody>{average_table}</tbody>
        </table>
      </article>

      <article class="panel pyramid-panel" data-panel="pyramid" data-pyramid-panel data-default-date="{_e(latest)}" hidden>
        <div class="panel-title"><b>金字塔</b><span>週增減: <em data-total-change class="{'up' if latest_total_change >= 0 else 'down'}">{_fmt_signed(latest_total_change)} 人</em></span></div>
        <div class="date-pills">{''.join(f'<button type="button" data-date="{_e(date)}" class="{ "active" if date == latest else "" }">{_short_date(date)}</button>' for date in dates)}</div>
        <div class="pyramid-head"><span>變動</span><span>持股人數</span><span>級距</span><span>持股比例</span><span>變動</span></div>
        <div class="pyramid" data-pyramid-body></div>
        <footer class="totals"><span>總人數: <b data-total-holders>{_fmt_int(latest_total)}</b></span><span>週增減: <em data-total-change class="{'up' if latest_total_change >= 0 else 'down'}">{_fmt_signed(latest_total_change)} 人</em></span></footer>
        <script type="application/json" id="{_e(target_id)}">{pyramid_json}</script>
      </article>
    </section>"""


def _html_page(body: str, title: str) -> str:
    return f"""<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{_e(title)}</title>
<style>
* {{ box-sizing: border-box; }}
body {{ margin: 0; background: #0b1116; color: {TEXT}; font-family: "Microsoft JhengHei", "Noto Sans TC", Arial, sans-serif; }}
.wrap {{ max-width: 1160px; margin: 0 auto; padding: 22px; display: grid; gap: 24px; }}
.page-title {{ display:flex; justify-content:space-between; align-items:end; gap:16px; }}
.page-title h1 {{ margin:0; font-size:28px; font-weight:700; }}
.page-title span {{ color:{MUTED}; }}
.phone {{ width: min(100%, 520px); margin: 0 auto; background: {BG}; border: 1px solid #26313a; border-radius: 8px; overflow: hidden; box-shadow: 0 16px 60px rgba(0,0,0,.35); }}
.quote {{ display:flex; justify-content:space-between; gap:14px; padding:14px 14px 10px; background:#020303; border-top:48px solid #101721; }}
.price {{ font-size:50px; line-height:1; font-weight:600; }}
.change {{ margin-top:8px; font-size:24px; }}
.down {{ color:{GREEN}; }}
.up {{ color:{RED}; }}
.stock {{ margin-top:8px; color:#fff077; font-size:18px; }}
.quote-source {{ min-width:150px; align-self:start; border:1px solid #345a30; padding:10px; text-align:center; }}
.quote-source span {{ display:block; font-size:18px; color:#fff; margin-bottom:8px; }}
.quote-source b {{ display:block; color:{GREEN}; font-size:20px; font-weight:500; }}
.quote-source small {{ display:block; color:{MUTED}; font-size:13px; margin-top:10px; }}
.tabs {{ display:grid; grid-template-columns:repeat(4,1fr); background:#1b2024; border-top:1px solid #26313a; border-bottom:1px solid #26313a; }}
.tabs button {{ border:0; background:transparent; text-align:center; padding:9px 2px; color:#aeb4ba; font:inherit; font-size:20px; cursor:pointer; white-space:nowrap; }}
.tabs button.active {{ color:#4ba0dd; border-bottom:3px solid #4ba0dd; }}
.panel {{ background:#050607; border-bottom:10px solid #1b2024; }}
.panel[hidden] {{ display:none; }}
.panel-title {{ display:flex; justify-content:space-between; align-items:center; padding:12px 14px 4px; font-size:20px; }}
.panel-title span {{ color:#e7d05d; }}
.custodian-filters {{ display:grid; grid-template-columns:1fr 1fr; gap:10px; padding:10px 14px 8px; font-size:20px; background:#050607; }}
.custodian-filters span {{ display:flex; align-items:center; justify-content:space-between; gap:8px; }}
.custodian-filters b {{ display:inline-block; background:#282d30; border-radius:8px; padding:8px 12px; font-weight:500; }}
.chart-svg {{ width:100%; display:block; background:#141719; }}
.grid {{ stroke:{GRID}; stroke-width:1; }}
.axis {{ fill:#f5f5f5; font-size:22px; }}
table {{ width:100%; border-collapse:collapse; table-layout:fixed; }}
th {{ color:#a5acb2; font-size:17px; font-weight:500; padding:7px 6px; background:#090b0d; white-space:nowrap; }}
td {{ padding:12px 6px; border-top:1px solid #23282d; font-size:22px; text-align:right; white-space:nowrap; }}
td:first-child, th:first-child {{ text-align:left; padding-left:12px; }}
.date-pills {{ display:flex; gap:8px; padding:8px; overflow-x:auto; scrollbar-width:thin; }}
.date-pills button {{ flex:0 0 96px; border:0; color:{TEXT}; background:#282d30; border-radius:8px; text-align:center; padding:8px 4px; font:inherit; font-size:19px; cursor:pointer; }}
.date-pills button.active {{ outline:2px solid #3389c9; }}
.pyramid-head, .pyramid-row {{ display:grid; grid-template-columns: 1fr 1.7fr 1.1fr 1.7fr 1fr; align-items:center; gap:8px; }}
.pyramid-head {{ color:#a5acb2; padding:0 10px 5px; font-size:15px; text-align:center; }}
.pyramid-row {{ min-height:31px; padding:0 10px; border-top:1px solid #15191d; font-size:20px; }}
.holders, .ratio {{ position:relative; display:block; height:26px; line-height:26px; text-align:right; padding-right:6px; }}
.holders span, .ratio span {{ position:absolute; inset:0 auto 0 0; background:{BLUE}; opacity:.88; }}
.holders b, .ratio b {{ position:relative; font-weight:500; }}
.range {{ color:#e8ecef; background:#293033; text-align:center; height:31px; line-height:31px; }}
.delta {{ font-size:17px; }}
.totals {{ display:flex; justify-content:space-between; padding:12px 12px 18px; border-top:1px solid #23282d; font-size:20px; }}
@media (min-width: 1080px) {{ .targets {{ display:grid; grid-template-columns:repeat(2, minmax(0, 1fr)); gap:24px; align-items:start; }} }}
@media (max-width: 560px) {{ .wrap {{ padding:0; }} .phone {{ border-radius:0; border-left:0; border-right:0; }} .price {{ font-size:44px; }} td {{ font-size:19px; }} }}
</style>
</head>
<body>
<main class="wrap">
  <div class="page-title"><h1>{_e(title)}</h1><span>Visual dashboard</span></div>
  <div class="targets">{body}</div>
</main>
<script>
(() => {{
  const fmtInt = (value) => Math.round(Number(value) || 0).toLocaleString("en-US");
  const fmtSignedInt = (value) => {{
    const n = Math.round(Number(value) || 0);
    return `${{n >= 0 ? "+" : ""}}${{n.toLocaleString("en-US")}} 人`;
  }};
  const fmtPercent = (value) => `${{(Number(value) || 0).toFixed(2)}}%`;
  const fmtSignedPercent = (value) => {{
    const n = Number(value) || 0;
    return `${{n >= 0 ? "+" : ""}}${{n.toFixed(2)}}%`;
  }};
  const tone = (value) => (Number(value) >= 0 ? "up" : "down");

  function renderPanel(panel, date) {{
    const dataNode = panel.querySelector('script[type="application/json"]');
    if (!dataNode) return;
    const data = JSON.parse(dataNode.textContent);
    const week = data.weeks[date] || data.weeks[data.dates[0]];
    if (!week) return;

    panel.querySelectorAll(".date-pills button").forEach((button) => {{
      button.classList.toggle("active", button.dataset.date === week.date);
    }});

    const body = panel.querySelector("[data-pyramid-body]");
    body.innerHTML = week.rows.map((row) => `
      <div class="pyramid-row">
        <span class="delta ${{tone(row.holdersChange)}}">${{fmtSignedInt(row.holdersChange)}}</span>
        <span class="holders"><span style="width:${{row.holdersWidth.toFixed(1)}}%"></span><b>${{fmtInt(row.holders)}}人</b></span>
        <span class="range">${{row.range}}</span>
        <span class="ratio"><span style="width:${{row.ratioWidth.toFixed(1)}}%"></span><b>${{fmtPercent(row.ratio)}}</b></span>
        <span class="delta ${{tone(row.ratioChange)}}">${{fmtSignedPercent(row.ratioChange)}}</span>
      </div>`).join("");

    panel.querySelectorAll("[data-total-holders]").forEach((node) => {{
      node.textContent = fmtInt(week.totalHolders);
    }});
    panel.querySelectorAll("[data-total-change]").forEach((node) => {{
      node.textContent = fmtSignedInt(week.totalHoldersChange);
      node.classList.toggle("up", week.totalHoldersChange >= 0);
      node.classList.toggle("down", week.totalHoldersChange < 0);
    }});
  }}

  document.querySelectorAll("[data-pyramid-panel]").forEach((panel) => {{
    panel.addEventListener("click", (event) => {{
      const button = event.target.closest(".date-pills button");
      if (button) renderPanel(panel, button.dataset.date);
    }});
    renderPanel(panel, panel.dataset.defaultDate);
  }});

  document.querySelectorAll(".phone").forEach((phone) => {{
    const tabs = phone.querySelector("[data-tabs]");
    if (!tabs) return;
    tabs.addEventListener("click", (event) => {{
      const button = event.target.closest("[data-view]");
      if (!button) return;
      const view = button.dataset.view;
      tabs.querySelectorAll("[data-view]").forEach((item) => {{
        item.classList.toggle("active", item === button);
      }});
      phone.querySelectorAll("[data-panel]").forEach((panel) => {{
        panel.hidden = panel.dataset.panel !== view;
      }});
    }});
  }});
}})();
</script>
</body>
</html>"""


def save_visual_outputs(
    result: BuildResult,
    shareholder_result: ShareholderBuildResult,
    outdir: Path,
    keep_latest: bool = True,
) -> dict[str, Path]:
    latest_date = shareholder_result.selected_dates[0]
    archive_dir = outdir / "archive" / latest_date
    archive_dir.mkdir(parents=True, exist_ok=True)

    body = "".join(
        _render_target(result, shareholder_result, outdir, item.code, item.name)
        for item in shareholder_result.targets
    )
    html_text = _html_page(body, f"籌碼視覺報表 {latest_date}")
    html_path = archive_dir / f"visual_dashboard_{latest_date}.html"
    html_path.write_text(html_text, encoding="utf-8")

    paths = {"html": html_path}
    if keep_latest:
        latest_path = outdir / "latest_visual_dashboard.html"
        shutil.copy2(html_path, latest_path)
        paths["latest_html"] = latest_path
    return paths
