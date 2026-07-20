"""Build a standalone Plotly chart with composable source/species filters."""

from __future__ import annotations

import html
import json
from collections import defaultdict
from datetime import timedelta
from pathlib import Path
from typing import Iterable

import plotly.graph_objects as go

from .models import MinuteCount
from .translations import tr


SOURCE_COLORS = [
    "#1976D2", "#F9A825", "#D32F2F", "#388E3C", "#7B1FA2", "#00838F",
    "#F57C00", "#5D4037", "#455A64", "#C2185B", "#512DA8", "#00796B",
]
SOURCE_SYMBOLS = [
    "circle", "triangle-up", "square", "diamond", "cross", "star",
    "triangle-down", "pentagon", "hexagon", "x", "hourglass", "bowtie",
]


def _checkboxes(values: list[str], group: str) -> str:
    return "".join(
        f'<label><input type="checkbox" data-filter="{group}" '
        f'value="{html.escape(value, quote=True)}" checked> {html.escape(value)}</label>'
        for value in values
    )


def _contrast_text_color(hex_color: str) -> str:
    red, green, blue = (int(hex_color[index:index + 2], 16) for index in (1, 3, 5))
    luminance = (0.299 * red) + (0.587 * green) + (0.114 * blue)
    return "#202124" if luminance > 155 else "#FFFFFF"


def build_chart_html(counts: Iterable[MinuteCount], title: str, language: str = "pl") -> str:
    rows = list(counts)
    if not rows:
        raise ValueError("Brak danych do przedstawienia na wykresie.")
    sources = sorted({row.source for row in rows})
    species = sorted({row.species for row in rows}, key=str.casefold)
    time_minimum = min(row.minute for row in rows) - timedelta(hours=24)
    time_maximum = max(row.minute for row in rows) + timedelta(hours=24)
    style = {
        source: (
            SOURCE_COLORS[index % len(SOURCE_COLORS)],
            SOURCE_SYMBOLS[index % len(SOURCE_SYMBOLS)],
        )
        for index, source in enumerate(sources)
    }
    grouped: dict[tuple[str, str], list[MinuteCount]] = defaultdict(list)
    for row in rows:
        grouped[(row.source, row.species)].append(row)
    coordinate_series: dict[tuple[str, object, int], list[str]] = defaultdict(list)
    for row in rows:
        coordinate_series[(row.source, row.minute, row.count)].append(row.species)

    figure = go.Figure()
    trace_filters = []
    for (source, taxon), items in sorted(grouped.items(), key=lambda item: (item[0][0], item[0][1].casefold())):
        color, symbol = style[source]
        badge_counts = []
        for item in items:
            overlapping_series = coordinate_series[(source, item.minute, item.count)]
            representative = min(overlapping_series, key=str.casefold)
            badge_counts.append(
                len(overlapping_series)
                if representative == taxon and len(overlapping_series) > 1
                else 0
            )
        figure.add_trace(go.Scatter(
            x=[item.minute for item in items],
            y=[item.count for item in items],
            mode="markers+text",
            name=f"{source} · {taxon}",
            marker={
                "color": color,
                "symbol": symbol,
                "size": [26 if count > 1 else 14 for count in badge_counts],
                "line": {"width": 1, "color": "#202124"},
            },
            text=[str(count) if count > 1 else "" for count in badge_counts],
            textposition="middle center",
            textfont={"color": _contrast_text_color(color), "size": 13, "weight": 700},
            customdata=[[source, taxon] for _ in items],
            hovertemplate=(
                "<b>%{customdata[0]}</b> · %{customdata[1]}: %{y}<extra></extra>"
            ),
        ))
        trace_filters.append({"source": source, "species": taxon})

    figure.update_layout(
        title=title,
        template="plotly_white",
        height=760,
        hovermode="x unified",
        xaxis={
            "title": tr(language, "x_axis"),
            "minallowed": time_minimum,
            "maxallowed": time_maximum,
            "rangeslider": {
                "visible": True,
                "range": [time_minimum, time_maximum],
            },
            "showspikes": True,
            "spikemode": "across",
            "spikesnap": "cursor",
        },
        yaxis={"title": tr(language, "y_axis"), "rangemode": "tozero", "dtick": 1},
        legend={"title": {"text": tr(language, "legend")}, "groupclick": "toggleitem"},
        margin={"l": 70, "r": 30, "t": 70, "b": 80},
    )
    chart = figure.to_html(
        full_html=False,
        include_plotlyjs=True,
        div_id="parallel-bat-chart",
        config={"responsive": True, "displaylogo": False, "scrollZoom": True},
    )
    source_boxes = _checkboxes(sources, "source")
    species_boxes = _checkboxes(species, "species")
    filters_json = json.dumps(trace_filters, ensure_ascii=False).replace("</", "<\\/")
    table_rows_json = json.dumps([
        {
            "minute": row.minute.isoformat(timespec="seconds"),
            "source": row.source,
            "species": row.species,
            "count": row.count,
        }
        for row in rows
    ], ensure_ascii=False).replace("</", "<\\/")
    table_labels_json = json.dumps({
        "visibleRows": tr(language, "visible_rows", count="{count}"),
        "selectionActive": tr(language, "selection_active"),
        "clearSelection": tr(language, "clear_selection"),
        "noData": tr(language, "no_visible_data"),
    }, ensure_ascii=False).replace("</", "<\\/")
    safe_title = html.escape(title)
    filters_aria = html.escape(tr(language, "filters_aria"), quote=True)
    sources_label = html.escape(tr(language, "sources"))
    species_label = html.escape(tr(language, "species"))
    all_label = html.escape(tr(language, "all"))
    none_label = html.escape(tr(language, "none"))
    print_label = html.escape(tr(language, "print_png"))
    table_title = html.escape(tr(language, "table_title"))
    table_minute = html.escape(tr(language, "table_minute"))
    table_source = html.escape(tr(language, "table_source"))
    table_manual_id = html.escape(tr(language, "table_manual_id"))
    table_count = html.escape(tr(language, "table_count"))
    return f"""<!doctype html>
<html lang="{html.escape(language, quote=True)}">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{safe_title}</title>
  <style>
    :root {{ font-family: system-ui, -apple-system, "Segoe UI", sans-serif; color: #202124; }}
    body {{ margin: 0; background: #f6f8fa; }}
    header {{ display: flex; align-items: center; justify-content: space-between; gap: 16px; padding: 16px 22px 4px; }}
    header h1 {{ margin: 0; }}
    .export-button {{ border: 1px solid #0969da; border-radius: 6px; background: #0969da; color: white; font-weight: 650; padding: 8px 14px; }}
    .export-button:hover {{ background: #0757b5; }}
    .filters {{ display: grid; grid-template-columns: minmax(220px, 1fr) 2fr; gap: 14px; padding: 10px 22px; }}
    fieldset {{ background: white; border: 1px solid #d0d7de; border-radius: 8px; max-height: 170px; overflow: auto; }}
    legend {{ font-weight: 650; }}
    label {{ display: inline-block; margin: 4px 12px 4px 0; white-space: nowrap; }}
    button {{ margin: 3px 6px 5px 0; padding: 4px 9px; cursor: pointer; }}
    .chart-wrap {{ min-height: 760px; margin: 8px 22px 22px; background: white; border: 1px solid #d0d7de; border-radius: 8px; }}
    .table-section {{ margin: 0 22px 28px; padding: 16px; background: white; border: 1px solid #d0d7de; border-radius: 8px; }}
    .table-heading {{ display: flex; align-items: baseline; justify-content: space-between; gap: 12px; margin-bottom: 10px; }}
    .table-heading h2 {{ margin: 0; font-size: 1.2rem; }}
    #visible-row-count {{ color: #57606a; }}
    .table-scroll {{ max-height: 480px; overflow: auto; border: 1px solid #d8dee4; border-radius: 6px; }}
    table {{ width: 100%; border-collapse: collapse; font-variant-numeric: tabular-nums; }}
    th, td {{ padding: 7px 10px; border-bottom: 1px solid #d8dee4; text-align: left; }}
    th {{ position: sticky; top: 0; z-index: 1; background: #f6f8fa; }}
    td:last-child, th:last-child {{ text-align: right; }}
    tbody tr:nth-child(even) {{ background: #fbfcfd; }}
    .no-data {{ padding: 20px; text-align: center !important; color: #57606a; }}
    @media (max-width: 800px) {{ .filters {{ grid-template-columns: 1fr; }} }}
  </style>
</head>
<body>
  <header>
    <h1>{safe_title}</h1>
    <button id="export-png" class="export-button" type="button">{print_label}</button>
  </header>
  <section class="filters" aria-label="{filters_aria}">
    <fieldset><legend>{sources_label}</legend>
      <button type="button" data-action="all" data-group="source">{all_label}</button>
      <button type="button" data-action="none" data-group="source">{none_label}</button><br>{source_boxes}
    </fieldset>
    <fieldset><legend>{species_label}</legend>
      <button type="button" data-action="all" data-group="species">{all_label}</button>
      <button type="button" data-action="none" data-group="species">{none_label}</button><br>{species_boxes}
    </fieldset>
  </section>
  <main class="chart-wrap">{chart}</main>
  <section class="table-section">
    <div class="table-heading">
      <h2>{table_title}</h2>
      <span id="visible-row-count"></span>
    </div>
    <div class="table-scroll">
      <table>
        <thead><tr><th>{table_minute}</th><th>{table_source}</th><th>{table_manual_id}</th><th>{table_count}</th></tr></thead>
        <tbody id="visible-data-body"></tbody>
      </table>
    </div>
  </section>
  <script>
    const traceFilters = {filters_json};
    const tableRows = {table_rows_json};
    const tableLabels = {table_labels_json};
    const chartElement = document.getElementById('parallel-bat-chart');
    let selectedTableKeys = null;
    let clearingSelection = false;
    let updatingMarkerBadges = false;
    function selected(group) {{
      return new Set(Array.from(document.querySelectorAll(`input[data-filter="${{group}}"]:checked`)).map(x => x.value));
    }}
    function seriesKey(source, species) {{
      return JSON.stringify([source, species]);
    }}
    function tableKey(minute, source, species) {{
      return JSON.stringify([minute, source, species]);
    }}
    function visibleSeries() {{
      const result = new Set();
      traceFilters.forEach((item, index) => {{
        const visibility = chartElement.data[index].visible;
        if (visibility !== false && visibility !== 'legendonly') {{
          result.add(seriesKey(item.source, item.species));
        }}
      }});
      return result;
    }}
    function visibleTimeRange() {{
      const axis = chartElement._fullLayout && chartElement._fullLayout.xaxis;
      if (!axis || !axis.range) return [null, null];
      return [new Date(axis.range[0]).getTime(), new Date(axis.range[1]).getTime()];
    }}
    function updateMarkerBadges() {{
      if (updatingMarkerBadges) return Promise.resolve();
      updatingMarkerBadges = true;
      const summaries = new Map();
      const texts = chartElement.data.map(trace => trace.x.map(() => ''));
      const sizes = chartElement.data.map(trace => trace.x.map(() => 14));
      traceFilters.forEach((filter, curveNumber) => {{
        const trace = chartElement.data[curveNumber];
        if (trace.visible === false || trace.visible === 'legendonly') return;
        trace.x.forEach((minute, pointNumber) => {{
          const count = Number(trace.y[pointNumber]);
          const key = JSON.stringify([filter.source, String(minute), count]);
          const current = summaries.get(key);
          if (!current) {{
            summaries.set(key, {{seriesCount: 1, curveNumber, pointNumber, count}});
          }} else {{
            current.seriesCount += 1;
            if (curveNumber < current.curveNumber) {{
              current.curveNumber = curveNumber;
              current.pointNumber = pointNumber;
            }}
          }}
        }});
      }});
      summaries.forEach(summary => {{
        const badge = summary.seriesCount;
        if (badge > 1) {{
          texts[summary.curveNumber][summary.pointNumber] = String(badge);
          sizes[summary.curveNumber][summary.pointNumber] = 26;
        }}
      }});
      const updates = chartElement.data.map((_trace, curveNumber) =>
        Plotly.restyle(chartElement, {{
          text: [texts[curveNumber]],
          'marker.size': [sizes[curveNumber]]
        }}, [curveNumber])
      );
      return Promise.all(updates).finally(() => {{ updatingMarkerBadges = false; }});
    }}
    function appendCell(row, value, className = '') {{
      const cell = document.createElement('td');
      cell.textContent = value;
      if (className) cell.className = className;
      row.appendChild(cell);
    }}
    function renderTable() {{
      const body = document.getElementById('visible-data-body');
      const activeSeries = visibleSeries();
      const [rangeStart, rangeEnd] = visibleTimeRange();
      const visibleRows = tableRows.filter(row => {{
        const timestamp = new Date(row.minute).getTime();
        const inRange = (rangeStart === null || timestamp >= rangeStart) &&
          (rangeEnd === null || timestamp <= rangeEnd);
        const inSelection = selectedTableKeys === null ||
          selectedTableKeys.has(tableKey(row.minute, row.source, row.species));
        return inRange && inSelection && activeSeries.has(seriesKey(row.source, row.species));
      }});
      body.replaceChildren();
      if (visibleRows.length === 0) {{
        const row = document.createElement('tr');
        const cell = document.createElement('td');
        cell.colSpan = 4;
        cell.className = 'no-data';
        cell.textContent = tableLabels.noData;
        row.appendChild(cell);
        body.appendChild(row);
      }} else {{
        const fragment = document.createDocumentFragment();
        visibleRows.forEach(item => {{
          const row = document.createElement('tr');
          appendCell(row, item.minute.slice(0, 10) + ' ' + item.minute.slice(11, 16));
          appendCell(row, item.source);
          appendCell(row, item.species);
          appendCell(row, String(item.count));
          fragment.appendChild(row);
        }});
        body.appendChild(fragment);
      }}
      let rowCount = tableLabels.visibleRows.replace('{{count}}', String(visibleRows.length));
      if (selectedTableKeys !== null) rowCount += ' · ' + tableLabels.selectionActive;
      document.getElementById('visible-row-count').textContent = rowCount;
    }}
    function applyPointSelection(eventData) {{
      if (clearingSelection) return;
      if (!eventData || !Array.isArray(eventData.points)) {{
        selectedTableKeys = null;
        renderTable();
        return;
      }}
      selectedTableKeys = new Set(eventData.points.map(point => {{
        const filter = traceFilters[point.curveNumber];
        const minute = chartElement.data[point.curveNumber].x[point.pointNumber];
        return tableKey(String(minute), filter.source, filter.species);
      }}));
      renderTable();
    }}
    function clearPointSelection() {{
      clearingSelection = true;
      Plotly.relayout(chartElement, {{selections: []}}).then(() => {{
        chartElement.data.forEach(trace => {{ trace.selectedpoints = null; }});
        return Plotly.redraw(chartElement);
      }}).finally(() => {{
        selectedTableKeys = null;
        clearingSelection = false;
        renderTable();
      }});
    }}
    function addClearSelectionModebarButton() {{
      const modebar = chartElement.querySelector('.modebar');
      if (!modebar || modebar.querySelector('[data-clear-selection]')) return;
      const group = document.createElement('div');
      group.className = 'modebar-group';
      group.setAttribute('data-clear-selection-group', 'true');
      const button = document.createElement('a');
      button.className = 'modebar-btn';
      button.rel = 'tooltip';
      button.setAttribute('data-title', tableLabels.clearSelection);
      button.setAttribute('aria-label', tableLabels.clearSelection);
      button.setAttribute('data-clear-selection', 'true');
      button.innerHTML = '<svg viewBox="0 0 1000 1000" class="icon" height="1em" width="1em"><path d="M185 120L500 435 815 120 880 185 565 500 880 815 815 880 500 565 185 880 120 815 435 500 120 185Z"></path></svg>';
      button.addEventListener('click', clearPointSelection);
      group.appendChild(button);
      modebar.appendChild(group);
    }}
    function applyFilters() {{
      const activeSources = selected('source');
      const activeSpecies = selected('species');
      const visible = traceFilters.map(item => activeSources.has(item.source) && activeSpecies.has(item.species));
      Plotly.restyle(chartElement, {{visible: visible}})
        .then(updateMarkerBadges)
        .then(renderTable);
    }}
    document.querySelectorAll('input[data-filter]').forEach(input => input.addEventListener('change', applyFilters));
    document.querySelectorAll('button[data-action]').forEach(button => button.addEventListener('click', () => {{
      const checked = button.dataset.action === 'all';
      document.querySelectorAll(`input[data-filter="${{button.dataset.group}}"]`).forEach(input => input.checked = checked);
      applyFilters();
    }}));
    document.getElementById('export-png').addEventListener('click', () => {{
      Plotly.downloadImage(chartElement, {{
        format: 'png',
        filename: 'parallel_bat_activity_current_view',
        width: Math.max(chartElement.clientWidth, 1200),
        height: 760,
        scale: 2
      }});
    }});
    function initialiseTable() {{
      if (typeof chartElement.on !== 'function' || !chartElement._fullLayout) {{
        window.requestAnimationFrame(initialiseTable);
        return;
      }}
      chartElement.on('plotly_relayout', renderTable);
      chartElement.on('plotly_restyle', renderTable);
      chartElement.on('plotly_legendclick', () => setTimeout(() => {{
        updateMarkerBadges().then(renderTable);
      }}, 0));
      chartElement.on('plotly_legenddoubleclick', () => setTimeout(() => {{
        updateMarkerBadges().then(renderTable);
      }}, 0));
      chartElement.on('plotly_afterplot', addClearSelectionModebarButton);
      chartElement.on('plotly_selected', applyPointSelection);
      chartElement.on('plotly_deselect', () => {{
        selectedTableKeys = null;
        renderTable();
      }});
      addClearSelectionModebarButton();
      updateMarkerBadges().then(renderTable);
    }}
    initialiseTable();
  </script>
</body>
</html>"""


def write_chart(path: Path, counts: Iterable[MinuteCount], title: str, language: str = "pl") -> Path:
    output = Path(path)
    output.write_text(build_chart_html(counts, title, language), encoding="utf-8")
    return output
