#!/usr/bin/env python3
"""
Local drag-and-drop LSV conversion app.

Run:
  python scripts/lsv_tool_app.py

Then open:
  http://127.0.0.1:8765
"""

from __future__ import annotations

import cgi
import csv
import io
import json
import argparse
import math
import time
import webbrowser
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any

from extract_lsv_vfim import CurveMeta, extract_vfim, sheet_title, write_xlsx


HOST = "127.0.0.1"
PORT = 8765
WORKSPACE_TMP = Path.cwd() / "lsv_vfim_output" / "_tmp"


HTML = r"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>LSV 数据转换工具</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #f6f7f9;
      --panel: #ffffff;
      --text: #1c2430;
      --muted: #647083;
      --line: #d8dde6;
      --accent: #1d7f74;
      --accent-dark: #12665d;
      --soft: #e9f5f2;
      --danger: #b3261e;
      --shadow: 0 18px 50px rgba(30, 38, 50, 0.10);
    }

    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      font-family: "Microsoft YaHei", "Segoe UI", Arial, sans-serif;
      background: var(--bg);
      color: var(--text);
    }

    main {
      width: min(1180px, calc(100vw - 40px));
      margin: 0 auto;
      padding: 34px 0 44px;
    }

    header {
      display: flex;
      align-items: flex-end;
      justify-content: space-between;
      gap: 20px;
      margin-bottom: 22px;
    }

    h1 {
      margin: 0 0 8px;
      font-size: 30px;
      line-height: 1.2;
      font-weight: 760;
      letter-spacing: 0;
    }

    .sub {
      margin: 0;
      color: var(--muted);
      font-size: 14px;
      line-height: 1.6;
    }

    .status {
      padding: 10px 14px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--panel);
      color: var(--muted);
      font-size: 13px;
      white-space: nowrap;
    }

    .drop {
      display: grid;
      place-items: center;
      min-height: 172px;
      border: 2px dashed #9bb7b2;
      border-radius: 8px;
      background: linear-gradient(180deg, #ffffff 0%, #f0faf7 100%);
      box-shadow: var(--shadow);
      margin-bottom: 18px;
      transition: border-color 140ms ease, transform 140ms ease, background 140ms ease;
    }

    .drop.drag {
      border-color: var(--accent);
      transform: translateY(-1px);
      background: #e9f8f4;
    }

    .drop-inner {
      text-align: center;
      padding: 20px;
    }

    .drop-title {
      font-size: 18px;
      font-weight: 700;
      margin-bottom: 10px;
    }

    .drop-text {
      color: var(--muted);
      font-size: 13px;
      margin-bottom: 18px;
    }

    button, .file-button {
      border: 0;
      border-radius: 8px;
      background: var(--accent);
      color: #fff;
      min-height: 38px;
      padding: 0 16px;
      font-size: 14px;
      font-weight: 650;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
    }

    button:hover, .file-button:hover { background: var(--accent-dark); }
    button.secondary {
      background: #edf1f5;
      color: var(--text);
      border: 1px solid var(--line);
    }
    button.secondary:hover { background: #e1e7ed; }
    button.danger {
      background: #fff2f1;
      color: var(--danger);
      border: 1px solid #f1c5c1;
    }
    button.danger:hover { background: #ffe5e2; }
    button:disabled {
      opacity: .55;
      cursor: not-allowed;
    }

    input[type="file"] { display: none; }

    .toolbar {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 16px;
      align-items: end;
      margin: 18px 0;
    }

    .defaults {
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      align-items: end;
    }

    label {
      display: grid;
      gap: 6px;
      color: var(--muted);
      font-size: 12px;
      font-weight: 650;
    }

    input {
      width: 116px;
      height: 38px;
      border-radius: 8px;
      border: 1px solid var(--line);
      background: #fff;
      padding: 0 10px;
      color: var(--text);
      font-size: 14px;
      outline: none;
    }

    input:focus {
      border-color: var(--accent);
      box-shadow: 0 0 0 3px rgba(29, 127, 116, .14);
    }

    .table-wrap {
      overflow: auto;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--panel);
      box-shadow: var(--shadow);
    }

    table {
      width: 100%;
      border-collapse: collapse;
      min-width: 780px;
    }

    th, td {
      border-bottom: 1px solid var(--line);
      padding: 10px 12px;
      text-align: left;
      font-size: 13px;
      vertical-align: middle;
    }

    th {
      position: sticky;
      top: 0;
      background: #f9fbfc;
      color: #3f4b5a;
      font-size: 12px;
      letter-spacing: 0;
      z-index: 1;
    }

    tr:last-child td { border-bottom: 0; }
    td.name {
      max-width: 360px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
      font-weight: 650;
    }

    td.size { color: var(--muted); width: 92px; }
    td.actions { width: 82px; text-align: right; }
    td input { width: 100%; min-width: 92px; }
    td.sheet input { min-width: 160px; }

    .empty {
      text-align: center;
      padding: 36px 14px;
      color: var(--muted);
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
    }

    .footer {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 14px;
      margin-top: 18px;
    }

    .note {
      color: var(--muted);
      font-size: 13px;
      line-height: 1.5;
    }

    .error {
      color: var(--danger);
      font-size: 13px;
      margin-top: 10px;
      min-height: 20px;
    }

    @media (max-width: 760px) {
      main { width: min(100vw - 24px, 1180px); padding-top: 22px; }
      header, .toolbar, .footer { display: grid; }
      .status { white-space: normal; }
      .defaults { display: grid; grid-template-columns: 1fr 1fr; }
      input { width: 100%; }
      h1 { font-size: 24px; }
    }
  </style>
</head>
<body>
  <main>
    <header>
      <div>
        <h1>LSV 数据转换工具</h1>
        <p class="sub">拖入原始 .DTA 文件，填写每条曲线的 pH 和面积，导出完整处理表和 Origin 作图表。</p>
      </div>
      <div class="status" id="status">本地运行，文件不会上传到外网</div>
    </header>

    <section class="drop" id="drop">
      <div class="drop-inner">
        <div class="drop-title">把 LSV 原始文件拖到这里</div>
        <div class="drop-text">支持多个文件；也可以点击按钮选择。文件名相同会自动覆盖旧项。</div>
        <label class="file-button" for="fileInput">选择文件</label>
        <input id="fileInput" type="file" multiple />
      </div>
    </section>

    <section class="toolbar">
      <div class="defaults">
        <label>Tafel min j
          <input id="tafelMinJ" type="number" step="0.1" value="10" />
        </label>
        <label>Tafel max j
          <input id="tafelMaxJ" type="number" step="0.1" value="100" />
        </label>
        <label>默认 pH
          <input id="defaultPh" type="number" step="0.001" placeholder="例如 8.41" />
        </label>
        <label>默认面积 cm²
          <input id="defaultArea" type="number" step="0.0001" placeholder="例如 0.25" />
        </label>
        <button class="secondary" id="applyDefaults" type="button">应用到空白项</button>
      </div>
      <button class="danger" id="clearFiles" type="button">清空</button>
    </section>

    <div id="fileArea" class="empty">还没有文件。拖入或选择 `.DTA.###` 文件后，会在这里填写每条曲线的参数。</div>

    <div class="footer">
      <div class="note">导出的完整 Excel 包含 A-E 五列；Origin 表使用每条曲线的 C/E 列；同时生成黑底 LSV 曲线图 PNG。</div>
      <button id="exportBtn" type="button" disabled>导出结果</button>
    </div>
    <div class="error" id="error"></div>
  </main>

  <script>
    const state = new Map();
    const drop = document.getElementById('drop');
    const input = document.getElementById('fileInput');
    const fileArea = document.getElementById('fileArea');
    const exportBtn = document.getElementById('exportBtn');
    const errorBox = document.getElementById('error');
    const statusBox = document.getElementById('status');

    function safeSheetName(name) {
      return name.replace(/\.[^.]+$/g, '').replace(/[\[\]:*?/\\]/g, '_').slice(0, 31) || 'LSV';
    }

    function formatSize(size) {
      if (size < 1024) return `${size} B`;
      if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
      return `${(size / 1024 / 1024).toFixed(2)} MB`;
    }

    function addFiles(files) {
      for (const file of files) {
        state.set(file.name, {
          file,
          ph: document.getElementById('defaultPh').value || '',
          area: document.getElementById('defaultArea').value || '',
          sheetName: safeSheetName(file.name)
        });
      }
      render();
    }

    function render() {
      errorBox.textContent = '';
      exportBtn.disabled = state.size === 0;
      statusBox.textContent = state.size ? `已加入 ${state.size} 个文件` : '本地运行，文件不会上传到外网';
      if (!state.size) {
        fileArea.className = 'empty';
        fileArea.textContent = '还没有文件。拖入或选择 `.DTA.###` 文件后，会在这里填写每条曲线的参数。';
        return;
      }

      fileArea.className = 'table-wrap';
      const rows = [...state.values()].map((item) => `
        <tr>
          <td class="name" title="${item.file.name}">${item.file.name}</td>
          <td class="size">${formatSize(item.file.size)}</td>
          <td><input data-field="ph" data-name="${item.file.name}" type="number" step="0.001" value="${item.ph}" placeholder="pH"></td>
          <td><input data-field="area" data-name="${item.file.name}" type="number" step="0.0001" value="${item.area}" placeholder="cm²"></td>
          <td class="sheet"><input data-field="sheetName" data-name="${item.file.name}" value="${item.sheetName}" placeholder="sheet 名"></td>
          <td class="actions"><button class="danger" data-remove="${item.file.name}" type="button">删除</button></td>
        </tr>
      `).join('');

      fileArea.innerHTML = `
        <table>
          <thead>
            <tr>
              <th>文件</th>
              <th>大小</th>
              <th>pH</th>
              <th>面积 cm²</th>
              <th>Sheet 名</th>
              <th></th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      `;

      fileArea.querySelectorAll('input').forEach((el) => {
        el.addEventListener('input', () => {
          const item = state.get(el.dataset.name);
          item[el.dataset.field] = el.value;
        });
      });
      fileArea.querySelectorAll('button[data-remove]').forEach((btn) => {
        btn.addEventListener('click', () => {
          state.delete(btn.dataset.remove);
          render();
        });
      });
    }

    ['dragenter', 'dragover'].forEach((eventName) => {
      drop.addEventListener(eventName, (event) => {
        event.preventDefault();
        drop.classList.add('drag');
      });
    });
    ['dragleave', 'drop'].forEach((eventName) => {
      drop.addEventListener(eventName, (event) => {
        event.preventDefault();
        drop.classList.remove('drag');
      });
    });
    drop.addEventListener('drop', (event) => addFiles(event.dataTransfer.files));
    input.addEventListener('change', () => addFiles(input.files));

    document.getElementById('applyDefaults').addEventListener('click', () => {
      const ph = document.getElementById('defaultPh').value;
      const area = document.getElementById('defaultArea').value;
      for (const item of state.values()) {
        if (!item.ph && ph) item.ph = ph;
        if (!item.area && area) item.area = area;
      }
      render();
    });

    document.getElementById('clearFiles').addEventListener('click', () => {
      state.clear();
      input.value = '';
      render();
    });

    exportBtn.addEventListener('click', async () => {
      errorBox.textContent = '';
      const items = [...state.values()];
      for (const item of items) {
        if (!item.ph || !item.area) {
          errorBox.textContent = '请先为每个文件填写 pH 和面积。';
          return;
        }
      }

      const form = new FormData();
      const tafelMinJ = Number(document.getElementById('tafelMinJ').value || 10);
      const tafelMaxJ = Number(document.getElementById('tafelMaxJ').value || 100);
      if (!(tafelMinJ > 0) || !(tafelMaxJ > tafelMinJ)) {
        errorBox.textContent = 'Tafel range must satisfy: 0 < min j < max j.';
        return;
      }
      const manifest = items.map((item) => ({
        name: item.file.name,
        pH: Number(item.ph),
        area_cm2: Number(item.area),
        sheet_name: item.sheetName || safeSheetName(item.file.name)
      }));
      form.append('manifest', JSON.stringify(manifest));
      form.append('tafel_min_j', String(tafelMinJ));
      form.append('tafel_max_j', String(tafelMaxJ));
      for (const item of items) form.append('files', item.file, item.file.name);

      exportBtn.disabled = true;
      exportBtn.textContent = '正在生成...';
      try {
        const response = await fetch('/api/export', { method: 'POST', body: form });
        if (!response.ok) {
          const text = await response.text();
          throw new Error(text || `HTTP ${response.status}`);
        }
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `LSV_export_${new Date().toISOString().slice(0, 19).replace(/[:T]/g, '')}.zip`;
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
        statusBox.textContent = '导出完成';
      } catch (error) {
        errorBox.textContent = error.message || String(error);
      } finally {
        exportBtn.disabled = state.size === 0;
        exportBtn.textContent = '导出结果';
      }
    });
  </script>
</body>
</html>
"""


def build_origin_tables(
    curves: list[tuple[Path, list[Any], CurveMeta]],
    resistance_ohm: float = 3.8504,
    rhe_offset_v: float = 0.197,
    ph_slope_v: float = 0.0591,
) -> tuple[bytes, bytes]:
    from openpyxl import Workbook
    from openpyxl.styles import Font

    used_titles: set[str] = set()
    rows_by_curve: list[tuple[str, list[tuple[float, float]]]] = []
    max_rows = 0

    for source, points, meta in curves:
        title = sheet_title(meta.sheet_name or source.name, used_titles)
        rhe_shift = rhe_offset_v + ph_slope_v * meta.ph
        rows = [
            (
                point.vf_v + rhe_shift,
                point.im_a * 1000 / meta.area_cm2,
            )
            for point in points
        ]
        rows_by_curve.append((title, rows))
        max_rows = max(max_rows, len(rows))

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Origin_CE"

    csv_buffer = io.StringIO(newline="")
    csv_writer = csv.writer(csv_buffer)

    header: list[str] = []
    for title, _rows in rows_by_curve:
        header.extend([f"{title}_v_V_vs_RHE", f"{title}_current_density_mA_cm2"])
    sheet.append(header)
    csv_writer.writerow(header)
    for cell in sheet[1]:
        cell.font = Font(bold=True)

    for row_index in range(max_rows):
        row: list[float | str] = []
        for _title, rows in rows_by_curve:
            if row_index < len(rows):
                row.extend(rows[row_index])
            else:
                row.extend(["", ""])
        sheet.append(row)
        csv_writer.writerow(row)

    for col_index in range(1, len(header) + 1):
        sheet.column_dimensions[sheet.cell(row=1, column=col_index).column_letter].width = 24

    xlsx_buffer = io.BytesIO()
    workbook.save(xlsx_buffer)
    return xlsx_buffer.getvalue(), csv_buffer.getvalue().encode("utf-8-sig")


def nice_number(value: float, round_value: bool) -> float:
    if value <= 0:
        return 1.0
    exponent = math.floor(math.log10(value))
    fraction = value / 10**exponent
    if round_value:
        if fraction < 1.5:
            nice_fraction = 1
        elif fraction < 3:
            nice_fraction = 2
        elif fraction < 7:
            nice_fraction = 5
        else:
            nice_fraction = 10
    else:
        if fraction <= 1:
            nice_fraction = 1
        elif fraction <= 2:
            nice_fraction = 2
        elif fraction <= 5:
            nice_fraction = 5
        else:
            nice_fraction = 10
    return nice_fraction * 10**exponent


def nice_ticks(min_value: float, max_value: float, target_count: int = 5) -> list[float]:
    if math.isclose(min_value, max_value):
        span = abs(max_value) or 1
        min_value -= span * 0.5
        max_value += span * 0.5
    span = nice_number(max_value - min_value, False)
    step = nice_number(span / max(target_count - 1, 1), True)
    tick_min = math.floor(min_value / step) * step
    tick_max = math.ceil(max_value / step) * step
    ticks = []
    value = tick_min
    guard = 0
    while value <= tick_max + step * 0.5 and guard < 100:
        ticks.append(round(value, 10))
        value += step
        guard += 1
    return ticks


def format_tick(value: float) -> str:
    if abs(value) >= 100:
        return f"{value:.0f}"
    if abs(value) >= 10:
        return f"{value:.1f}".rstrip("0").rstrip(".")
    return f"{value:.2f}".rstrip("0").rstrip(".")


def find_font(size: int):
    from PIL import ImageFont

    candidates = [
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simhei.ttf",
        "C:/Windows/Fonts/simsun.ttc",
        "C:/Windows/Fonts/arial.ttf",
    ]
    for candidate in candidates:
        if Path(candidate).exists():
            return ImageFont.truetype(candidate, size=size)
    return ImageFont.load_default()


def draw_text_centered(draw: Any, position: tuple[float, float], text: str, font: Any, fill: str) -> None:
    bbox = draw.textbbox((0, 0), text, font=font)
    width = bbox[2] - bbox[0]
    height = bbox[3] - bbox[1]
    draw.text((position[0] - width / 2, position[1] - height / 2), text, font=font, fill=fill)


def draw_rotated_label(image: Any, text: str, center: tuple[int, int], font: Any, fill: str) -> None:
    from PIL import Image, ImageDraw

    bbox = font.getbbox(text)
    width = bbox[2] - bbox[0] + 16
    height = bbox[3] - bbox[1] + 16
    label = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    label_draw = ImageDraw.Draw(label)
    label_draw.text((8 - bbox[0], 8 - bbox[1]), text, font=font, fill=fill)
    rotated = label.rotate(90, expand=True)
    image.alpha_composite(rotated, (int(center[0] - rotated.width / 2), int(center[1] - rotated.height / 2)))


def build_lsv_plot_png(
    curves: list[tuple[Path, list[Any], CurveMeta]],
    resistance_ohm: float = 3.8504,
    rhe_offset_v: float = 0.197,
    ph_slope_v: float = 0.0591,
) -> bytes:
    from PIL import Image, ImageDraw

    palette = ["#6a6a6a", "#ff4d4d", "#2f80ed", "#43bf79", "#b46ae5", "#e2b500", "#00a6a6", "#f28e2b"]
    plot_curves: list[tuple[str, str, list[tuple[float, float]]]] = []
    all_x: list[float] = []
    all_y: list[float] = []
    used_titles: set[str] = set()

    for index, (source, points, meta) in enumerate(curves):
        title = sheet_title(meta.sheet_name or source.name, used_titles)
        rhe_shift = rhe_offset_v + ph_slope_v * meta.ph
        rows = [
            (
                point.vf_v + rhe_shift,
                point.im_a * 1000 / meta.area_cm2,
            )
            for point in points
        ]
        rows = [(x, y) for x, y in rows if math.isfinite(x) and math.isfinite(y)]
        if not rows:
            continue
        plot_curves.append((title, palette[index % len(palette)], rows))
        all_x.extend(x for x, _y in rows)
        all_y.extend(y for _x, y in rows)

    if not plot_curves:
        raise ValueError("没有可绘制的曲线数据。")

    width, height = 1800, 1400
    left, right, top, bottom = 290, 120, 150, 190
    plot_left, plot_top = left, top
    plot_right, plot_bottom = width - right, height - bottom
    plot_width = plot_right - plot_left
    plot_height = plot_bottom - plot_top

    # Match the user's reference plotting window.
    x_min, x_max = 1.0, 1.8
    y_min, y_max = 0.0, 250.0
    x_ticks = [1.0, 1.2, 1.4, 1.6, 1.8]
    y_ticks = [0.0, 100.0, 200.0]

    def x_to_px(x_value: float) -> float:
        return plot_left + (x_value - x_min) / (x_max - x_min) * plot_width

    def y_to_px(y_value: float) -> float:
        return plot_bottom - (y_value - y_min) / (y_max - y_min) * plot_height

    image = Image.new("RGBA", (width, height), "#ffffff")
    draw = ImageDraw.Draw(image)
    font_tick = find_font(54)
    font_label = find_font(68)
    font_legend = find_font(40)
    font_small = find_font(34)

    axis_color = "#000000"
    tick_color = "#000000"
    label_color = "#000000"
    draw.line((plot_left, plot_bottom, plot_right, plot_bottom), fill=axis_color, width=6)
    draw.line((plot_left, plot_top, plot_left, plot_bottom), fill=axis_color, width=6)

    for tick in x_ticks:
        x = x_to_px(tick)
        draw.line((x, plot_bottom, x, plot_bottom + 12), fill=tick_color, width=5)
        draw_text_centered(draw, (x, plot_bottom + 62), format_tick(tick), font_tick, label_color)

    for tick in y_ticks:
        y = y_to_px(tick)
        draw.line((plot_left - 12, y, plot_left, y), fill=tick_color, width=5)
        bbox = draw.textbbox((0, 0), format_tick(tick), font=font_tick)
        draw.text((plot_left - 32 - (bbox[2] - bbox[0]), y - (bbox[3] - bbox[1]) / 2), format_tick(tick), font=font_tick, fill=label_color)

    draw_text_centered(draw, ((plot_left + plot_right) / 2, height - 68), "电位/(V vs RHE)", font_label, label_color)
    draw_rotated_label(image, "j/(mA cm-2)", (76, (plot_top + plot_bottom) // 2), font_label, label_color)

    for title, color, rows in plot_curves:
        visible_rows = [(x, y) for x, y in rows if x_min <= x <= x_max and y_min <= y <= y_max]
        points = [(x_to_px(x), y_to_px(y)) for x, y in visible_rows]
        if len(points) >= 2:
            draw.line(points, fill=color, width=12, joint="curve")

    legend_x, legend_y = 340, 210
    line_len = 130
    row_gap = 64
    for index, (title, color, _rows) in enumerate(plot_curves):
        y = legend_y + index * row_gap
        draw.line((legend_x, y, legend_x + line_len, y), fill=color, width=13)
        draw.text((legend_x + line_len + 20, y - 25), title, font=font_legend, fill=label_color)

    note_lines = [
        "工作电极：自动导出曲线",
        "对电极：碳棒电极",
        "参比电极：Ag/AgCl",
    ]
    note_y = legend_y + len(plot_curves) * row_gap + 30
    for line in note_lines:
        draw.text((legend_x, note_y), line, font=font_small, fill=label_color)
        note_y += 50

    output = io.BytesIO()
    image.convert("RGB").save(output, format="PNG", optimize=True)
    return output.getvalue()


def linear_regression(points: list[tuple[float, float]]) -> tuple[float, float, float]:
    if len(points) < 2:
        raise ValueError("At least two points are required for linear regression.")
    n = len(points)
    sum_x = sum(x for x, _y in points)
    sum_y = sum(y for _x, y in points)
    mean_x = sum_x / n
    mean_y = sum_y / n
    ss_xx = sum((x - mean_x) ** 2 for x, _y in points)
    if math.isclose(ss_xx, 0.0):
        raise ValueError("X values are identical; cannot fit Tafel slope.")
    ss_xy = sum((x - mean_x) * (y - mean_y) for x, y in points)
    slope = ss_xy / ss_xx
    intercept = mean_y - slope * mean_x
    ss_tot = sum((y - mean_y) ** 2 for _x, y in points)
    ss_res = sum((y - (slope * x + intercept)) ** 2 for x, y in points)
    r_squared = 1.0 if math.isclose(ss_tot, 0.0) else 1.0 - ss_res / ss_tot
    return slope, intercept, r_squared


def compute_tafel_fits(
    curves: list[tuple[Path, list[Any], CurveMeta]],
    min_j: float,
    max_j: float,
    rhe_offset_v: float = 0.197,
    ph_slope_v: float = 0.0591,
) -> list[dict[str, Any]]:
    if not (min_j > 0 and max_j > min_j):
        raise ValueError("Tafel fit range must satisfy: 0 < min j < max j.")

    used_titles: set[str] = set()
    results: list[dict[str, Any]] = []
    for source, points, meta in curves:
        title = sheet_title(meta.sheet_name or source.name, used_titles)
        rhe_shift = rhe_offset_v + ph_slope_v * meta.ph
        all_rows: list[tuple[float, float, float, float]] = []
        fit_points: list[tuple[float, float]] = []

        for point in points:
            e_rhe = point.vf_v + rhe_shift
            current_density = point.im_a * 1000 / meta.area_cm2
            if current_density <= 0:
                continue
            overpotential_mv = (e_rhe - 1.23) * 1000
            log_j = math.log10(current_density)
            all_rows.append((e_rhe, current_density, log_j, overpotential_mv))
            if min_j <= current_density <= max_j:
                fit_points.append((log_j, overpotential_mv))

        slope = intercept = r_squared = None
        status = "OK"
        if len(fit_points) >= 2:
            slope, intercept, r_squared = linear_regression(fit_points)
        else:
            status = "Not enough points in selected j range"

        results.append(
            {
                "title": title,
                "source_file": str(source),
                "ph": meta.ph,
                "area_cm2": meta.area_cm2,
                "min_j": min_j,
                "max_j": max_j,
                "all_rows": all_rows,
                "fit_points": fit_points,
                "slope_mv_dec": slope,
                "intercept_mv": intercept,
                "r_squared": r_squared,
                "fit_count": len(fit_points),
                "status": status,
            }
        )
    return results


def build_tafel_workbook(fits: list[dict[str, Any]]) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Font

    workbook = Workbook()
    summary = workbook.active
    summary.title = "Tafel_summary"
    summary.append(
        [
            "curve",
            "source_file",
            "pH",
            "area_cm2",
            "fit_min_j_mA_cm2",
            "fit_max_j_mA_cm2",
            "points_used",
            "slope_mV_dec",
            "intercept_mV",
            "R_squared",
            "status",
        ]
    )
    for cell in summary[1]:
        cell.font = Font(bold=True)

    used_sheets = {summary.title}
    for fit in fits:
        summary.append(
            [
                fit["title"],
                fit["source_file"],
                fit["ph"],
                fit["area_cm2"],
                fit["min_j"],
                fit["max_j"],
                fit["fit_count"],
                fit["slope_mv_dec"],
                fit["intercept_mv"],
                fit["r_squared"],
                fit["status"],
            ]
        )

        sheet = workbook.create_sheet(sheet_title(f"Tafel_{fit['title']}", used_sheets))
        sheet.append(["E_RHE_V", "j_mA_cm2", "log10_j", "eta_mV", "used_for_fit"])
        for cell in sheet[1]:
            cell.font = Font(bold=True)
        selected = set(fit["fit_points"])
        for e_rhe, current_density, log_j, overpotential_mv in fit["all_rows"]:
            sheet.append(
                [
                    e_rhe,
                    current_density,
                    log_j,
                    overpotential_mv,
                    "Y" if (log_j, overpotential_mv) in selected else "",
                ]
            )
        for column, width in {"A": 14, "B": 16, "C": 14, "D": 14, "E": 14}.items():
            sheet.column_dimensions[column].width = width
        sheet.freeze_panes = "A2"

    for column, width in {
        "A": 22,
        "B": 64,
        "C": 10,
        "D": 12,
        "E": 16,
        "F": 16,
        "G": 13,
        "H": 16,
        "I": 16,
        "J": 12,
        "K": 28,
    }.items():
        summary.column_dimensions[column].width = width
    summary.freeze_panes = "A2"

    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


def build_tafel_plot_png(fits: list[dict[str, Any]]) -> bytes:
    from PIL import Image, ImageDraw

    plot_series = [fit for fit in fits if fit["fit_points"]]
    if not plot_series:
        raise ValueError("No Tafel points available for plotting.")

    palette = ["#6a6a6a", "#ff4d4d", "#2f80ed", "#43bf79", "#b46ae5", "#e2b500", "#00a6a6", "#f28e2b"]
    all_x = [x for fit in plot_series for x, _y in fit["fit_points"]]
    all_y = [y for fit in plot_series for _x, y in fit["fit_points"]]
    for fit in plot_series:
        if fit["slope_mv_dec"] is not None:
            all_y.extend(
                [
                    fit["slope_mv_dec"] * min(all_x) + fit["intercept_mv"],
                    fit["slope_mv_dec"] * max(all_x) + fit["intercept_mv"],
                ]
            )

    width, height = 1800, 1400
    left, right, top, bottom = 300, 120, 150, 190
    plot_left, plot_top = left, top
    plot_right, plot_bottom = width - right, height - bottom
    plot_width = plot_right - plot_left
    plot_height = plot_bottom - plot_top

    x_ticks = nice_ticks(min(all_x), max(all_x), 5)
    y_ticks = nice_ticks(min(all_y), max(all_y), 5)
    x_min, x_max = min(x_ticks), max(x_ticks)
    y_min, y_max = min(y_ticks), max(y_ticks)

    def x_to_px(x_value: float) -> float:
        return plot_left + (x_value - x_min) / (x_max - x_min) * plot_width

    def y_to_px(y_value: float) -> float:
        return plot_bottom - (y_value - y_min) / (y_max - y_min) * plot_height

    image = Image.new("RGBA", (width, height), "#ffffff")
    draw = ImageDraw.Draw(image)
    font_tick = find_font(50)
    font_label = find_font(64)
    font_legend = find_font(38)
    axis_color = "#000000"
    label_color = "#000000"

    draw.line((plot_left, plot_bottom, plot_right, plot_bottom), fill=axis_color, width=6)
    draw.line((plot_left, plot_top, plot_left, plot_bottom), fill=axis_color, width=6)

    for tick in x_ticks:
        x = x_to_px(tick)
        draw.line((x, plot_bottom, x, plot_bottom + 12), fill=axis_color, width=5)
        draw_text_centered(draw, (x, plot_bottom + 62), format_tick(tick), font_tick, label_color)
    for tick in y_ticks:
        y = y_to_px(tick)
        draw.line((plot_left - 12, y, plot_left, y), fill=axis_color, width=5)
        label = format_tick(tick)
        bbox = draw.textbbox((0, 0), label, font=font_tick)
        draw.text((plot_left - 32 - (bbox[2] - bbox[0]), y - (bbox[3] - bbox[1]) / 2), label, font=font_tick, fill=label_color)

    draw_text_centered(draw, ((plot_left + plot_right) / 2, height - 68), "log10 j/(mA cm-2)", font_label, label_color)
    draw_rotated_label(image, "η/mV", (78, (plot_top + plot_bottom) // 2), font_label, label_color)

    legend_x, legend_y = 340, 210
    for index, fit in enumerate(plot_series):
        color = palette[index % len(palette)]
        points = [(x_to_px(x), y_to_px(y)) for x, y in fit["fit_points"]]
        for x, y in points:
            draw.ellipse((x - 7, y - 7, x + 7, y + 7), fill=color)
        if fit["slope_mv_dec"] is not None:
            x1 = min(x for x, _y in fit["fit_points"])
            x2 = max(x for x, _y in fit["fit_points"])
            y1 = fit["slope_mv_dec"] * x1 + fit["intercept_mv"]
            y2 = fit["slope_mv_dec"] * x2 + fit["intercept_mv"]
            draw.line((x_to_px(x1), y_to_px(y1), x_to_px(x2), y_to_px(y2)), fill=color, width=9)

        legend_yi = legend_y + index * 62
        label = fit["title"]
        if fit["slope_mv_dec"] is not None:
            label = f"{label}: {fit['slope_mv_dec']:.1f} mV/dec, R2={fit['r_squared']:.4f}"
        draw.line((legend_x, legend_yi, legend_x + 120, legend_yi), fill=color, width=10)
        draw.text((legend_x + 145, legend_yi - 24), label, font=font_legend, fill=label_color)

    output = io.BytesIO()
    image.convert("RGB").save(output, format="PNG", optimize=True)
    return output.getvalue()


def make_response_zip(files: dict[str, bytes]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, content in files.items():
            archive.writestr(name, content)
    return buffer.getvalue()


def parse_form(handler: BaseHTTPRequestHandler) -> tuple[list[dict[str, Any]], list[tuple[str, bytes]], dict[str, float]]:
    form = cgi.FieldStorage(
        fp=handler.rfile,
        headers=handler.headers,
        environ={
            "REQUEST_METHOD": "POST",
            "CONTENT_TYPE": handler.headers.get("Content-Type", ""),
        },
        keep_blank_values=True,
    )

    manifest_raw = form.getfirst("manifest", "[]")
    manifest = json.loads(manifest_raw)
    options = {
        "tafel_min_j": float(form.getfirst("tafel_min_j", "10")),
        "tafel_max_j": float(form.getfirst("tafel_max_j", "100")),
    }
    uploaded = form["files"] if "files" in form else []
    if not isinstance(uploaded, list):
        uploaded = [uploaded]

    files: list[tuple[str, bytes]] = []
    for item in uploaded:
        if not getattr(item, "filename", ""):
            continue
        files.append((Path(item.filename).name, item.file.read()))

    return manifest, files, options


class Handler(BaseHTTPRequestHandler):
    server_version = "LSVTool/1.0"

    def log_message(self, format: str, *args: Any) -> None:
        timestamp = time.strftime("%H:%M:%S")
        print(f"[{timestamp}] {self.address_string()} {format % args}")

    def send_bytes(self, status: int, content_type: str, body: bytes, filename: str | None = None) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        if filename:
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        if self.path not in ("/", "/index.html"):
            self.send_bytes(404, "text/plain; charset=utf-8", "Not found".encode("utf-8"))
            return
        self.send_bytes(200, "text/html; charset=utf-8", HTML.encode("utf-8"))

    def do_POST(self) -> None:
        if self.path != "/api/export":
            self.send_bytes(404, "text/plain; charset=utf-8", "Not found".encode("utf-8"))
            return

        try:
            manifest, uploaded_files, options = parse_form(self)
            meta_by_name = {
                item["name"]: CurveMeta(
                    ph=float(item["pH"]),
                    area_cm2=float(item["area_cm2"]),
                    sheet_name=(item.get("sheet_name") or item["name"]),
                )
                for item in manifest
            }

            if not uploaded_files:
                raise ValueError("没有收到文件。")

            curves: list[tuple[Path, list[Any], CurveMeta]] = []
            WORKSPACE_TMP.mkdir(parents=True, exist_ok=True)
            run_dir = WORKSPACE_TMP / f"run_{int(time.time() * 1000)}"
            run_dir.mkdir()
            try:
                tmp = run_dir
                for filename, content in uploaded_files:
                    if filename not in meta_by_name:
                        raise ValueError(f"缺少参数：{filename}")
                    source = tmp / filename
                    source.write_bytes(content)
                    points = extract_vfim(source)
                    curves.append((Path(filename), points, meta_by_name[filename]))

                full_path = tmp / "LSV_full.xlsx"
                write_xlsx(full_path, curves, resistance_ohm=3.8504, rhe_offset_v=0.197, ph_slope_v=0.0591)
                origin_xlsx, origin_csv = build_origin_tables(curves)
                plot_png = build_lsv_plot_png(curves)
                tafel_fits = compute_tafel_fits(
                    curves,
                    min_j=options["tafel_min_j"],
                    max_j=options["tafel_max_j"],
                    rhe_offset_v=0.197,
                    ph_slope_v=0.0591,
                )
                tafel_xlsx = build_tafel_workbook(tafel_fits)
                tafel_png = build_tafel_plot_png(tafel_fits)
                archive = make_response_zip(
                    {
                        "LSV_full.xlsx": full_path.read_bytes(),
                        "Origin_CE.xlsx": origin_xlsx,
                        "Origin_CE.csv": origin_csv,
                        "LSV_plot.png": plot_png,
                        "Tafel_fit.xlsx": tafel_xlsx,
                        "Tafel_plot.png": tafel_png,
                    }
                )
            finally:
                for child in sorted(run_dir.rglob("*"), reverse=True):
                    if child.is_file():
                        child.unlink(missing_ok=True)
                    elif child.is_dir():
                        child.rmdir()
                run_dir.rmdir()

            self.send_bytes(200, "application/zip", archive, filename="LSV_export.zip")
        except Exception as exc:  # noqa: BLE001 - user-facing local tool error.
            message = str(exc).encode("utf-8")
            self.send_bytes(400, "text/plain; charset=utf-8", message)


def main() -> int:
    parser = argparse.ArgumentParser(description="Start the local LSV conversion web app.")
    parser.add_argument("--host", default=HOST, help="Host to bind. Default: 127.0.0.1")
    parser.add_argument("--port", type=int, default=PORT, help="Port to bind. Default: 8765")
    parser.add_argument("--no-open", action="store_true", help="Do not open the browser automatically.")
    args = parser.parse_args()

    server = ThreadingHTTPServer((args.host, args.port), Handler)
    url = f"http://{args.host}:{args.port}"
    print(f"LSV 数据转换工具已启动：{url}")
    print("按 Ctrl+C 停止。")
    if not args.no_open:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止。")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
