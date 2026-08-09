//! grid geometry and the target-agnostic display list: turns a workbook
//! viewport into draw commands. never imports canvas, dom, or any raster backend.

pub mod display_list;
pub mod geometry;
pub mod region;

use std::ops::Range;

use serde::{Deserialize, Serialize};

use xlsx_model::numfmt::{builtin_format_code, format_value};
use xlsx_model::styles::{Border, BorderEdge, BorderStyle, FormatCode, Stylesheet, Theme};
use xlsx_model::value::CellValue;
use xlsx_model::workbook::{Cell, Sheet};
use xlsx_model::{CellRange, CellRef, Fill, HAlign, MAX_COLS, MAX_ROWS, SheetId, VAlign, Workbook};

pub use display_list::{
    Align, DisplayList, DrawCmd, GridMeta, HyperlinkRegion, PathCmd, Rect, scaled,
};
pub use geometry::GridGeometry;
pub use region::{viewport_for_range, viewport_for_used_range};

/// a scrolled window in pixels; `x`/`y` offset the non-frozen body.
#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub struct Viewport {
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
}

const BACKGROUND_COLOR: &str = "#ffffff";
const GRIDLINE_COLOR: &str = "#d4d4d4";
const TEXT_COLOR: &str = "#000000";
const BORDER_COLOR: &str = "#000000";
const PANE_DIVIDER_COLOR: &str = "#8a8a8a";
const HYPERLINK_COLOR: &str = "#0563c1";
const GRIDLINE_WIDTH: f32 = 1.0;
const PANE_DIVIDER_WIDTH: f32 = 2.0;
const FONT_SIZE_PT: f32 = 11.0;
const TEXT_PAD_PX: f32 = 2.0;
// rough calibri-like ascent/descent as fractions of the font size.
const ASCENT_RATIO: f32 = 0.7;
const DESCENT_RATIO: f32 = 0.2;

// ghost pair colors, matching the docx revision palette (del struck / ins).
const GHOST_DEL_COLOR: &str = "#c62828";
const GHOST_INS_COLOR: &str = "#2e7d32";
const GHOST_DEL_HIGHLIGHT: &str = "#c628281a";
const GHOST_INS_HIGHLIGHT: &str = "#2e7d321a";
// conservative per-char advance estimate as a fraction of the font size.
const GHOST_CHAR_W_RATIO: f32 = 0.6;
const GHOST_GAP_PX: f32 = 6.0;
const GHOST_MIN_SCALE: f32 = 0.6;

/// a pending edit rendered as a ghost pair in place of the cell's committed
/// text: `old_text` struck in red, `new_text` in green.
#[derive(Debug, Clone, PartialEq)]
pub struct GhostEdit {
    pub row: u32,
    pub col: u32,
    pub old_text: String,
    pub new_text: String,
    pub alignment_value: CellValue,
}

struct GhostFont {
    size: f32,
    family: Option<String>,
    bold: bool,
    italic: bool,
    underline: bool,
}

#[derive(Clone, Copy)]
struct AxisTrack {
    index: u32,
    raw_start: f32,
    start: f32,
    end: f32,
    pinned: bool,
}

struct AxisLayout {
    tracks: Vec<AxisTrack>,
    ranges: Vec<Range<u32>>,
    divider: Option<f32>,
    frozen: u32,
    scroll: f32,
}

#[derive(Clone, Copy)]
struct AxisSpan {
    raw_start: f32,
    raw_end: f32,
    start: f32,
    end: f32,
}

impl AxisLayout {
    fn new(
        limit: u32,
        frozen: u32,
        scroll: f32,
        extent: f32,
        edge: impl Fn(u32) -> f32,
        at: impl Fn(f32) -> u32,
        hidden: impl Fn(u32) -> bool,
    ) -> Self {
        let frozen = frozen.min(limit);
        let scroll = if frozen > 0 { scroll.max(0.0) } else { scroll };
        let frozen_extent = edge(frozen);
        let mut tracks = Vec::new();
        if frozen > 0 {
            for index in 0..frozen {
                let start = edge(index);
                if start >= extent {
                    break;
                }
                if hidden(index) {
                    continue;
                }
                tracks.push(AxisTrack {
                    index,
                    raw_start: start,
                    start,
                    end: edge(index + 1),
                    pinned: true,
                });
            }
        }

        if frozen < limit && frozen_extent < extent {
            let body_extent = extent - frozen_extent;
            let origin = frozen_extent + scroll;
            let first = at(origin).max(frozen).min(limit - 1);
            let last = at(origin + body_extent).max(first).min(limit - 1);
            let mut clamp_to_divider = frozen > 0;
            for index in first..=last {
                if hidden(index) {
                    continue;
                }
                let raw_start = edge(index) - scroll;
                let raw_end = edge(index + 1) - scroll;
                tracks.push(AxisTrack {
                    index,
                    raw_start,
                    start: if clamp_to_divider {
                        raw_start.max(frozen_extent)
                    } else {
                        raw_start
                    },
                    end: raw_end,
                    pinned: false,
                });
                clamp_to_divider = false;
            }
        }

        let mut ranges: Vec<Range<u32>> = Vec::new();
        for track in &tracks {
            match ranges.last_mut() {
                Some(range) if range.end == track.index => range.end += 1,
                _ => ranges.push(track.index..track.index + 1),
            }
        }
        Self {
            tracks,
            ranges,
            divider: (frozen > 0 && frozen_extent < extent).then_some(frozen_extent),
            frozen,
            scroll,
        }
    }

    fn start(&self) -> u32 {
        self.tracks.first().map_or(0, |track| track.index)
    }

    fn indices(&self) -> Option<Vec<u32>> {
        let start = self.start();
        self.tracks
            .iter()
            .enumerate()
            .any(|(offset, track)| track.index != start + offset as u32)
            .then(|| self.tracks.iter().map(|track| track.index).collect())
    }

    fn offsets(&self) -> Vec<f32> {
        let Some(first) = self.tracks.first() else {
            return Vec::new();
        };
        std::iter::once(first.start)
            .chain(self.tracks.iter().map(|track| track.end))
            .collect()
    }

    fn contains(&self, index: u32) -> bool {
        self.tracks
            .binary_search_by_key(&index, |track| track.index)
            .is_ok()
    }

    fn intersects(&self, start: u32, end: u32) -> bool {
        self.tracks
            .iter()
            .any(|track| (start..=end).contains(&track.index))
    }

    fn span(&self, start: u32, end: u32, edge: impl Fn(u32) -> f32) -> Option<AxisSpan> {
        let first = self
            .tracks
            .binary_search_by_key(&start, |track| track.index)
            .ok()?;
        let track = self.tracks[first];
        let raw_end = if track.pinned {
            edge(end.saturating_add(1).min(self.frozen))
        } else {
            edge(end.saturating_add(1)) - self.scroll
        };
        Some(AxisSpan {
            raw_start: track.raw_start,
            raw_end,
            start: track.start,
            end: raw_end,
        })
    }
}

#[derive(Clone, Copy)]
struct CellBox {
    x: f32,
    y: f32,
    w: f32,
    h: f32,
    clip: Rect,
}

/// build the display list for one viewport of one sheet. commands are emitted
/// background -> fills -> gridlines -> borders -> text -> pane dividers.
pub fn build_display_list(wb: &Workbook, sheet: SheetId, viewport: &Viewport) -> DisplayList {
    build_display_list_with_ghosts(wb, sheet, viewport, &[])
}

/// like [`build_display_list`], with ghost pairs painted in place of the
/// committed text of each ghost's cell.
pub fn build_display_list_with_ghosts(
    wb: &Workbook,
    sheet: SheetId,
    viewport: &Viewport,
    ghosts: &[GhostEdit],
) -> DisplayList {
    let mut commands = Vec::new();

    commands.push(DrawCmd::FillRect {
        x: 0.0,
        y: 0.0,
        w: viewport.width,
        h: viewport.height,
        color: BACKGROUND_COLOR.to_string(),
    });

    let Some(sheet_ref) = wb.sheet(sheet) else {
        return DisplayList {
            width: viewport.width,
            height: viewport.height,
            commands,
            grid: GridMeta::default(),
            hyperlinks: Vec::new(),
        };
    };

    let styles = &wb.styles;
    let theme = &styles.theme;
    let hyperlink_color = theme.slot(10).unwrap_or(HYPERLINK_COLOR);
    let geom = GridGeometry::new(sheet_ref);
    let (frozen_rows, frozen_cols) = sheet_ref
        .freeze_pane
        .map_or((0, 0), |pane| (pane.rows, pane.cols));
    let rows = AxisLayout::new(
        MAX_ROWS,
        frozen_rows,
        viewport.y,
        viewport.height,
        |row| geom.row_y(row),
        |y| geom.row_at_y(y),
        |row| sheet_ref.is_row_hidden(row),
    );
    let cols = AxisLayout::new(
        MAX_COLS,
        frozen_cols,
        viewport.x,
        viewport.width,
        |col| geom.col_x(col),
        |x| geom.col_at_x(x),
        |_| false,
    );

    let grid = GridMeta {
        start_row: rows.start(),
        start_col: cols.start(),
        row_indices: rows.indices(),
        col_indices: cols.indices(),
        row_offsets: rows.offsets(),
        col_offsets: cols.offsets(),
    };
    let hyperlinks = sheet_ref
        .hyperlinks
        .iter()
        .filter(|link| {
            rows.intersects(link.range.start.row, link.range.end.row)
                && cols.intersects(link.range.start.col, link.range.end.col)
        })
        .map(|link| HyperlinkRegion {
            top: link.range.start.row,
            left: link.range.start.col,
            bottom: link.range.end.row,
            right: link.range.end.col,
            external_target: link.external_target.clone(),
            location: link.location.clone(),
            tooltip: link.tooltip.clone(),
        })
        .collect();
    let anchors = visible_anchors(sheet_ref, &rows, &cols);
    let changed_ghost_cells: std::collections::HashSet<(u32, u32)> = ghosts
        .iter()
        .filter(|g| g.old_text != g.new_text)
        .map(|g| (g.row, g.col))
        .collect();

    for &(at, cell) in &anchors {
        let Some(style) = cell.style else { continue };
        let Some(Fill::Solid(color)) = styles.fill_for(style) else {
            continue;
        };
        let Some(hex) = color.resolve(theme) else {
            continue;
        };
        let Some(cell_box) = cell_box(&geom, &rows, &cols, sheet_ref, at) else {
            continue;
        };
        let clip = cell_box.clip;
        commands.push(DrawCmd::FillRect {
            x: clip.x,
            y: clip.y,
            w: clip.w,
            h: clip.h,
            color: hex,
        });
    }

    let row_offsets = rows.offsets();
    let col_offsets = cols.offsets();
    let top = row_offsets.first().copied().unwrap_or(0.0);
    let bottom = row_offsets.last().copied().unwrap_or(0.0);
    let left = col_offsets.first().copied().unwrap_or(0.0);
    let right = col_offsets.last().copied().unwrap_or(0.0);
    for &x in &col_offsets {
        commands.push(DrawCmd::Line {
            x1: x,
            y1: top,
            x2: x,
            y2: bottom,
            width: GRIDLINE_WIDTH,
            color: GRIDLINE_COLOR.to_string(),
            style: None,
        });
    }
    for &y in &row_offsets {
        commands.push(DrawCmd::Line {
            x1: left,
            y1: y,
            x2: right,
            y2: y,
            width: GRIDLINE_WIDTH,
            color: GRIDLINE_COLOR.to_string(),
            style: None,
        });
    }

    for &(at, cell) in &anchors {
        let Some(style) = cell.style else { continue };
        let Some(border) = styles.border_for(style) else {
            continue;
        };
        emit_borders(
            &mut commands,
            &geom,
            &rows,
            &cols,
            sheet_ref,
            styles,
            theme,
            at,
            border,
        );
    }

    for &(at, cell) in &anchors {
        if changed_ghost_cells.contains(&(at.row, at.col)) {
            continue;
        }
        let hyperlink = sheet_ref.hyperlink_at(at);
        let Some((text, color)) = cell_display_text(styles, wb.date_system, cell).or_else(|| {
            hyperlink
                .filter(|link| link.range.start == at)
                .and_then(|link| link.display.clone())
                .filter(|display| !display.is_empty())
                .map(|display| (display, hyperlink_color.to_string()))
        }) else {
            continue;
        };
        let color = if hyperlink.is_some() {
            hyperlink_color.to_string()
        } else {
            color
        };

        let Some(cell_box) = cell_box(&geom, &rows, &cols, sheet_ref, at) else {
            continue;
        };
        let font = cell.style.and_then(|s| styles.font_for(s));
        let size = font
            .and_then(|f| f.size_pt)
            .map(|p| p as f32)
            .unwrap_or(FONT_SIZE_PT);
        let align = resolve_align(styles, cell);
        let valign = cell
            .style
            .and_then(|s| styles.alignment_for(s))
            .and_then(|a| a.v);

        let tx = match align {
            Align::Left => cell_box.x + TEXT_PAD_PX,
            Align::Right => cell_box.x + cell_box.w - TEXT_PAD_PX,
            Align::Center => cell_box.x + cell_box.w / 2.0,
        };
        let ty = baseline_y(cell_box.y, cell_box.h, size, valign);

        commands.push(DrawCmd::Text {
            x: tx,
            y: ty,
            text,
            font_size: size,
            color,
            clip: cell_box.clip,
            align,
            bold: font.is_some_and(|f| f.bold),
            italic: font.is_some_and(|f| f.italic),
            underline: hyperlink.is_some() || font.is_some_and(|f| f.underline),
            strike: font.is_some_and(|f| f.strike),
            highlight: None,
            dashed_underline: false,
            font_family: font.and_then(|f| f.name.clone()),
            ghost: false,
        });
    }

    for link in &sheet_ref.hyperlinks {
        let at = link.range.start;
        if sheet_ref.cell(at).is_some() || !rows.contains(at.row) || !cols.contains(at.col) {
            continue;
        }
        let Some(text) = link.display.as_ref().filter(|display| !display.is_empty()) else {
            continue;
        };
        let Some(cell_box) = cell_box(&geom, &rows, &cols, sheet_ref, at) else {
            continue;
        };
        commands.push(DrawCmd::Text {
            x: cell_box.x + TEXT_PAD_PX,
            y: baseline_y(cell_box.y, cell_box.h, FONT_SIZE_PT, None),
            text: text.clone(),
            font_size: FONT_SIZE_PT,
            color: hyperlink_color.to_string(),
            clip: cell_box.clip,
            align: Align::Left,
            bold: false,
            italic: false,
            underline: true,
            strike: false,
            highlight: None,
            dashed_underline: false,
            font_family: None,
            ghost: false,
        });
    }

    for ghost in ghosts {
        if !rows.contains(ghost.row) || !cols.contains(ghost.col) {
            continue;
        }
        let at = CellRef::new(ghost.row, ghost.col);
        let cell = sheet_ref.cell(at);
        let font = cell.and_then(|c| c.style).and_then(|s| styles.font_for(s));
        let font = GhostFont {
            size: font
                .and_then(|font| font.size_pt)
                .map(|size| size as f32)
                .unwrap_or(FONT_SIZE_PT),
            family: font.and_then(|font| font.name.clone()),
            bold: font.is_some_and(|font| font.bold),
            italic: font.is_some_and(|font| font.italic),
            underline: font.is_some_and(|font| font.underline),
        };
        let Some(bx) = cell_box(&geom, &rows, &cols, sheet_ref, at) else {
            continue;
        };
        let align = resolve_align_with_value(styles, cell, &ghost.alignment_value);
        emit_ghost(&mut commands, ghost, bx, font, align);
    }

    emit_charts(
        &mut commands,
        wb,
        sheet_ref,
        &geom,
        viewport,
        frozen_rows,
        frozen_cols,
    );

    if let Some(x) = cols.divider {
        commands.push(DrawCmd::Line {
            x1: x,
            y1: 0.0,
            x2: x,
            y2: viewport.height,
            width: PANE_DIVIDER_WIDTH,
            color: PANE_DIVIDER_COLOR.to_string(),
            style: None,
        });
    }
    if let Some(y) = rows.divider {
        commands.push(DrawCmd::Line {
            x1: 0.0,
            y1: y,
            x2: viewport.width,
            y2: y,
            width: PANE_DIVIDER_WIDTH,
            color: PANE_DIVIDER_COLOR.to_string(),
            style: None,
        });
    }

    DisplayList {
        width: viewport.width,
        height: viewport.height,
        commands,
        grid,
        hyperlinks,
    }
}

const EMU_PER_PX: f64 = 9525.0;

/// paint each anchored chart, clipped to the pane its anchor cell lives in so
/// floating charts never bleed across a frozen-pane divider.
fn emit_charts(
    commands: &mut Vec<DrawCmd>,
    wb: &Workbook,
    sheet: &Sheet,
    geom: &GridGeometry,
    viewport: &Viewport,
    frozen_rows: u32,
    frozen_cols: u32,
) {
    use ooxml_drawingml::chart::{PlotChart, PlotRect, plot_chart_into};

    if sheet.drawings.is_empty() {
        return;
    }
    let frozen_x = geom.col_x(frozen_cols);
    let frozen_y = geom.row_y(frozen_rows);
    for drawing in &sheet.drawings {
        let (x0, y0, x1, y1, anchor_col, anchor_row) = match &drawing.anchor {
            xlsx_model::DrawingAnchor::Cell {
                from,
                to,
                extent_emu,
            } => {
                let x0 = f64::from(geom.col_x(from.col)) + from.col_offset_emu as f64 / EMU_PER_PX;
                let y0 = f64::from(geom.row_y(from.row)) + from.row_offset_emu as f64 / EMU_PER_PX;
                let (x1, y1) = match (to, extent_emu) {
                    (Some(to), _) => (
                        f64::from(geom.col_x(to.col)) + to.col_offset_emu as f64 / EMU_PER_PX,
                        f64::from(geom.row_y(to.row)) + to.row_offset_emu as f64 / EMU_PER_PX,
                    ),
                    (None, Some((cx, cy))) => {
                        (x0 + *cx as f64 / EMU_PER_PX, y0 + *cy as f64 / EMU_PER_PX)
                    }
                    (None, None) => continue,
                };
                (x0, y0, x1, y1, from.col, from.row)
            }
            xlsx_model::DrawingAnchor::Absolute {
                pos_emu,
                extent_emu,
            } => {
                let x0 = pos_emu.0 as f64 / EMU_PER_PX;
                let y0 = pos_emu.1 as f64 / EMU_PER_PX;
                let x1 = x0 + extent_emu.0 as f64 / EMU_PER_PX;
                let y1 = y0 + extent_emu.1 as f64 / EMU_PER_PX;
                (
                    x0,
                    y0,
                    x1,
                    y1,
                    geom.col_at_x(x0 as f32),
                    geom.row_at_y(y0 as f32),
                )
            }
        };
        if x1 <= x0 || y1 <= y0 {
            continue;
        }

        let axis = |pinned: bool, pos: f64, scroll: f32, frozen: u32, extent: f32, edge: f32| {
            if pinned {
                (pos, 0.0f32, edge.min(extent))
            } else {
                (
                    pos - f64::from(scroll),
                    if frozen > 0 { edge } else { 0.0 },
                    extent,
                )
            }
        };
        let (vx, clip_left, clip_right) = axis(
            frozen_cols > 0 && anchor_col < frozen_cols,
            x0,
            viewport.x,
            frozen_cols,
            viewport.width,
            frozen_x,
        );
        let (vy, clip_top, clip_bottom) = axis(
            frozen_rows > 0 && anchor_row < frozen_rows,
            y0,
            viewport.y,
            frozen_rows,
            viewport.height,
            frozen_y,
        );
        let (w, h) = (x1 - x0, y1 - y0);
        if clip_right <= clip_left
            || clip_bottom <= clip_top
            || vx + w <= f64::from(clip_left)
            || vy + h <= f64::from(clip_top)
            || vx >= f64::from(clip_right)
            || vy >= f64::from(clip_bottom)
        {
            continue;
        }

        let clip = Rect {
            x: clip_left,
            y: clip_top,
            w: clip_right - clip_left,
            h: clip_bottom - clip_top,
        };
        commands.push(DrawCmd::PushClip {
            x: clip.x,
            y: clip.y,
            w: clip.w,
            h: clip.h,
        });
        let chart = refreshed_chart(wb, sheet, &drawing.chart);
        plot_chart_into(
            &PlotChart::from(&chart),
            PlotRect { x: vx, y: vy, w, h },
            &mut ChartSink { commands, clip },
        );
        commands.push(DrawCmd::PopClip {});
    }
}

/// The most cells one chart series reads back, bounding the per-frame cost of
/// a formula that spans a huge range.
const MAX_CHART_SERIES_CELLS: usize = 4096;

/// Re-reads each series' cached data from its source range so charts track
/// live cell edits; ranges that no longer resolve keep their cached values.
fn refreshed_chart(
    wb: &Workbook,
    host: &Sheet,
    chart: &ooxml_drawingml::chart::ChartSpace,
) -> ooxml_drawingml::chart::ChartSpace {
    let mut chart = chart.clone();
    let group_series = chart
        .plot_groups
        .iter_mut()
        .flat_map(|group| group.series.iter_mut());
    for series in chart.series.iter_mut().chain(group_series) {
        if let Some(cells) = series
            .value_formula
            .as_deref()
            .and_then(|formula| range_cells(wb, host, formula))
        {
            series.values = cells
                .into_iter()
                .map(|cell| match cell.map(|cell| &cell.value) {
                    Some(CellValue::Number { value }) => *value,
                    Some(CellValue::Bool { value }) => f64::from(*value),
                    _ => f64::NAN,
                })
                .collect();
        }
        if let Some(cells) = series
            .category_formula
            .as_deref()
            .and_then(|formula| range_cells(wb, host, formula))
        {
            series.categories = cells
                .into_iter()
                .map(|cell| match cell.map(|cell| &cell.value) {
                    Some(CellValue::Text { value }) => value.clone(),
                    Some(CellValue::Number { value }) => trimmed_number(*value),
                    Some(CellValue::Bool { value }) => {
                        if *value { "TRUE" } else { "FALSE" }.to_owned()
                    }
                    _ => String::new(),
                })
                .collect();
        }
    }
    chart
}

/// Cells of a `Sheet1!$B$2:$B$4`-style reference in row-major order, or none
/// when the reference does not parse, names an unknown sheet, or is too big.
fn range_cells<'a>(
    wb: &'a Workbook,
    host: &'a Sheet,
    formula: &str,
) -> Option<Vec<Option<&'a Cell>>> {
    let (name, range) = xlsx_model::parse_sheet_range(formula)?;
    let sheet = match name {
        Some(name) => wb.sheet_by_name(&name).map(|(_, sheet)| sheet)?,
        None => host,
    };
    let rows = range.end.row.checked_sub(range.start.row)? as usize + 1;
    let cols = range.end.col.checked_sub(range.start.col)? as usize + 1;
    if rows.checked_mul(cols)? > MAX_CHART_SERIES_CELLS {
        return None;
    }
    let mut cells = Vec::with_capacity(rows * cols);
    for row in range.start.row..=range.end.row {
        for col in range.start.col..=range.end.col {
            cells.push(sheet.cell(CellRef::new(row, col)));
        }
    }
    Some(cells)
}

/// A short display form for numeric category labels.
fn trimmed_number(value: f64) -> String {
    if value.fract() == 0.0 && value.abs() < 1e15 {
        format!("{}", value as i64)
    } else {
        format!("{value}")
    }
}

/// translates chart plot ops into display-list commands as they are emitted.
struct ChartSink<'a> {
    commands: &'a mut Vec<DrawCmd>,
    clip: Rect,
}

impl ooxml_drawingml::chart::PlotSink for ChartSink<'_> {
    fn push_op(&mut self, op: ooxml_drawingml::chart::PlotOp) {
        use ooxml_drawingml::GeometryPathCommand;
        use ooxml_drawingml::chart::PlotOp;

        match op {
            PlotOp::Rect { x, y, w, h, fill } => self.commands.push(DrawCmd::FillRect {
                x: x as f32,
                y: y as f32,
                w: w as f32,
                h: h as f32,
                color: fill,
            }),
            PlotOp::Text {
                text,
                x,
                baseline_y,
                width: _,
                font,
                color,
            } => self.commands.push(DrawCmd::Text {
                x: x as f32,
                y: baseline_y as f32,
                text,
                font_size: font.size_px as f32,
                color,
                clip: self.clip,
                align: Align::Left,
                bold: font.weight >= 600,
                italic: false,
                underline: false,
                strike: false,
                highlight: None,
                dashed_underline: false,
                font_family: None,
                ghost: false,
            }),
            PlotOp::Line {
                x1,
                y1,
                x2,
                y2,
                color,
                width,
            } => self.commands.push(DrawCmd::Line {
                x1: x1 as f32,
                y1: y1 as f32,
                x2: x2 as f32,
                y2: y2 as f32,
                width: width as f32,
                color,
                style: None,
            }),
            PlotOp::Path {
                commands,
                fill,
                stroke,
                ..
            } => self.commands.push(DrawCmd::Path {
                commands: commands
                    .into_iter()
                    .map(|command| match command {
                        GeometryPathCommand::Move { x, y } => PathCmd::MoveTo {
                            x: x as f32,
                            y: y as f32,
                        },
                        GeometryPathCommand::Line { x, y } => PathCmd::LineTo {
                            x: x as f32,
                            y: y as f32,
                        },
                        GeometryPathCommand::Quad { cpx, cpy, x, y } => PathCmd::QuadTo {
                            cx: cpx as f32,
                            cy: cpy as f32,
                            x: x as f32,
                            y: y as f32,
                        },
                        GeometryPathCommand::Cubic {
                            cp1x,
                            cp1y,
                            cp2x,
                            cp2y,
                            x,
                            y,
                        } => PathCmd::CubicTo {
                            x1: cp1x as f32,
                            y1: cp1y as f32,
                            x2: cp2x as f32,
                            y2: cp2y as f32,
                            x: x as f32,
                            y: y as f32,
                        },
                        GeometryPathCommand::Close => PathCmd::Close {},
                    })
                    .collect(),
                fill: Some(fill),
                stroke: stroke.as_ref().map(|stroke| stroke.color.clone()),
                stroke_width: stroke.map_or(0.0, |stroke| stroke.width as f32),
            }),
        }
    }
}

/// paint one pending edit inside a cell box.
fn emit_ghost(
    commands: &mut Vec<DrawCmd>,
    ghost: &GhostEdit,
    cell_box: CellBox,
    font: GhostFont,
    single_align: Align,
) {
    let old = ghost.old_text.as_str();
    let new = ghost.new_text.as_str();
    if old == new {
        return;
    }

    let (cx0, cy0, cw, ch) = (cell_box.x, cell_box.y, cell_box.w, cell_box.h);
    let clip = cell_box.clip;
    let x = cx0 + TEXT_PAD_PX;
    let avail = (cw - 2.0 * TEXT_PAD_PX).max(0.0);
    let full_size = font.size;

    let mut line = |x: f32,
                    y: f32,
                    text: String,
                    size: f32,
                    color: &str,
                    align: Align,
                    strike: bool,
                    preview: bool| {
        commands.push(DrawCmd::Text {
            x,
            y,
            text,
            font_size: size,
            color: color.to_string(),
            clip,
            align,
            bold: font.bold,
            italic: font.italic,
            underline: font.underline,
            strike,
            highlight: Some(
                if preview {
                    GHOST_INS_HIGHLIGHT
                } else {
                    GHOST_DEL_HIGHLIGHT
                }
                .to_string(),
            ),
            dashed_underline: preview,
            font_family: font.family.clone(),
            ghost: preview,
        });
    };

    if old.is_empty() || new.is_empty() {
        let (text, color, strike, preview) = if old.is_empty() {
            (new, GHOST_INS_COLOR, false, true)
        } else {
            (old, GHOST_DEL_COLOR, true, false)
        };
        let x = match single_align {
            Align::Left => cx0 + TEXT_PAD_PX,
            Align::Right => cx0 + cw - TEXT_PAD_PX,
            Align::Center => cx0 + cw / 2.0,
        };
        line(
            x,
            baseline_y(cy0, ch, full_size, None),
            ellipsize(text, avail, full_size),
            full_size,
            color,
            single_align,
            strike,
            preview,
        );
        return;
    }

    let old_width = ghost_text_width(old, full_size);
    let new_width = ghost_text_width(new, full_size);
    if old_width + GHOST_GAP_PX + new_width <= avail {
        let baseline = baseline_y(cy0, ch, full_size, None);
        line(
            x,
            baseline,
            old.to_string(),
            full_size,
            GHOST_DEL_COLOR,
            Align::Left,
            true,
            false,
        );
        line(
            x + old_width + GHOST_GAP_PX,
            baseline,
            new.to_string(),
            full_size,
            GHOST_INS_COLOR,
            Align::Left,
            false,
            true,
        );
        return;
    }

    let line_ratio = ASCENT_RATIO + DESCENT_RATIO;
    let scale = (ch / (2.0 * full_size * line_ratio)).clamp(GHOST_MIN_SCALE, 1.0);
    let size = full_size * scale;
    let line_h = size * line_ratio;
    let top = cy0 + ((ch - 2.0 * line_h) / 2.0).max(0.0);
    let first_baseline = top + size * ASCENT_RATIO;
    line(
        x,
        first_baseline,
        ellipsize(new, avail, size),
        size,
        GHOST_INS_COLOR,
        Align::Left,
        false,
        true,
    );
    line(
        x,
        first_baseline + line_h,
        ellipsize(old, avail, size),
        size,
        GHOST_DEL_COLOR,
        Align::Left,
        true,
        false,
    );
}

/// estimated advance width of `text` at `size`, deliberately generous so fit
/// decisions err toward ellipsizing rather than overlap.
fn ghost_text_width(text: &str, size: f32) -> f32 {
    text.chars().count() as f32 * size * GHOST_CHAR_W_RATIO
}

/// `text` unchanged when its estimate fits `budget`, else a truncated prefix
/// ending in `…`.
fn ellipsize(text: &str, budget: f32, size: f32) -> String {
    if ghost_text_width(text, size) <= budget {
        return text.to_string();
    }
    let char_w = size * GHOST_CHAR_W_RATIO;
    let keep = ((budget / char_w) as i32 - 1).max(0) as usize;
    let prefix: String = text.chars().take(keep).collect();
    format!("{prefix}…")
}

/// visible cells that draw: inside the range and not a covered merge cell
/// (only a merge's anchor draws). yields `(anchor, cell)` in row-major order.
fn visible_anchors<'a>(
    sheet: &'a Sheet,
    rows: &AxisLayout,
    cols: &AxisLayout,
) -> Vec<(CellRef, &'a xlsx_model::Cell)> {
    let mut cells = Vec::new();
    for row_range in &rows.ranges {
        for col_range in &cols.ranges {
            cells.extend(sheet.iter_cells_in_rect(row_range.clone(), col_range.clone()));
        }
    }
    cells.sort_unstable_by_key(|(at, _)| (at.row, at.col));
    cells.dedup_by_key(|(at, _)| (at.row, at.col));
    cells.retain(|(at, _)| match covering_merge(&sheet.merges, *at) {
        Some(merge) => merge.start == *at,
        None => true,
    });
    cells
}

/// the merge (if any) that covers a cell.
fn covering_merge(merges: &[CellRange], at: CellRef) -> Option<CellRange> {
    merges.iter().copied().find(|m| m.contains(at))
}

/// viewport-local `(x, y, w, h)` of a cell's box, spanning its merged range
/// when `at` anchors one.
fn cell_box(
    geom: &GridGeometry,
    rows: &AxisLayout,
    cols: &AxisLayout,
    sheet: &Sheet,
    at: CellRef,
) -> Option<CellBox> {
    let (end_col, end_row) = match covering_merge(&sheet.merges, at) {
        Some(merge) => (merge.end.col, merge.end.row),
        None => (at.col, at.row),
    };
    let col = cols.span(at.col, end_col, |column| geom.col_x(column))?;
    let row = rows.span(at.row, end_row, |row| geom.row_y(row))?;
    Some(CellBox {
        x: col.raw_start,
        y: row.raw_start,
        w: col.raw_end - col.raw_start,
        h: row.raw_end - row.raw_start,
        clip: Rect {
            x: col.start,
            y: row.start,
            w: col.end - col.start,
            h: row.end - row.start,
        },
    })
}

/// display string and resolved font color for a cell, or `None` when it renders
/// nothing. a `[Red]`-style number-format color overrides the font color.
fn cell_display_text(
    styles: &Stylesheet,
    date_system: xlsx_model::DateSystem,
    cell: &xlsx_model::Cell,
) -> Option<(String, String)> {
    if matches!(cell.value, CellValue::Empty) {
        return None;
    }
    let code = format_code_for_cell(styles, cell);
    let formatted = format_value(&cell.value, &code, date_system);
    if formatted.text.is_empty() {
        return None;
    }
    let font = cell.style.and_then(|s| styles.font_for(s));
    let color = formatted
        .color
        .or_else(|| {
            font.and_then(|f| f.color.as_ref())
                .and_then(|c| c.resolve(&styles.theme))
        })
        .unwrap_or_else(|| TEXT_COLOR.to_string());
    Some((formatted.text, color))
}

/// the number-format code a cell's xf resolves to; general when unset or when
/// a builtin id is not modeled.
fn format_code_for_cell(styles: &Stylesheet, cell: &xlsx_model::Cell) -> String {
    match cell.style.map(|s| styles.format_code_for(s)) {
        Some(FormatCode::Custom(c)) => c.to_string(),
        Some(FormatCode::Builtin(id)) => builtin_format_code(id).unwrap_or("General").to_string(),
        None => "General".to_string(),
    }
}

/// the exact string the grid would paint for `cell`, number-format aware.
/// empty cells and formats that yield nothing render as "".
pub fn display_text(
    styles: &Stylesheet,
    date_system: xlsx_model::DateSystem,
    cell: &xlsx_model::Cell,
) -> String {
    if matches!(cell.value, CellValue::Empty) {
        return String::new();
    }
    let code = format_code_for_cell(styles, cell);
    format_value(&cell.value, &code, date_system).text
}

/// horizontal anchor for a cell: an explicit xf alignment wins, otherwise the
/// value type decides (numbers right, booleans center, text/errors left).
fn resolve_align(styles: &Stylesheet, cell: &xlsx_model::Cell) -> Align {
    resolve_align_with_value(styles, Some(cell), &cell.value)
}

fn resolve_align_with_value(
    styles: &Stylesheet,
    cell: Option<&xlsx_model::Cell>,
    value: &CellValue,
) -> Align {
    let type_default = match value {
        CellValue::Number { .. } => Align::Right,
        CellValue::Bool { .. } => Align::Center,
        _ => Align::Left,
    };
    let h = cell
        .and_then(|cell| cell.style)
        .and_then(|s| styles.alignment_for(s))
        .and_then(|a| a.h);
    match h {
        Some(HAlign::Left) | Some(HAlign::Fill) | Some(HAlign::Justify) => Align::Left,
        Some(HAlign::Right) => Align::Right,
        Some(HAlign::Center) | Some(HAlign::CenterContinuous) | Some(HAlign::Distributed) => {
            Align::Center
        }
        Some(HAlign::General) | None => type_default,
    }
}

/// baseline y for a cell's text given its vertical alignment; unset (or center)
/// keeps the centered baseline.
fn baseline_y(cy0: f32, ch: f32, size: f32, valign: Option<VAlign>) -> f32 {
    match valign {
        Some(VAlign::Top) => cy0 + TEXT_PAD_PX + size * ASCENT_RATIO,
        Some(VAlign::Bottom) => cy0 + ch - TEXT_PAD_PX - size * DESCENT_RATIO,
        _ => cy0 + (ch + size * ASCENT_RATIO) / 2.0,
    }
}

/// emit the set edges of a cell's border. a shared interior edge draws once:
/// the bottom (right) edge is skipped when the neighbor declares its own top (left) edge.
#[allow(clippy::too_many_arguments)]
fn emit_borders(
    commands: &mut Vec<DrawCmd>,
    geom: &GridGeometry,
    rows: &AxisLayout,
    cols: &AxisLayout,
    sheet: &Sheet,
    styles: &Stylesheet,
    theme: &Theme,
    at: CellRef,
    border: &Border,
) {
    let Some(cell_box) = cell_box(geom, rows, cols, sheet, at) else {
        return;
    };
    let (x, y) = (cell_box.x, cell_box.y);
    let (x2, y2) = (x + cell_box.w, y + cell_box.h);
    let clip = cell_box.clip;
    let (clip_x2, clip_y2) = (clip.x + clip.w, clip.y + clip.h);
    let (end_col, end_row) = match covering_merge(&sheet.merges, at) {
        Some(m) => (m.end.col, m.end.row),
        None => (at.col, at.row),
    };

    if let Some(edge) = &border.top
        && y >= clip.y
        && y <= clip_y2
    {
        commands.push(border_line(clip.x, y, clip_x2, y, edge, theme));
    }
    if let Some(edge) = &border.left
        && x >= clip.x
        && x <= clip_x2
    {
        commands.push(border_line(x, clip.y, x, clip_y2, edge, theme));
    }
    if let Some(edge) = &border.bottom
        && y2 >= clip.y
        && y2 <= clip_y2
        && !neighbor_edge(sheet, styles, end_row + 1, at.col, |b| b.top.is_some())
    {
        commands.push(border_line(clip.x, y2, clip_x2, y2, edge, theme));
    }
    if let Some(edge) = &border.right
        && x2 >= clip.x
        && x2 <= clip_x2
        && !neighbor_edge(sheet, styles, at.row, end_col + 1, |b| b.left.is_some())
    {
        commands.push(border_line(x2, clip.y, x2, clip_y2, edge, theme));
    }
}

/// true when the cell at `(row, col)` has a border satisfying `pick`.
fn neighbor_edge(
    sheet: &Sheet,
    styles: &Stylesheet,
    row: u32,
    col: u32,
    pick: impl Fn(&Border) -> bool,
) -> bool {
    sheet
        .cell(CellRef::new(row, col))
        .and_then(|c| c.style)
        .and_then(|s| styles.border_for(s))
        .is_some_and(pick)
}

/// one border edge as a `Line`, mapping the weight to a stroke width and dash
/// style; an unset edge color resolves to black, matching excel's automatic color.
fn border_line(x1: f32, y1: f32, x2: f32, y2: f32, edge: &BorderEdge, theme: &Theme) -> DrawCmd {
    let (width, style) = border_stroke(edge.style);
    let color = edge
        .color
        .as_ref()
        .and_then(|c| c.resolve(theme))
        .unwrap_or_else(|| BORDER_COLOR.to_string());
    DrawCmd::Line {
        x1,
        y1,
        x2,
        y2,
        width,
        color,
        style,
    }
}

/// map a border weight to a `(stroke width, dash style)`.
fn border_stroke(style: BorderStyle) -> (f32, Option<String>) {
    match style {
        BorderStyle::Hair => (1.0, Some("dotted".to_string())),
        BorderStyle::Thin => (1.0, None),
        BorderStyle::Medium => (2.0, None),
        BorderStyle::Thick => (3.0, None),
        BorderStyle::Dashed => (1.0, Some("dashed".to_string())),
        BorderStyle::Dotted => (1.0, Some("dotted".to_string())),
        BorderStyle::Double => (1.0, Some("double".to_string())),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::Hyperlink;
    use xlsx_model::workbook::{AnchorCell, Cell, DrawingAnchor, FreezePane, Sheet, SheetDrawing};

    fn text_cell(s: &str) -> Cell {
        Cell {
            value: CellValue::Text { value: s.into() },
            ..Cell::default()
        }
    }
    fn num_cell(n: f64) -> Cell {
        Cell {
            value: CellValue::Number { value: n },
            ..Cell::default()
        }
    }

    #[test]
    fn structural_order_and_clip_rect() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(CellRef::new(0, 0), num_cell(42.0));
        sheet.set_cell(CellRef::new(0, 1), text_cell("hi"));
        let long = "a very long label that overflows its cell";
        sheet.set_cell(CellRef::new(0, 2), text_cell(long));
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);

        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: 400.0,
            height: 100.0,
        };
        let dl = build_display_list(&wb, SheetId(0), &vp);

        assert_eq!(dl.width, 400.0);
        assert!(matches!(dl.commands[0], DrawCmd::FillRect { .. }));

        let first_text = dl
            .commands
            .iter()
            .position(|c| matches!(c, DrawCmd::Text { .. }));
        let last_line = dl
            .commands
            .iter()
            .rposition(|c| matches!(c, DrawCmd::Line { .. }));
        assert!(first_text.is_some() && last_line.is_some());
        assert!(last_line.unwrap() < first_text.unwrap());

        let texts: Vec<_> = dl
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCmd::Text { .. }))
            .collect();
        assert_eq!(texts.len(), 3);

        let long_text = texts
            .iter()
            .find_map(|c| match c {
                DrawCmd::Text { text, clip, .. } if text == long => Some(clip),
                _ => None,
            })
            .unwrap();
        let dc = geometry::col_chars_to_px(geometry::DEFAULT_COL_WIDTH_CHARS);
        assert_eq!(long_text.x, dc * 2.0);
        assert_eq!(long_text.w, dc);
    }

    #[test]
    fn renders_hyperlink_indication_and_hit_regions() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(CellRef::new(0, 0), text_cell("Website"));
        sheet.hyperlinks.push(Hyperlink {
            range: CellRange::parse_a1("A1:B1").unwrap(),
            external_target: Some("https://example.com".into()),
            location: None,
            tooltip: Some("Open site".into()),
            display: None,
        });
        sheet.hyperlinks.push(Hyperlink {
            range: CellRange::parse_a1("C3").unwrap(),
            external_target: None,
            location: Some("Sheet1!A1".into()),
            tooltip: None,
            display: Some("Jump".into()),
        });
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);

        let dl = build_display_list(
            &wb,
            SheetId(0),
            &Viewport {
                x: 0.0,
                y: 0.0,
                width: 400.0,
                height: 120.0,
            },
        );

        assert_eq!(dl.hyperlinks.len(), 2);
        assert_eq!(dl.hyperlinks[0].right, 1);
        assert_eq!(dl.hyperlinks[0].tooltip.as_deref(), Some("Open site"));
        let text = dl
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCmd::Text {
                    text,
                    color,
                    underline,
                    ..
                } => Some((text.as_str(), color.as_str(), *underline)),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert!(text.contains(&("Website", HYPERLINK_COLOR, true)));
        assert!(text.contains(&("Jump", HYPERLINK_COLOR, true)));
    }

    #[test]
    fn grid_meta_covers_visible_boundaries() {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Sheet1"));
        let dc = geometry::col_chars_to_px(geometry::DEFAULT_COL_WIDTH_CHARS);
        let dr = geometry::row_pt_to_px(geometry::DEFAULT_ROW_HEIGHT_PT);

        let vp = Viewport {
            x: dc * 1.5,
            y: dr * 2.5,
            width: dc * 2.0,
            height: dr * 1.0,
        };
        let dl = build_display_list(&wb, SheetId(0), &vp);

        assert_eq!(dl.grid.start_col, 1);
        assert_eq!(dl.grid.start_row, 2);
        assert_eq!(dl.grid.col_indices, None);
        assert_eq!(dl.grid.row_indices, None);
        assert_eq!(dl.grid.col_offsets.len(), 4);
        assert_eq!(dl.grid.row_offsets.len(), 3);
        assert!((dl.grid.col_offsets[0] - (dc * 1.0 - vp.x)).abs() < 0.01);
        assert!((dl.grid.row_offsets[0] - (dr * 2.0 - vp.y)).abs() < 0.01);
    }

    #[test]
    fn frozen_tracks_stay_pinned_while_the_body_scrolls() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.freeze_pane = Some(FreezePane::new(1, 1, CellRef::new(4, 3)));
        sheet.set_cell(CellRef::new(0, 0), text_cell("pinned"));
        sheet.set_cell(CellRef::new(4, 3), text_cell("body"));
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);
        let dc = geometry::col_chars_to_px(geometry::DEFAULT_COL_WIDTH_CHARS);
        let dr = geometry::row_pt_to_px(geometry::DEFAULT_ROW_HEIGHT_PT);
        let dl = build_display_list(
            &wb,
            SheetId(0),
            &Viewport {
                x: dc * 2.0,
                y: dr * 3.0,
                width: dc * 3.0,
                height: dr * 3.0,
            },
        );

        assert_eq!(dl.grid.start_col, 0);
        assert_eq!(dl.grid.start_row, 0);
        assert_eq!(dl.grid.col_indices.as_deref(), Some(&[0, 3, 4, 5][..]));
        assert_eq!(dl.grid.row_indices.as_deref(), Some(&[0, 4, 5, 6][..]));
        assert_eq!(
            dl.grid.col_offsets,
            vec![0.0, dc, dc * 2.0, dc * 3.0, dc * 4.0]
        );
        assert_eq!(
            dl.grid.row_offsets,
            vec![0.0, dr, dr * 2.0, dr * 3.0, dr * 4.0]
        );

        let clips = dl
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCmd::Text { text, clip, .. } => Some((text.as_str(), *clip)),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(
            clips[0],
            (
                "pinned",
                Rect {
                    x: 0.0,
                    y: 0.0,
                    w: dc,
                    h: dr
                }
            )
        );
        assert_eq!(
            clips[1],
            (
                "body",
                Rect {
                    x: dc,
                    y: dr,
                    w: dc,
                    h: dr
                }
            )
        );
        assert!(dl.commands.iter().any(|command| matches!(
            command,
            DrawCmd::Line { x1, x2, width, color, .. }
                if *x1 == dc && *x2 == dc && *width == PANE_DIVIDER_WIDTH
                    && color == PANE_DIVIDER_COLOR
        )));
        assert!(dl.commands.iter().any(|command| matches!(
            command,
            DrawCmd::Line { y1, y2, width, color, .. }
                if *y1 == dr && *y2 == dr && *width == PANE_DIVIDER_WIDTH
                    && color == PANE_DIVIDER_COLOR
        )));
    }

    fn ghost(row: u32, col: u32, old: &str, new: &str) -> GhostEdit {
        ghost_with_alignment_value(
            row,
            col,
            old,
            new,
            CellValue::Text {
                value: new.to_string(),
            },
        )
    }

    fn ghost_with_alignment_value(
        row: u32,
        col: u32,
        old: &str,
        new: &str,
        alignment_value: CellValue,
    ) -> GhostEdit {
        GhostEdit {
            row,
            col,
            old_text: old.into(),
            new_text: new.into(),
            alignment_value,
        }
    }

    fn text_cmds(dl: &DisplayList) -> Vec<(&str, &str, bool, Align)> {
        dl.commands
            .iter()
            .filter_map(|c| match c {
                DrawCmd::Text {
                    text,
                    color,
                    strike,
                    align,
                    ..
                } => Some((text.as_str(), color.as_str(), *strike, *align)),
                _ => None,
            })
            .collect()
    }

    fn ghost_flags(dl: &DisplayList) -> Vec<bool> {
        dl.commands
            .iter()
            .filter_map(|c| match c {
                DrawCmd::Text { ghost, .. } => Some(*ghost),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn ghost_pair_prefers_old_then_new_on_one_line() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(CellRef::new(0, 0), num_cell(10.0));
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);

        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: 200.0,
            height: 60.0,
        };
        let dl = build_display_list_with_ghosts(&wb, SheetId(0), &vp, &[ghost(0, 0, "10", "42")]);

        let texts = text_cmds(&dl);
        assert_eq!(texts.len(), 2);
        assert_eq!(texts[0], ("10", GHOST_DEL_COLOR, true, Align::Left));
        assert_eq!(texts[1], ("42", GHOST_INS_COLOR, false, Align::Left));
        assert_eq!(ghost_flags(&dl), vec![false, true]);

        let lines: Vec<(f32, f32, f32)> = dl
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCmd::Text {
                    x, y, font_size, ..
                } => Some((*x, *y, *font_size)),
                _ => None,
            })
            .collect();
        assert!(lines[0].0 < lines[1].0);
        assert_eq!(lines[0].1, lines[1].1);
        assert_eq!((lines[0].2, lines[1].2), (FONT_SIZE_PT, FONT_SIZE_PT));
    }

    #[test]
    fn ghost_insertion_paints_green_only() {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Sheet1"));

        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: 200.0,
            height: 60.0,
        };
        let dl = build_display_list_with_ghosts(
            &wb,
            SheetId(0),
            &vp,
            &[ghost_with_alignment_value(
                1,
                1,
                "",
                "7",
                CellValue::Number { value: 7.0 },
            )],
        );

        let texts = text_cmds(&dl);
        assert_eq!(texts, vec![("7", GHOST_INS_COLOR, false, Align::Right)]);
        assert_eq!(ghost_flags(&dl), vec![true]);
        let (x, clip) = dl
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCmd::Text { x, clip, .. } => Some((*x, *clip)),
                _ => None,
            })
            .unwrap();
        assert_eq!(x, clip.x + clip.w - TEXT_PAD_PX);
    }

    #[test]
    fn single_ghosts_honor_explicit_alignment_and_deleted_value_type() {
        let mut wb = Workbook::default();
        let style = wb
            .styles
            .intern_cell_format(&xlsx_model::CellFormat {
                alignment: xlsx_model::Alignment {
                    h: Some(HAlign::Left),
                    ..xlsx_model::Alignment::default()
                },
                ..xlsx_model::CellFormat::default()
            })
            .unwrap();
        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(
            CellRef::new(0, 0),
            Cell {
                style: Some(style),
                ..Cell::default()
            },
        );
        sheet.set_cell(CellRef::new(1, 0), num_cell(7.0));
        wb.sheets.push(sheet);

        let dl = build_display_list_with_ghosts(
            &wb,
            SheetId(0),
            &Viewport {
                x: 0.0,
                y: 0.0,
                width: 200.0,
                height: 60.0,
            },
            &[
                ghost_with_alignment_value(0, 0, "", "7", CellValue::Number { value: 7.0 }),
                ghost_with_alignment_value(1, 0, "7", "", CellValue::Number { value: 7.0 }),
            ],
        );

        assert_eq!(
            text_cmds(&dl),
            vec![
                ("7", GHOST_INS_COLOR, false, Align::Left),
                ("7", GHOST_DEL_COLOR, true, Align::Right),
            ]
        );
    }

    #[test]
    fn stacked_ghost_pair_puts_new_value_on_top() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(CellRef::new(0, 0), text_cell("previous long value"));
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);

        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: 200.0,
            height: 60.0,
        };
        let dl = build_display_list_with_ghosts(
            &wb,
            SheetId(0),
            &vp,
            &[ghost(0, 0, "previous long value", "replacement long value")],
        );

        let texts = text_cmds(&dl);
        assert_eq!(texts.len(), 2);
        assert!(texts[0].0.ends_with('…') && !texts[0].2);
        assert!(texts[1].0.ends_with('…') && texts[1].2);
        assert_eq!((texts[0].3, texts[1].3), (Align::Left, Align::Left));

        let lines: Vec<_> = dl
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCmd::Text { y, font_size, .. } => Some((*y, *font_size)),
                _ => None,
            })
            .collect();
        assert!(lines[0].0 < lines[1].0);
        assert_eq!((lines[0].1, lines[1].1), (FONT_SIZE_PT, FONT_SIZE_PT));
    }

    #[test]
    fn short_rows_shrink_stacked_ghosts_without_overlap() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.row_heights.insert(0, 7.5);
        sheet.set_cell(CellRef::new(0, 0), num_cell(10.0));
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);

        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: 200.0,
            height: 80.0,
        };
        let dl = build_display_list_with_ghosts(
            &wb,
            SheetId(0),
            &vp,
            &[ghost(0, 0, "previous", "replacement")],
        );

        let lines: Vec<(f32, f32)> = dl
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCmd::Text { y, font_size, .. } => Some((*y, *font_size)),
                _ => None,
            })
            .collect();
        let texts = text_cmds(&dl);
        assert!(!texts[0].2 && texts[1].2);
        assert_eq!(lines.len(), 2);
        assert_eq!(
            (lines[0].1, lines[1].1),
            (
                FONT_SIZE_PT * GHOST_MIN_SCALE,
                FONT_SIZE_PT * GHOST_MIN_SCALE
            )
        );
        assert!(
            lines[1].0 - lines[0].0
                >= FONT_SIZE_PT * GHOST_MIN_SCALE * (ASCENT_RATIO + DESCENT_RATIO) - 0.01
        );
    }

    #[test]
    fn equal_formatted_values_keep_the_committed_cell_text() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(CellRef::new(0, 0), num_cell(4855.0));
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);

        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: 200.0,
            height: 60.0,
        };
        let dl =
            build_display_list_with_ghosts(&wb, SheetId(0), &vp, &[ghost(0, 0, "4855", "4855")]);

        assert_eq!(
            text_cmds(&dl),
            vec![("4855", TEXT_COLOR, false, Align::Right)]
        );
        assert_eq!(ghost_flags(&dl), vec![false]);
        assert_eq!(
            dl.commands
                .iter()
                .filter(|command| matches!(command, DrawCmd::FillRect { .. }))
                .count(),
            1
        );
    }

    #[test]
    fn ghost_runs_carry_revision_highlights_and_new_underline() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(CellRef::new(0, 0), num_cell(10.0));
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);
        let dl = build_display_list_with_ghosts(
            &wb,
            SheetId(0),
            &Viewport {
                x: 0.0,
                y: 0.0,
                width: 200.0,
                height: 60.0,
            },
            &[ghost(0, 0, "10", "42")],
        );
        let styles: Vec<_> = dl
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCmd::Text {
                    highlight,
                    dashed_underline,
                    ..
                } => Some((highlight.as_deref(), *dashed_underline)),
                _ => None,
            })
            .collect();
        assert_eq!(
            styles,
            vec![
                (Some(GHOST_DEL_HIGHLIGHT), false),
                (Some(GHOST_INS_HIGHLIGHT), true)
            ]
        );
    }

    #[test]
    fn merge_anchor_draws_covered_cells_skip() {
        let mut sheet = Sheet::new("Sheet1");
        sheet
            .merges
            .push(CellRange::new(CellRef::new(0, 0), CellRef::new(0, 1)));
        sheet.set_cell(CellRef::new(0, 0), text_cell("merged"));
        sheet.set_cell(CellRef::new(0, 1), text_cell("covered"));
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);

        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: 400.0,
            height: 100.0,
        };
        let dl = build_display_list(&wb, SheetId(0), &vp);

        let texts: Vec<_> = dl
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCmd::Text { text, clip, .. } => Some((text.clone(), *clip)),
                _ => None,
            })
            .collect();
        assert_eq!(texts.len(), 1);
        assert_eq!(texts[0].0, "merged");
        let dc = geometry::col_chars_to_px(geometry::DEFAULT_COL_WIDTH_CHARS);
        assert!((texts[0].1.w - dc * 2.0).abs() < 0.01);
    }

    fn pie_chart() -> ooxml_drawingml::chart::ChartSpace {
        use ooxml_drawingml::chart::{ChartSeries, ChartSpace};
        ChartSpace {
            chart_type: "pie".to_owned(),
            title: Some("Sales".to_owned()),
            legend: None,
            series: vec![ChartSeries {
                name: Some("Revenue".to_owned()),
                categories: vec!["North".to_owned(), "South".to_owned()],
                values: vec![10.0, 25.0],
                color: "#4472C4".to_owned(),
                index: None,
                order: None,
                category_formula: None,
                value_formula: None,
                axis_ids: None,
                points: None,
                grouping: None,
                marker: None,
                smooth: None,
            }],
            axes: None,
            plot_groups: Vec::new(),
            axis_list: None,
        }
    }

    fn anchored_chart(from_col: u32, from_row: u32, to_col: u32, to_row: u32) -> SheetDrawing {
        SheetDrawing {
            anchor: DrawingAnchor::Cell {
                from: AnchorCell {
                    col: from_col,
                    col_offset_emu: 0,
                    row: from_row,
                    row_offset_emu: 0,
                },
                to: Some(AnchorCell {
                    col: to_col,
                    col_offset_emu: 0,
                    row: to_row,
                    row_offset_emu: 0,
                }),
                extent_emu: None,
            },
            chart: pie_chart(),
            created: false,
        }
    }

    fn clip_span(dl: &DisplayList) -> Option<(usize, usize)> {
        let push = dl
            .commands
            .iter()
            .position(|c| matches!(c, DrawCmd::PushClip { .. }))?;
        let pop = dl
            .commands
            .iter()
            .position(|c| matches!(c, DrawCmd::PopClip {}))?;
        Some((push, pop))
    }

    #[test]
    fn charts_paint_paths_inside_a_clip_pair() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(CellRef::new(0, 0), num_cell(1.0));
        sheet.drawings.push(anchored_chart(1, 1, 6, 12));
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);

        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: 800.0,
            height: 400.0,
        };
        let dl = build_display_list(&wb, SheetId(0), &vp);

        let (push, pop) = clip_span(&dl).expect("chart must emit a clip pair");
        assert!(push < pop);
        let inside = &dl.commands[push + 1..pop];
        assert!(inside.iter().any(|c| matches!(c, DrawCmd::Path { .. })));
        assert!(inside.iter().any(|c| matches!(c, DrawCmd::FillRect { .. })));
        assert!(
            inside
                .iter()
                .any(|c| matches!(c, DrawCmd::Text { text, .. } if text == "Sales"))
        );
        assert!(
            !dl.commands[..push]
                .iter()
                .chain(&dl.commands[pop + 1..])
                .any(|c| matches!(c, DrawCmd::Path { .. })),
            "paths only paint inside the chart clip"
        );
    }

    #[test]
    fn charts_clip_to_the_body_pane_and_track_scroll() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(CellRef::new(0, 0), num_cell(1.0));
        sheet.freeze_pane = Some(FreezePane {
            rows: 2,
            cols: 1,
            top_left: CellRef::new(2, 1),
        });
        sheet.drawings.push(anchored_chart(2, 3, 8, 14));
        let geom = GridGeometry::new(&sheet);
        let divider_x = geom.col_x(1);
        let divider_y = geom.row_y(2);
        let chart_x = geom.col_x(2);
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);

        let vp = Viewport {
            x: 10.0,
            y: 25.0,
            width: 800.0,
            height: 400.0,
        };
        let dl = build_display_list(&wb, SheetId(0), &vp);

        let (push, pop) = clip_span(&dl).expect("chart must emit a clip pair");
        let DrawCmd::PushClip { x, y, w, h } = dl.commands[push] else {
            unreachable!()
        };
        assert_eq!(x, divider_x);
        assert_eq!(y, divider_y);
        assert_eq!(w, vp.width - divider_x);
        assert_eq!(h, vp.height - divider_y);

        let background = dl.commands[push + 1..pop]
            .iter()
            .find_map(|c| match c {
                DrawCmd::FillRect { x, .. } => Some(*x),
                _ => None,
            })
            .expect("chart background");
        assert!((background - (chart_x - vp.x)).abs() < 0.01);
    }

    #[test]
    fn chart_series_track_live_cell_values() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(CellRef::new(1, 0), text_cell("B\u{1eafc}"));
        sheet.set_cell(CellRef::new(2, 0), text_cell("Nam"));
        sheet.set_cell(CellRef::new(1, 1), num_cell(10.0));
        sheet.set_cell(CellRef::new(2, 1), num_cell(25.0));
        let mut drawing = anchored_chart(2, 1, 8, 14);
        drawing.chart.chart_type = "column".to_owned();
        let series = &mut drawing.chart.series[0];
        series.categories = vec!["Old1".to_owned(), "Old2".to_owned()];
        series.values = vec![1.0, 1.0];
        series.category_formula = Some("Sheet1!$A$2:$A$3".to_owned());
        series.value_formula = Some("Sheet1!$B$2:$B$3".to_owned());
        sheet.drawings.push(drawing);
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);

        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: 800.0,
            height: 400.0,
        };
        let dl = build_display_list(&wb, SheetId(0), &vp);
        let (push, pop) = clip_span(&dl).expect("chart must emit a clip pair");
        let labels: Vec<&str> = dl.commands[push + 1..pop]
            .iter()
            .filter_map(|c| match c {
                DrawCmd::Text { text, .. } => Some(text.as_str()),
                _ => None,
            })
            .collect();
        assert!(labels.contains(&"B\u{1eafc}"), "{labels:?}");
        assert!(!labels.contains(&"Old1"), "{labels:?}");
    }

    #[test]
    fn off_screen_charts_are_culled() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.drawings.push(anchored_chart(40, 0, 46, 12));
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);

        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: 400.0,
            height: 300.0,
        };
        let dl = build_display_list(&wb, SheetId(0), &vp);
        assert!(clip_span(&dl).is_none());
    }
}
