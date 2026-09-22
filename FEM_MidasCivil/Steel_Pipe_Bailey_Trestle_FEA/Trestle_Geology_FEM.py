"""地质信息识别与参数设置插件 — 主入口与 UI (tkinter版本)"""
from __future__ import annotations

# 1. 标准库
import copy
import math
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Callable

# 2. 第三方库
import ezdxf
import openpyxl
from openpyxl.styles import Alignment
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backend_bases import ResizeEvent
from matplotlib.figure import Figure
from matplotlib.patches import Polygon as MplPolygon
from ezdxf.entities import DXFGraphic
from ezdxf.math import Vec2
from shapely.affinity import translate
from shapely.geometry import LineString, MultiLineString, Point, Polygon
from shapely.ops import polygonize, unary_union


# ═══════════════════════════════════════════════════════════════════════
# 数据模型
# ═══════════════════════════════════════════════════════════════════════

@dataclass
class ExtractionResult:
    """DXF 数据提取结果"""
    boundary_lines: list[LineString] = field(default_factory=list)
    water_level_lines: list[LineString] = field(default_factory=list)
    water_level_elevation: float = 0.0
    y_origin: float = 0.0
    soil_lines: list[LineString] = field(default_factory=list)
    x_range: tuple[float, float] = (0.0, 0.0)


@dataclass
class PolygonInfo:
    """识别出的土层多边形信息"""
    polygon: Polygon = field(default_factory=lambda: Polygon())
    centroid_x: float = 0.0
    centroid_y: float = 0.0
    system_id: str = ""
    color: str = ""
    exterior_coords: list[tuple[float, float]] = field(default_factory=list)


@dataclass
class SoilParamRecord:
    """土层参数记录"""
    library_id: str = ""
    soil_name: str = ""
    unit_weight: float = 0.0
    cohesion: float = 0.0
    friction_angle: float = 0.0
    qsik: float = 0.0
    qpk: float = 0.0


@dataclass
class SaveResult:
    """保存结果"""
    saved: list[SoilParamRecord] = field(default_factory=list)
    skipped: list[SoilParamRecord] = field(default_factory=list)
    conflicts: list[tuple[SoilParamRecord, SoilParamRecord]] = field(default_factory=list)


# ═══════════════════════════════════════════════════════════════════════
# 颜色常量
# ═══════════════════════════════════════════════════════════════════════

LAYER_COLORS = [
    '#E6B3B3', '#B3D9E6', '#B3E6CC', '#E6D9B3', '#D9B3E6', '#B3E6E6',
    '#E6CCB3', '#CCE6B3', '#E6B3D9', '#B3B3E6', '#E6E6B3', '#B3E6B3',
    '#E6B3A0', '#D9E6B3', '#B3CCE6', '#E6B3CC',
]
HIGHLIGHT_COLOR = '#FFD700'


def get_layer_color(index: int) -> str:
    return LAYER_COLORS[index % len(LAYER_COLORS)]


# DXF 提取
# ═══════════════════════════════════════════════════════════════════════

DXF_SCALE = 0.001  # mm → m

DEFAULT_CONFIG = {
    "boundary_line": {"layer_names": ["Defpoints", "DEFPOINTS", "BOUNDARY"], "colors": [7]},
    "water_level_line": {"layer_names": ["5文本", "WATER", "水位"], "colors": [5]},
    "snap_tolerance": 0.01, "min_area_ratio": 0.01,
}


def _entity_to_segments(entity: DXFGraphic) -> list[LineString]:
    segments: list[LineString] = []
    dtype = entity.dxftype()

    if dtype == 'LINE':
        s, e = Vec2(entity.dxf.start), Vec2(entity.dxf.end)
        if s.distance(e) > 1e-9:
            segments.append(LineString([(s.x * DXF_SCALE, s.y * DXF_SCALE),
                                        (e.x * DXF_SCALE, e.y * DXF_SCALE)]))
    elif dtype == 'LWPOLYLINE':
        pts = [Vec2(p[:2]) for p in entity.get_points(format='xy')]
        for i in range(len(pts) - 1):
            if pts[i].distance(pts[i + 1]) > 1e-9:
                segments.append(LineString([(pts[i].x * DXF_SCALE, pts[i].y * DXF_SCALE),
                                            (pts[i + 1].x * DXF_SCALE, pts[i + 1].y * DXF_SCALE)]))
        if entity.closed and len(pts) > 2 and pts[-1].distance(pts[0]) > 1e-9:
            segments.append(LineString([(pts[-1].x * DXF_SCALE, pts[-1].y * DXF_SCALE),
                                        (pts[0].x * DXF_SCALE, pts[0].y * DXF_SCALE)]))
    elif dtype == 'POLYLINE':
        verts = list(entity.vertices)
        for i in range(len(verts) - 1):
            p1, p2 = Vec2(verts[i].dxf.location), Vec2(verts[i + 1].dxf.location)
            if p1.distance(p2) > 1e-9:
                segments.append(LineString([(p1.x * DXF_SCALE, p1.y * DXF_SCALE),
                                            (p2.x * DXF_SCALE, p2.y * DXF_SCALE)]))
        if entity.is_closed and len(verts) > 2:
            pl, pf = Vec2(verts[-1].dxf.location), Vec2(verts[0].dxf.location)
            if pl.distance(pf) > 1e-9:
                segments.append(LineString([(pl.x * DXF_SCALE, pl.y * DXF_SCALE),
                                            (pf.x * DXF_SCALE, pf.y * DXF_SCALE)]))
    elif dtype == 'ARC':
        segments.extend(_arc_to_segments(entity))
    elif dtype == 'CIRCLE':
        segments.extend(_circle_to_segments(entity))
    elif dtype in ('XLINE', 'RAY'):
        segments.extend(_xline_to_segments(entity))
    return segments


def _arc_to_segments(entity: DXFGraphic, n: int = 36) -> list[LineString]:
    c = Vec2(entity.dxf.center); r = entity.dxf.radius
    sa = math.radians(entity.dxf.start_angle)
    ea = math.radians(entity.dxf.end_angle)
    if ea < sa: ea += 2 * math.pi
    pts = [((c.x + r * math.cos(sa + (ea - sa) * i / n)) * DXF_SCALE,
            (c.y + r * math.sin(sa + (ea - sa) * i / n)) * DXF_SCALE) for i in range(n + 1)]
    return [LineString([pts[i], pts[i + 1]]) for i in range(len(pts) - 1)]


def _circle_to_segments(entity: DXFGraphic, n: int = 72) -> list[LineString]:
    c = Vec2(entity.dxf.center); r = entity.dxf.radius
    pts = [((c.x + r * math.cos(2 * math.pi * i / n)) * DXF_SCALE,
            (c.y + r * math.sin(2 * math.pi * i / n)) * DXF_SCALE) for i in range(n)]
    pts.append(pts[0])
    return [LineString([pts[i], pts[i + 1]]) for i in range(len(pts) - 1)]


def _xline_to_segments(entity: DXFGraphic) -> list[LineString]:
    start = Vec2(entity.dxf.start)
    try: direction = Vec2(entity.dxf.unit_vector)
    except AttributeError:
        try: direction = Vec2(Vec2(entity.dxf.second_point).x - start.x, Vec2(entity.dxf.second_point).y - start.y)
        except AttributeError: return []
    sx, sy = start.x * DXF_SCALE, start.y * DXF_SCALE
    dx, dy = direction.x, direction.y
    norm = (dx ** 2 + dy ** 2) ** 0.5
    if norm < 1e-9: return []
    dx, dy = dx / norm, dy / norm
    ext = 200.0
    return [LineString([(sx - dx * ext, sy - dy * ext), (sx + dx * ext, sy + dy * ext)])]


def _is_vertical(seg: LineString, tol: float = 5.0) -> bool:
    x0, y0 = seg.coords[0]; x1, y1 = seg.coords[-1]
    dx, dy = abs(x1 - x0), abs(y1 - y0)
    return dy > 1e-9 and math.degrees(math.atan2(dx, dy)) < tol


def _is_horizontal(seg: LineString, tol: float = 5.0) -> bool:
    x0, y0 = seg.coords[0]; x1, y1 = seg.coords[-1]
    dx, dy = abs(x1 - x0), abs(y1 - y0)
    return dx > 1e-9 and math.degrees(math.atan2(dy, dx)) < tol


class DXFDataExtractor:
    """从 DXF 文件提取边界线、水位线、土层线段"""

    def __init__(self, config: dict | None = None):
        self._config = config or DEFAULT_CONFIG

    def extract(self, filepath: str) -> ExtractionResult:
        doc = ezdxf.readfile(filepath)
        msp = doc.modelspace()

        y_origin = self._find_y_origin(msp)
        boundary_segments, water_segments, soil_segments = self._extract_by_layer(msp)

        if not boundary_segments:
            boundary_segments = self._detect_boundary_by_geometry(soil_segments)
            b_ids = set(id(s) for s in boundary_segments)
            soil_segments = [s for s in soil_segments if id(s) not in b_ids]

        dy = -y_origin
        self._offset_y(boundary_segments, dy)
        self._offset_y(water_segments, dy)
        self._offset_y(soil_segments, dy)
        water_elevation = water_segments[0].coords[0][1] if water_segments else 0.0
        x_min, x_max = self._calculate_x_range(boundary_segments, soil_segments)
        dx = -x_min
        self._offset_x(boundary_segments, dx)
        self._offset_x(water_segments, dx)
        self._offset_x(soil_segments, dx)
        # 平移后更新 x_min 和 x_max
        x_max += dx
        x_min = 0.0

        return ExtractionResult(
            boundary_lines=boundary_segments,
            water_level_lines=water_segments,
            water_level_elevation=water_elevation,
            y_origin=y_origin,
            soil_lines=soil_segments,
            x_range=(x_min, x_max),
        )

    def _find_y_origin(self, msp) -> float:
        for e in msp:
            if e.dxftype() in ('XLINE', 'RAY') and e.dxf.layer == '0':
                return e.dxf.start.y * DXF_SCALE
        return 0.0

    def _offset_y(self, segments: list[LineString], dy: float) -> None:
        for i, seg in enumerate(segments):
            segments[i] = LineString([(x, y + dy) for x, y in seg.coords])

    def _offset_x(self, segments: list[LineString], dx: float) -> None:
        for i, seg in enumerate(segments):
            segments[i] = LineString([(x + dx, y) for x, y in seg.coords])

    def _extract_by_layer(self, msp):
        b_names = set(n.upper() for n in self._config["boundary_line"]["layer_names"])
        w_names = set(n.upper() for n in self._config["water_level_line"]["layer_names"])
        w_colors = set(self._config["water_level_line"]["colors"])
        b_segs, w_segs, s_segs = [], [], []

        for entity in msp:
            if entity.dxftype() not in ('LINE', 'LWPOLYLINE', 'POLYLINE', 'ARC', 'CIRCLE', 'XLINE', 'RAY'):
                continue
            layer = entity.dxf.layer.upper()
            color = entity.dxf.color
            segs = _entity_to_segments(entity)

            if layer == '0':
                continue
            elif layer in b_names:
                b_segs.extend(segs)
            elif layer in w_names or color in w_colors:
                w_segs.extend(segs)
            else:
                s_segs.extend(segs)

        return self._separate_boundary(b_segs), w_segs, s_segs

    def _separate_boundary(self, raw: list[LineString]) -> list[LineString]:
        v = [s for s in raw if _is_vertical(s, 10.0)]
        h = [s for s in raw if _is_horizontal(s, 10.0)]
        # 如果边界段已由预处理分割（段数>3），返回全部
        if len(v) + len(h) > 3:
            return v + h
        v.sort(key=lambda s: s.length, reverse=True)
        h.sort(key=lambda s: s.length, reverse=True)
        return v[:2] + h[:1]

    def _detect_boundary_by_geometry(self, all_segs: list[LineString]) -> list[LineString]:
        v = [s for s in all_segs if _is_vertical(s, 10.0)]
        h = [s for s in all_segs if _is_horizontal(s, 10.0)]
        v.sort(key=lambda s: s.length, reverse=True)
        result = v[:2]
        if h:
            min_y = min(s.coords[0][1] for s in h)
            candidates = [s for s in h if abs(s.coords[0][1] - min_y) < 1.0]
            candidates.sort(key=lambda s: s.length, reverse=True)
            if candidates: result.append(candidates[0])
        return result

    def _calculate_x_range(self, boundary, soil) -> tuple[float, float]:
        all_x = [c[0] for seg in boundary + soil for c in seg.coords]
        return (min(all_x), max(all_x)) if all_x else (0.0, 0.0)


# ═══════════════════════════════════════════════════════════════════════
# 拓扑多边形化
# ═══════════════════════════════════════════════════════════════════════

def _flatten(geometry) -> list[LineString]:
    """将几何对象展平为线段列表"""
    if geometry.is_empty: return []
    if isinstance(geometry, LineString): return [geometry]
    if isinstance(geometry, MultiLineString): return list(geometry.geoms)
    segs = []
    if hasattr(geometry, 'geoms'):
        for g in geometry.geoms: segs.extend(_flatten(g))
    return segs


class TopologyPolygonizer:
    def __init__(self, snap_tolerance: float = 0.01, min_area_ratio: float = 0.01):
        self._snap = snap_tolerance
        self._min_area_ratio = min_area_ratio

    def polygonize(self, soil_lines, boundary_lines) -> list[Polygon]:
        if not soil_lines: return []
        segs = _flatten(unary_union(MultiLineString(soil_lines + boundary_lines)))
        if not segs: return []
        polygons = list(polygonize(segs))
        polygons = self._merge_false_edges(polygons, soil_lines, boundary_lines)
        return [p for p in polygons if p.area > self._min_threshold(polygons)]

    def _merge_false_edges(self, polygons, soil_lines, boundary_lines):
        if len(polygons) <= 1: return polygons
        real_union = MultiLineString(list(soil_lines) + list(boundary_lines))
        TOL = 0.1
        merged = list(polygons)
        changed = True
        while changed:
            changed = False; new_merged = []; used = set()
            for i in range(len(merged)):
                if i in used: continue
                pa = merged[i]
                for j in range(i + 1, len(merged)):
                    if j in used: continue
                    pb = merged[j]
                    if not pa.intersects(pb): continue
                    shared = pa.intersection(pb.boundary)
                    if shared.is_empty: continue
                    if shared.distance(real_union) >= TOL:
                        union_p = pa.union(pb)
                        if union_p.geom_type == 'Polygon' and union_p.is_valid:
                            pa = union_p; used.add(j); changed = True
                new_merged.append(pa); used.add(i)
            merged = new_merged
        return merged

    def _min_threshold(self, polygons):
        if not polygons: return 0
        areas = sorted([p.area for p in polygons])
        return areas[len(areas) // 2] * self._min_area_ratio


# ═══════════════════════════════════════════════════════════════════════
# 空间排序
# ═══════════════════════════════════════════════════════════════════════

class SpatialSorter:
    def __init__(self, y_tolerance: float = 0.5):
        self._y_tolerance = y_tolerance

    def sort_and_assign_ids(self, polygons: list[Polygon]) -> list[PolygonInfo]:
        if not polygons: return []
        infos = []
        for poly in polygons:
            c = poly.centroid
            infos.append(PolygonInfo(polygon=poly, centroid_x=c.x, centroid_y=c.y,
                                     exterior_coords=list(poly.exterior.coords)))
        infos.sort(key=lambda p: (-p.centroid_y, p.centroid_x))
        for i, info in enumerate(infos): info.system_id = f"L{i + 1}"
        return infos


# ═══════════════════════════════════════════════════════════════════════
# Excel 参数库
# ═══════════════════════════════════════════════════════════════════════

PARAM_SHEET = "土层参数库"
HEADERS = ["土层编号", "土层名称", "重度(kN/m3)", "黏聚力c(kPa)",
           "内摩擦角φ(°)", "桩极限侧阻力标准值qsik(kPa)", "桩极限端阻力标准值qpk(kPa)"]

_CENTER_ALIGN = Alignment(horizontal='center', vertical='center')


def _apply_center_alignment(wb):
    """遍历所有 sheet 的所有非空 cell，统一设置居中对齐（对齐 Trestle_Excel_io_FEM 写入风格）"""
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row,
                                min_col=1, max_col=ws.max_column):
            for cell in row:
                if cell.value is not None:
                    cell.alignment = _CENTER_ALIGN


class ExcelParamLibrary:

    def __init__(self, excel_path: Optional[str] = None):
        self._excel_path = excel_path
    @property
    def excel_path(self): return self._excel_path

    def load_library(self) -> list[SoilParamRecord]:
        if not Path(self._excel_path).exists(): return []
        try: wb = openpyxl.load_workbook(self._excel_path, data_only=True)
        except PermissionError: raise RuntimeError(f"Excel 文件被占用：{self._excel_path}")
        if PARAM_SHEET not in wb.sheetnames: return []
        ws = wb[PARAM_SHEET]
        records = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[1] is None: continue
            records.append(SoilParamRecord(
                library_id=str(row[0]) if row[0] is not None else "",
                soil_name=str(row[1]) if row[1] else "",
                unit_weight=float(row[2]) if row[2] is not None else 0.0,
                cohesion=float(row[3]) if row[3] is not None else 0.0,
                friction_angle=float(row[4]) if row[4] is not None else 0.0,
                qsik=float(row[5]) if row[5] is not None else 0.0,
                qpk=float(row[6]) if row[6] is not None else 0.0,
            ))
        return records

    def save_with_dedup(self, records: list[SoilParamRecord]) -> SaveResult:
        existing = {r.soil_name: r for r in self.load_library()}
        saved, skipped, conflicts = [], [], []
        seen_names: set[str] = set()  # 内部去重：同一土层只写一条
        for r in records:
            if r.soil_name in seen_names:
                skipped.append(r)  # 本次输入中重复的土层，跳过
                continue
            seen_names.add(r.soil_name)
            if r.soil_name in existing:
                if self._equal(r, existing[r.soil_name]): skipped.append(r)
                else: conflicts.append((r, existing[r.soil_name]))
            else: saved.append(r)
        self._write_records(saved)
        return SaveResult(saved=saved, skipped=skipped, conflicts=conflicts)

    def _write_records(self, records):
        if not records: return
        if not Path(self._excel_path).exists(): self._create_workbook()
        try: wb = openpyxl.load_workbook(self._excel_path)
        except PermissionError: raise RuntimeError(f"Excel 文件被占用：{self._excel_path}")
        if PARAM_SHEET not in wb.sheetnames:
            ws = wb.create_sheet(PARAM_SHEET)
            for ci, h in enumerate(HEADERS, 1): ws.cell(row=1, column=ci, value=h)
        else: ws = wb[PARAM_SHEET]
        # 找最后一个非空行作为编号起点
        last_row = 1
        for r in range(ws.max_row, 0, -1):
            if any(ws.cell(r, c).value is not None for c in range(1, 8)):
                last_row = r; break
        nr = last_row + 1
        for r in records:
            ws.cell(row=nr, column=1, value=nr - 1)
            ws.cell(row=nr, column=2, value=r.soil_name)
            ws.cell(row=nr, column=3, value=r.unit_weight)
            ws.cell(row=nr, column=4, value=r.cohesion)
            ws.cell(row=nr, column=5, value=r.friction_angle)
            ws.cell(row=nr, column=6, value=r.qsik)
            ws.cell(row=nr, column=7, value=r.qpk)
            nr += 1
        _apply_center_alignment(wb)
        wb.save(self._excel_path)

    def _create_workbook(self):
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = PARAM_SHEET
        for ci, h in enumerate(HEADERS, 1): ws.cell(row=1, column=ci, value=h)
        wb.save(self._excel_path)

    def save_topography(self, polygons: list, sample_interval: float = 0.5,
                        records: list[SoilParamRecord] | None = None,
                        ground_line: list[tuple[float, float]] | None = None,
                        sheet_name: str | None = None) -> int:
        if not polygons: return 0
        sn = sheet_name
        if not Path(self._excel_path).exists(): self._create_workbook()
        try: wb = openpyxl.load_workbook(self._excel_path)
        except PermissionError: raise RuntimeError(f"Excel 文件被占用：{self._excel_path}")
        if sn in wb.sheetnames:
            ws = wb[sn]
            # 清除第2行起的旧数据（保留第1行表头）
            for r in range(2, ws.max_row + 1):
                for c in range(1, ws.max_column + 1):
                    ws.cell(r, c).value = None
        else:
            ws = wb.create_sheet(sn)
            ws.cell(1, 1, "土层编号标签")
            ws.cell(1, 2, "土顶线坐标")

        # === 优化后的混合 X 采样逻辑 ===
        all_x = [c[0] for info in polygons for c in info.polygon.exterior.coords]
        x_min, x_max = min(all_x), max(all_x)
        # 强制将左边界 X 坐标平移到 0
        x_offset = x_min

        # 1. 基础步长网格（插值兜底）
        n = int(math.ceil((x_max - x_min) / sample_interval)) + 1
        x_samples_base = [x_min + i * sample_interval for i in range(n)]
        if x_samples_base[-1] > x_max: x_samples_base[-1] = x_max

        # 2. 核心优化：将多边形所有关键节点的 X 坐标加入集合
        x_set = set(round(x, 4) for x in x_samples_base)
        for info in polygons:
            for c in info.polygon.exterior.coords:
                x_set.add(round(c[0], 4))

        # 3. 排序并偏移：强制左边界 X=0
        x_samples = sorted(x - x_offset for x in x_set)

        # 从第2行写数据（第1行为表头）
        for ri, info in enumerate(polygons):
            ws.cell(row=ri + 2, column=1, value=info.library_id)
            # 偏移多边形坐标使左边界 X=0
            shifted_poly = translate(info.polygon, -x_offset, 0)
            for ci, xs in enumerate(x_samples, 2):
                ty = self._poly_top_at_x(shifted_poly, xs)
                ws.cell(row=ri + 2, column=ci, value=f"({round(xs, 4)},{round(ty, 4)})" if ty is not None else "/")

        # 计算全局底边 Y 并写入末行
        all_y = [c[1] for info in polygons for c in info.polygon.exterior.coords]
        y_bottom = round(min(all_y) - 0.5, 4)
        bottom_row = len(polygons) + 2
        ws.cell(bottom_row, 1, "计算底边线y坐标")
        ws.cell(bottom_row, 2, y_bottom)

        # 保存地面线
        if ground_line:
            gl_name = "地面线"
            if gl_name in wb.sheetnames: del wb[gl_name]
            ws_gl = wb.create_sheet(gl_name)
            ws_gl.cell(1, 1, "X(m)"); ws_gl.cell(1, 2, "Y(m)")
            for i_gl, (x_gl, y_gl) in enumerate(ground_line, 2):
                ws_gl.cell(i_gl, 1, round(x_gl, 4))
                ws_gl.cell(i_gl, 2, round(y_gl, 4))
        _apply_center_alignment(wb)
        wb.save(self._excel_path)
        return len(polygons)

    @staticmethod
    def _poly_top_at_x(poly: Polygon, x: float) -> float | None:
        """返回多边形在 x 处的最高 Y 值（土顶线标高）"""
        b = poly.bounds
        if x < b[0] - 1e-4 or x > b[2] + 1e-4:
            return None
        vline = LineString([(x, b[1] - 1), (x, b[3] + 1)])
        ix = poly.intersection(vline)
        if ix.is_empty:
            return None
        if ix.geom_type == 'Point':
            return ix.y
        if ix.geom_type == 'MultiPoint':
            return max(pt.y for pt in ix.geoms)
        if ix.geom_type == 'LineString':
            return max(c[1] for c in ix.coords)
        if ix.geom_type == 'MultiLineString':
            return max(max(c[1] for c in ln.coords) for ln in ix.geoms)
        return None

    @staticmethod
    def _equal(a: SoilParamRecord, b: SoilParamRecord) -> bool:
        return (a.soil_name == b.soil_name
                and abs(a.unit_weight - b.unit_weight) < 1e-6
                and abs(a.cohesion - b.cohesion) < 1e-6
                and abs(a.friction_angle - b.friction_angle) < 1e-6
                and abs(a.qsik - b.qsik) < 1e-6
                and abs(a.qpk - b.qpk) < 1e-6)

# ═══════════════════════════════════════════════════════════════════════
# 辅助控件

# ═══════════════════════════════════════════════════════════════════════
# UI 组件 (tkinter版本)
# ═══════════════════════════════════════════════════════════════════════


class ParamRowWidget(tk.Frame):
    """参数行控件"""
    def __init__(self, parent, idx: int, system_id: str, color: str, 
                 library_records: list[SoilParamRecord], **kwargs):
        super().__init__(parent, **kwargs)
        self._idx = idx
        self._system_id = system_id
        self._library_records = library_records
        self._is_highlighted = False
        self._suppress_reset = False
        
        # 回调函数
        self._on_click_callback: Callable | None = None
        self._on_library_selected_callback: Callable | None = None
        
        self._setup_ui()
    
    def _setup_ui(self):
        """设置UI布局"""
        self.configure(relief='solid', bd=1, bg='#ffffff')
        
        # 主布局
        main_frame = tk.Frame(self, bg='#ffffff')
        main_frame.pack(fill='x', padx=4, pady=4)
        
        # 编号（显示system_id，如L1、L2、L3）
        self._id_label = tk.Label(main_frame, text=self._system_id, width=9, 
                                  relief='solid', bd=1, bg='#ECF0F1', font=('微软雅黑', 9, 'bold'))
        self._id_label.pack(side='left', padx=4)
        
        # 土层名称（初始为空，用户可输入）
        self._name_entry = tk.Entry(main_frame, width=12, relief='solid', bd=1)
        self._name_entry.pack(side='left', padx=4)
        self._name_entry.insert(0, "")  # 初始为空
        
        # 重度
        self._unit_weight_entry = tk.Entry(main_frame, width=10, relief='solid', bd=1)
        self._unit_weight_entry.pack(side='left', padx=4)
        self._unit_weight_entry.insert(0, "")
        
        # 粘聚力
        self._cohesion_entry = tk.Entry(main_frame, width=10, relief='solid', bd=1)
        self._cohesion_entry.pack(side='left', padx=4)
        self._cohesion_entry.insert(0, "")
        
        # 内摩擦角
        self._friction_entry = tk.Entry(main_frame, width=10, relief='solid', bd=1)
        self._friction_entry.pack(side='left', padx=4)
        self._friction_entry.insert(0, "")
        
        # qsik
        self._qsik_entry = tk.Entry(main_frame, width=23, relief='solid', bd=1)
        self._qsik_entry.pack(side='left', padx=5)
        self._qsik_entry.insert(0, "")
        
        # qpk
        self._qpk_entry = tk.Entry(main_frame, width=23, relief='solid', bd=1)
        self._qpk_entry.pack(side='left', padx=5)
        self._qpk_entry.insert(0, "")
        
        # 已有土层信息下拉框
        self._library_combo = ttk.Combobox(main_frame, width=16, state='readonly')
        self._library_combo.pack(side='left', padx=(7,0))
        self._library_combo['values'] = ["━ 选择已有土层 ━"] + [r.soil_name for r in self._library_records]
        self._library_combo.current(0)
        self._library_combo.bind('<<ComboboxSelected>>', self._on_library_change)
        
        # 绑定点击事件
        self.bind('<Button-1>', self._on_click)
        for child in self.winfo_children():
            child.bind('<Button-1>', self._on_click)

        # 用户修改参数 → 重置库选择（对齐 PySide6 版 textChanged；fill_from_record 有抑制）
        for entry in (self._name_entry, self._unit_weight_entry, self._cohesion_entry,
                      self._friction_entry, self._qsik_entry, self._qpk_entry):
            entry.bind('<KeyRelease>', self._on_edit_changed)
    
    def fill_from_record(self, r: SoilParamRecord):
        """从记录填充参数"""
        self._suppress_reset = True
        self._name_entry.delete(0, tk.END)
        self._name_entry.insert(0, r.soil_name)
        self._unit_weight_entry.delete(0, tk.END)
        self._unit_weight_entry.insert(0, str(r.unit_weight))
        self._cohesion_entry.delete(0, tk.END)
        self._cohesion_entry.insert(0, str(r.cohesion))
        self._friction_entry.delete(0, tk.END)
        self._friction_entry.insert(0, str(r.friction_angle))
        self._qsik_entry.delete(0, tk.END)
        self._qsik_entry.insert(0, str(r.qsik))
        self._qpk_entry.delete(0, tk.END)
        self._qpk_entry.insert(0, str(r.qpk))
        self._suppress_reset = False
    
    def _on_click(self, event):
        """点击事件处理"""
        if self._on_click_callback:
            self._on_click_callback(self._idx)
    
    def _on_library_change(self, event):
        """库选择变化事件"""
        idx = self._library_combo.current()
        if idx <= 0:
            return
        ri = idx - 1
        if 0 <= ri < len(self._library_records):
            r = self._library_records[ri]
            self.fill_from_record(r)
            if self._on_library_selected_callback:
                self._on_library_selected_callback(self._idx, r.library_id)

    def _on_edit_changed(self, event=None):
        """用户修改参数 → 重置库选择（对齐 PySide6 版）"""
        if not self._suppress_reset:
            self._library_combo.current(0)
    
    def set_on_click_callback(self, callback: Callable):
        """设置点击回调"""
        self._on_click_callback = callback
    
    def set_on_library_selected_callback(self, callback: Callable):
        """设置库选择回调"""
        self._on_library_selected_callback = callback
    
    def get_record(self) -> SoilParamRecord:
        """获取参数记录"""
        return SoilParamRecord(
            soil_name=self._name_entry.get(),
            unit_weight=self._get_float(self._unit_weight_entry.get()),
            cohesion=self._get_float(self._cohesion_entry.get()),
            friction_angle=self._get_float(self._friction_entry.get()),
            qsik=self._get_float(self._qsik_entry.get()),
            qpk=self._get_float(self._qpk_entry.get()),
        )
    
    def _get_float(self, value: str) -> float:
        """安全转换为浮点数"""
        try:
            return float(value) if value.strip() else 0.0
        except ValueError:
            return 0.0
    
    def has_empty_fields(self) -> bool:
        """检查是否有空字段（与 PySide6 版一致：全部 6 个参数）"""
        return (not self._name_entry.get().strip() or
                not self._unit_weight_entry.get().strip() or
                not self._cohesion_entry.get().strip() or
                not self._friction_entry.get().strip() or
                not self._qsik_entry.get().strip() or
                not self._qpk_entry.get().strip())
    
    def set_highlighted(self, highlighted: bool):
        """设置高亮状态"""
        self._is_highlighted = highlighted
        if highlighted:
            self.configure(bg='#FFF3CD')
            for child in self.winfo_children():
                if isinstance(child, tk.Frame):
                    child.configure(bg='#FFF3CD')
                elif isinstance(child, tk.Label):
                    child.configure(bg='#FFF3CD')
        else:
            self.configure(bg='#ffffff')
            for child in self.winfo_children():
                if isinstance(child, tk.Frame):
                    child.configure(bg='#ffffff')
                elif isinstance(child, tk.Label):
                    child.configure(bg='#ffffff')



class GeologicalCanvas(FigureCanvasTkAgg):
    """地质图绘制画布 (tkinter版本)

    继承 FigureCanvasTkAgg（与 PySide6 版继承 FigureCanvasQTAgg 结构一致），
    并覆盖两处后端行为以修复 Windows 高 DPI 下的图形裁切：

    - 禁用 _update_device_pixel_ratio（官方 DPI 修正）：它按物理像素重设
      canvas 请求尺寸（Windows 125% 缩放时 1480x581），而 tk 的 width/height
      是逻辑单位 → canvas 请求尺寸被撑大 1.25 倍，pack 压回容器后 photo
      （物理像素、1:1 显示）比画布大 → 图形被裁切、不居中。
    - 覆盖 resize()：matplotlib 默认用 Configure 事件的物理像素设置 figure，
      本实现以 winfo_width/height（逻辑像素）驱动，保证
      「画布尺寸 == figure 尺寸 == photo 尺寸」严格一致，图形填满图块区并居中。
    """
    def __init__(self, parent=None, **kwargs):
        self._fig = Figure(figsize=(6, 3), dpi=100)
        self._ax = self._fig.add_subplot(111)
        super().__init__(self._fig, master=parent)

        self._polygon_infos: list[PolygonInfo] = []
        self._extraction: ExtractionResult | None = None
        self._click_cb: Callable | None = None
        self._hover_cb: Callable | None = None
        self._wl_elevation: float | None = None
        self._highlighted_id = ""

        # 绑定鼠标事件
        self.mpl_connect('button_press_event', self._on_click)
        self.mpl_connect('motion_notify_event', self._on_hover)

    def _update_device_pixel_ratio(self, event=None):
        """禁用官方 DPI 修正：其把 canvas 请求尺寸设为物理像素，
        与 tk 逻辑坐标不一致，是 Windows 高 DPI 下图形被裁切的根源。"""
        pass

    def resize(self, event):
        """覆盖默认 resize：以逻辑像素（winfo_width）驱动 figure 与 photo，
        保持 画布 == figure == photo 严格一致（详见类注释）。"""
        w = self._tkcanvas.winfo_width()
        h = self._tkcanvas.winfo_height()
        if w <= 1 or h <= 1:
            return
        self.figure.set_size_inches(w / self.figure.dpi, h / self.figure.dpi, forward=False)
        self._tkcanvas.delete(self._tkcanvas_image_region)
        self._tkphoto.configure(width=w, height=h)
        self._tkcanvas_image_region = self._tkcanvas.create_image(w // 2, h // 2, image=self._tkphoto)
        ResizeEvent("resize_event", self)._process()
        self.draw_idle()
    
    def draw_water_level(self, elevation, x_min, x_max):
        """绘制水位线"""
        self._wl_elevation = elevation
        self._wl_x_min, self._wl_x_max = x_min, x_max
        self._render()
    
    def draw_geological_section(self, polygon_infos, extraction):
        """绘制地质剖面"""
        self._polygon_infos = polygon_infos
        self._extraction = extraction
        self._highlighted_id = ""
        self._render()
    
    def set_click_callback(self, cb):
        """设置点击回调"""
        self._click_cb = cb
    
    def set_hover_callback(self, cb):
        """设置悬停回调"""
        self._hover_cb = cb
    
    def highlight_polygon(self, sid):
        """高亮多边形"""
        self._highlighted_id = sid
        self._render()
    
    def _render(self):
        """渲染图形"""
        self._ax.clear()
        if not self._polygon_infos:
            self._ax.text(0.5, 0.5, '请导入 DXF 地形图', ha='center', va='center',
                          fontsize=14, transform=self._ax.transAxes, color='gray')
            self.draw()
            return

        for info in self._polygon_infos:
            coords = list(info.polygon.exterior.coords)
            hi = info.system_id == self._highlighted_id
            patch = MplPolygon(coords, closed=True, facecolor=info.color,
                               edgecolor=HIGHLIGHT_COLOR if hi else 'black',
                               linewidth=2.5 if hi else 0.8, alpha=0.9 if hi else 0.7)
            self._ax.add_patch(patch)
            self._ax.annotate(info.system_id, xy=(info.centroid_x, info.centroid_y),
                              fontsize=9, fontweight='bold', ha='center', va='center',
                              bbox=dict(boxstyle='round,pad=0.2', facecolor='white', alpha=0.85))

        if self._extraction and self._extraction.water_level_lines:
            wl_y = self._extraction.water_level_elevation
            for wl in self._extraction.water_level_lines:
                x, y = wl.xy
                self._ax.plot(x, y, color='#0066CC', linestyle='--', linewidth=1.5)
            max_x = max(x for wl in self._extraction.water_level_lines for x in wl.xy[0])
            self._ax.annotate(f'水位 {wl_y:.3f}m', xy=(max_x, wl_y),
                              fontsize=8, color='#0066CC', ha='right', va='bottom')
        elif self._wl_elevation is not None:
            wl_y = self._wl_elevation
            xs = getattr(self, '_wl_x_min', -10), getattr(self, '_wl_x_max', 10)
            self._ax.plot(xs, (wl_y, wl_y), color='#0066CC', linestyle='--', linewidth=1.5)
            self._ax.annotate(f'水位 {wl_y:.3f}m', xy=(xs[1], wl_y),
                              fontsize=8, color='#0066CC', ha='right', va='bottom')

        if self._extraction:
            for bl in self._extraction.boundary_lines:
                x, y = bl.xy
                self._ax.plot(x, y, color='black', linewidth=2)

        # 设置坐标轴范围
        all_bounds = [info.polygon.bounds for info in self._polygon_infos]
        min_x = min(b[0] for b in all_bounds)
        min_y = min(b[1] for b in all_bounds)
        max_x = max(b[2] for b in all_bounds)
        max_y = max(b[3] for b in all_bounds)
        
        dx = (max_x - min_x) * 0.05 if max_x > min_x else 1.0
        dy = (max_y - min_y) * 0.05 if max_y > min_y else 1.0
        
        self._ax.set_xlim(min_x - dx, max_x + dx)
        self._ax.set_ylim(min_y - dy, max_y + dy)

        self._ax.set_aspect('auto')
        self._ax.grid(True, alpha=0.2, linestyle=':', color='gray')
        self._ax.set_xlabel('X (m)', fontsize=10)
        self._ax.set_ylabel('标高 (m)', fontsize=10)
        wl_val = self._wl_elevation if self._wl_elevation is not None else (self._extraction.water_level_elevation if self._extraction else 0)
        self._ax.set_title(f'已识别 {len(self._polygon_infos)} 个土层 | 水位标高: {wl_val:.3f} m',
                           fontsize=11, fontweight='bold')
        self.draw()

    def _on_click(self, event):
        """点击事件处理"""
        if event.inaxes != self._ax or event.xdata is None or event.ydata is None:
            return
        pt = Point(event.xdata, event.ydata)
        for info in self._polygon_infos:
            if info.polygon.contains(pt):
                if self._click_cb:
                    self._click_cb(info.system_id)
                break

    def _on_hover(self, event):
        """悬停事件处理"""
        if event.inaxes != self._ax or event.xdata is None or event.ydata is None:
            return
        pt = Point(event.xdata, event.ydata)
        found = ""
        for info in self._polygon_infos:
            if info.polygon.contains(pt):
                found = info.system_id
                break
        if found != self._highlighted_id:
            self._highlighted_id = found
            self._render()
            if self._hover_cb and found:
                self._hover_cb(found)



class ParamEditorPanel(tk.Frame):
    """参数编辑面板"""
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self._rows: list[ParamRowWidget] = []
        self._row_clicked_callback: Callable | None = None
        self._library_selected_callback: Callable | None = None
        
        self._setup_ui()
    
    def _setup_ui(self):
        """设置UI布局"""
        # 标题
        sep_line = tb.Separator(master=self, orient=HORIZONTAL)
        sep_line.pack(fill=X)
        title = tk.Label(self, text="土层参数编辑表", font=('微软雅黑', 11, 'bold'),
                         bg='#2C3E50', fg='white', pady=8)
        title.pack(fill='x')
        
        # 滚动区域（垂直 + 水平滚动，保障窄屏/缩放时能访问最右侧"选择已有土层信息"列）
        self._scroll_canvas = tk.Canvas(self, bg='#ffffff')
        self._scrollbar = ttk.Scrollbar(self, orient='vertical', command=self._scroll_canvas.yview)
        self._hscrollbar = ttk.Scrollbar(self, orient='horizontal', command=self._scroll_canvas.xview)
        self._scroll_frame = tk.Frame(self._scroll_canvas, bg='#ffffff')
        
        self._scroll_frame.bind('<Configure>', 
            lambda e: self._scroll_canvas.configure(scrollregion=self._scroll_canvas.bbox('all')))
        
        self._scroll_canvas.create_window((0, 0), window=self._scroll_frame, anchor='nw')
        self._scroll_canvas.configure(yscrollcommand=self._scrollbar.set, xscrollcommand=self._hscrollbar.set)
        
        self._scrollbar.pack(side='right', fill='y')
        self._hscrollbar.pack(side='bottom', fill='x')
        self._scroll_canvas.pack(side='left', fill='both', expand=True)
        
        # 绑定鼠标滚轮（bind 而非 bind_all：widget 销毁后自动失效，避免 TclError）
        self._on_mousewheel = lambda e: self._scroll_canvas.yview_scroll(int(-1*(e.delta/120)), 'units')
        self._scroll_canvas.bind('<MouseWheel>', self._on_mousewheel)
        self._scroll_frame.bind('<MouseWheel>', self._on_mousewheel)
        
        # 表头放入 _scroll_frame 内部（顶部），随横向滚动与数据列对齐
        header_frame = tk.Frame(self._scroll_frame, bg='#ECF0F1', relief='solid', bd=1)
        header_frame.pack(fill='x', padx=0, pady=0)
        
        headers = [("土层编号", 80), ("土层名称", 100), ("重度(kN/m3)", 80), ("黏聚力c(kPa)", 80),
                   ("内摩擦角φ(°)", 80), ("桩极限侧阻力标准值qsik(kPa)", 180), ("桩极限端阻力标准值qpk(kPa)", 180), ("选择已有土层信息", 150)]
        
        for text, width in headers:
            label = tk.Label(header_frame, text=text, width=width//8,
                            font=('微软雅黑', 9, 'bold'), bg='#ECF0F1', fg='#2C3E50',
                            relief='solid', bd=1)
            label.pack(side='left', padx=0, pady=2)
    
    def generate_rows(self, polygon_infos, library_records=None):
        """生成参数行"""
        self.clear_rows()
        library_records = library_records or []
        for idx, info in enumerate(polygon_infos):
            row = ParamRowWidget(self._scroll_frame, idx, info.system_id, info.color, library_records)
            row.set_on_click_callback(self._on_click)
            row.set_on_library_selected_callback(self._on_library_selected)
            self._rows.append(row)
            row.pack(fill='x', padx=0, pady=0)
    
    def clear_rows(self):
        """清除所有行"""
        for r in self._rows:
            r.destroy()
        self._rows.clear()
    
    def get_all_records(self) -> list[SoilParamRecord]:
        """获取所有记录"""
        return [r.get_record() for r in self._rows]
    
    def highlight_row(self, idx):
        """高亮指定行（滚动到该行可见，对齐 PySide6 版 ensureWidgetVisible）"""
        for i, r in enumerate(self._rows):
            r.set_highlighted(i == idx)
        if 0 <= idx < len(self._rows):
            self._scroll_canvas.update_idletasks()
            bbox = self._scroll_canvas.bbox('all')
            if bbox and bbox[3] > 0:
                # 滚动到该行顶部附近（yview_moveto 的分数 = 行顶 y / 内容总高）
                row_y = self._rows[idx].winfo_y()
                self._scroll_canvas.yview_moveto(max(0, row_y - 10) / bbox[3])
    
    def _on_click(self, idx):
        """点击事件处理"""
        self.highlight_row(idx)
        if self._row_clicked_callback:
            self._row_clicked_callback(idx)
    
    def _on_library_selected(self, idx, library_id):
        """库选择事件处理"""
        if self._library_selected_callback:
            self._library_selected_callback(idx, library_id)
    
    def set_row_clicked_callback(self, callback: Callable):
        """设置行点击回调"""
        self._row_clicked_callback = callback
    
    def set_library_selected_callback(self, callback: Callable):
        """设置库选择回调"""
        self._library_selected_callback = callback
    
    @property
    def row_count(self) -> int:
        """获取行数"""
        return len(self._rows)
    
    def get_empty_rows(self) -> list[int]:
        """获取空行索引"""
        return [i + 1 for i, r in enumerate(self._rows) if r.has_empty_fields()]



class MainWindow(tk.Toplevel):
    """主窗口（独立运行或嵌入主程序均可用）

    使用 tk.Toplevel 而非 tb.Window：
    - tk.Toplevel 可接受 master 参数，作为子窗口嵌入主程序（对齐 PySide6 QMainWindow(parent)）
    - tb.Window 继承 tk.Tk（根窗口），不接受 master，独立运行时也不需要
    - ttkbootstrap 主题通过全局 tb.Style() 已在 geology_main / 主程序中初始化，ttk 控件样式不受影响
    """
    def __init__(self, parent=None, excel_path=None, sheet_name=None, water_level=None):
        super().__init__(master=parent)
        self._excel_path = excel_path
        self._sheet_name = sheet_name
        self._water_level = water_level
        self._extraction: ExtractionResult | None = None
        self._polygon_infos: list[PolygonInfo] = []
        self._library_records: list[SoilParamRecord] = []
        self._param_library = ExcelParamLibrary(self._excel_path)
        
        self._setup_window()
        self._setup_ui()
        self._setup_status_bar()
        self._load_library()
        self._load_from_excel()
    
    def _setup_window(self):
        """设置窗口属性"""
        self.title("地质信息识别与参数设置")
        # 按屏幕分辨率动态调整窗口大小，避免笔记本（窄屏）显示不全最右侧"选择已有土层信息"列
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        win_w = min(1200, int(sw * 0.9))
        win_h = min(1000, int(sh * 0.9))
        self.geometry(f"{win_w}x{win_h}")
        self.minsize(900, 600)
        self.resizable(True, True)
        
        # 设置窗口背景色
        self.configure(bg='#f5f5f5')
        
        # 设置关闭事件
        self.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def _setup_ui(self):
        """设置UI布局"""
        # 主容器
        main_frame = tk.Frame(self, bg='#f5f5f5')
        main_frame.pack(fill='both', expand=True, padx=8, pady=8)

        # 顶部：选择已有地形参数（水平通长：label + 下拉框占满整行）
        top_row = tk.Frame(main_frame, bg='#f5f5f5')
        top_row.pack(fill='x', pady=(0, 6))
        tk.Label(top_row, text="选择已有地形参数:", font=('微软雅黑', 10), bg='#f5f5f5').pack(side='left')
        self._topo_combo = ttk.Combobox(top_row, state="readonly", width=40)
        self._topo_combo.pack(side='left', padx=(6, 0), fill='x', expand=True)
        self._topo_combo.bind("<<ComboboxSelected>>", lambda e: self._on_topo_selected())
        self._reload_topo_list()
        
        # 底部按钮区域
        btn_frame = tk.Frame(main_frame, bg='#f5f5f5')
        btn_frame.pack(fill='x',side="bottom")
        # 保存按钮
        self._save_btn = tk.Button(btn_frame, text="保存", font=('微软雅黑', 10),
                                   bg='#27AE60', fg='white', relief='flat',
                                   command=self._on_save, width = 10)
        self._save_btn.pack(side='right', padx=5)
        # 取消按钮
        self._cancel_btn = tk.Button(btn_frame, text="取消", font=('微软雅黑', 10),
                                     bg='#95A5A6', fg='white', relief='flat',
                                     command=self._on_closing, width = 10)
        self._cancel_btn.pack(side='right', padx=5)


        # 导入按钮
        self._import_btn = tk.Button(main_frame, text="导入 DXF 地形图", 
                                     font=('微软雅黑', 11, 'bold'),
                                     bg='#2980B9', fg='white', relief='flat',
                                     height=2, command=self._on_import)
        self._import_btn.pack(fill='x', pady=(0, 6))
        
        # 地质信息图区（canvas + placeholder 切换）
        # 高度固定 400（pack_propagate(False) + fill='x' 不 expand，不随窗口拉伸）
        self._canvas_frame = tk.Frame(main_frame, bg='#f5f5f5', height=450)
        self._canvas_frame.pack_propagate(False)
        self._canvas_frame.pack(fill='x', pady=(0, 6))

        # 占位标签
        self._placeholder_lbl = tk.Label(self._canvas_frame, text="请导入 DXF 地形图",
                                         font=('微软雅黑', 14), fg='#888', bg='#f5f5f5',
                                         relief='solid', bd=2)
        self._placeholder_lbl.pack(fill='both', expand=True)

        # 地质图画布（导入/打开时切换显示：placeholder pack_forget + canvas pack）
        self._canvas = GeologicalCanvas(self._canvas_frame)
        
        # 参数编辑面板
        self._param_panel = ParamEditorPanel(main_frame)
        self._param_panel.pack(fill='both', expand=True, pady=(0, 6))
        
        # 设置回调
        self._canvas.set_click_callback(self._on_canvas_click)
        self._param_panel.set_row_clicked_callback(self._on_param_click)
        self._param_panel.set_library_selected_callback(self._on_library_selected)
        
        
        
        
    
    def _setup_status_bar(self):
        """设置状态栏"""
        self._status_bar = tk.Label(self, text="就绪", font=('微软雅黑', 9),
                                   bg='#ECF0F1', fg='#2C3E50', relief='sunken', anchor='w')
        self._status_bar.pack(side='bottom', fill='x')
    
    def _load_library(self):
        """加载土层参数库"""
        try:
            self._library_records = self._param_library.load_library()
        except Exception as e:
            messagebox.showerror("错误", f"加载土层参数库失败: {e}")
            self._library_records = []
    
    def _load_from_excel(self):
        """从 Excel 项目土层 Sheet 反推地质剖面并显示"""
        import openpyxl, re
        if not self._sheet_name or not Path(self._excel_path).exists():
            return
        wb = openpyxl.load_workbook(self._excel_path, data_only=True)
        if self._sheet_name not in wb.sheetnames:
            wb.close(); return
        ws = wb[self._sheet_name]

        # 逐行解析 (x,y) 坐标 + 参数（每行 = 一条边界线 + 可选参数）
        boundaries = []  # [(layer_id, [(x,y),...], SoilParamRecord|None)]
        saved_y_bottom = None
        for r in range(1, ws.max_row + 1):
            raw_a = ws.cell(r, 1).value
            if raw_a is not None and isinstance(raw_a, (int, float)) and raw_a < 0:
                saved_y_bottom = float(raw_a)
                continue
            coords = []
            sid = str(raw_a or "").strip()
            param_start = None
            for c in range(2, ws.max_column + 1):
                v = ws.cell(r, c).value
                if v is None or str(v).strip() in ('', '/'): continue
                m = re.match(r'\(\s*([\d.-]+)\s*,\s*([\d.-]+)\s*\)', str(v))
                if m: coords.append((float(m.group(1)), float(m.group(2))))
                else:
                    param_start = c
                    break
            if len(coords) >= 2:
                rec = None
                if param_start:
                    name_val = str(ws.cell(r, param_start).value or "")
                    rec = SoilParamRecord(
                        soil_name=name_val,
                        unit_weight=float(ws.cell(r, param_start + 1).value or 0),
                        cohesion=float(ws.cell(r, param_start + 2).value or 0),
                        friction_angle=float(ws.cell(r, param_start + 3).value or 0),
                        qsik=float(ws.cell(r, param_start + 4).value or 0),
                        qpk=float(ws.cell(r, param_start + 5).value or 0),
                    )
                boundaries.append((sid, coords, rec))

        # 读取土层参数库（Col A=ID, Col B=名称, Col C-G=参数）
        lib_map = {}
        try:
            if "土层参数库" in wb.sheetnames:
                ws_lib = wb["土层参数库"]
                for r in range(2, ws_lib.max_row + 1):
                    lib_id = str(ws_lib.cell(r, 1).value or "").strip()
                    if lib_id:
                        lib_map[lib_id] = SoilParamRecord(
                            soil_name=str(ws_lib.cell(r, 2).value or "").strip(),
                            unit_weight=float(ws_lib.cell(r, 3).value or 0),
                            cohesion=float(ws_lib.cell(r, 4).value or 0),
                            friction_angle=float(ws_lib.cell(r, 5).value or 0),
                            qsik=float(ws_lib.cell(r, 6).value or 0),
                            qpk=float(ws_lib.cell(r, 7).value or 0),
                        )
        except: pass
        wb.close()
        if len(boundaries) < 2:
            return

        # 1. 获取基础 X 范围和底边界
        all_y = [p[1] for _, pts, _ in boundaries for p in pts]
        y_bottom = saved_y_bottom if saved_y_bottom is not None else (min(all_y) - 0.5)

        base_x_set = set()
        for _, pts, _ in boundaries:
            for p in pts:
                base_x_set.add(p[0])

        common_x = sorted(list(base_x_set))
        if len(common_x) < 2: return

        # 定义一个纯 Python 的 1D 线性插值辅助函数
        def interpolate_y(x_val, pts):
            sorted_pts = sorted(pts, key=lambda p: p[0])
            xs = [p[0] for p in sorted_pts]
            ys = [p[1] for p in sorted_pts]
            if x_val < xs[0] or x_val > xs[-1]:
                return None
            for i in range(len(xs) - 1):
                x0, x1 = xs[i], xs[i+1]
                if x0 <= x_val <= x1:
                    if x0 == x1: return ys[i]
                    return ys[i] + (ys[i+1] - ys[i]) * (x_val - x0) / (x1 - x0)
            return None

        # 3. 按区间扫描，局部排序并生成梯形
        from collections import defaultdict
        layer_polygons = defaultdict(list)

        for k in range(len(common_x) - 1):
            x0 = common_x[k]
            x1 = common_x[k+1]
            x_mid = (x0 + x1) / 2.0
            
            if x1 - x0 < 1e-5: continue  # 忽略极小区间
            
            # 提取该区间内所有有效的土层顶线
            active_lines = []
            for sid, pts, rec in boundaries:
                y_mid = interpolate_y(x_mid, pts)
                if y_mid is not None:  # 如果是未延伸至此的透镜体，插值返回 None，自动被忽略
                    y0 = interpolate_y(x0, pts)
                    y1 = interpolate_y(x1, pts)
                    active_lines.append({'sid': sid, 'y0': y0, 'y_mid': y_mid, 'y1': y1})
            
            # 加入全局底边界作为保底线
            active_lines.append({'sid': 'BOTTOM', 'y0': y_bottom, 'y_mid': y_bottom, 'y1': y_bottom})
            
            # 【核心修正】在当前小区间内，按中点 Y 值进行真实的局部从上到下排序！不依赖全局平均值
            active_lines.sort(key=lambda item: -item['y_mid'])
            
            # 依次将相邻的两条线组成闭合梯形
            for i in range(len(active_lines) - 1):
                top_line = active_lines[i]
                bot_line = active_lines[i+1]
                
                # 如果两线重合（交点处），不生成面积
                if top_line['y_mid'] - bot_line['y_mid'] < 1e-5:
                    continue
                
                sid = top_line['sid']
                coords = [(x0, top_line['y0']), (x1, top_line['y1']),
                          (x1, bot_line['y1']), (x0, bot_line['y0'])]
                poly = Polygon(coords)
                if not poly.is_valid:
                    poly = poly.buffer(0)
                layer_polygons[sid].append(poly)

        # 4. 合并同一土层的所有碎片，提取有效多边形
        polygon_infos = []
        for sid, polys in layer_polygons.items():
            if not polys: continue
            
            # 将该土层的所有小梯形无缝合并为完整的大多边形
            merged = unary_union(polys)
            
            geom_list = []
            if merged.geom_type == 'Polygon': geom_list.append(merged)
            elif merged.geom_type == 'MultiPolygon': geom_list.extend(merged.geoms)
            
            for sub_geom in geom_list:
                if sub_geom.is_empty or sub_geom.area < 1e-4:
                    continue
                
                rec = None
                for s, _, r in boundaries:
                    if s == sid: rec = r; break
                if not rec and sid in lib_map:
                    rec = copy.deepcopy(lib_map[sid])
                    
                c = sub_geom.centroid
                info = PolygonInfo(polygon=sub_geom, centroid_x=c.x, centroid_y=c.y,
                                   system_id="", color="", exterior_coords=list(sub_geom.exterior.coords))
                info.library_id = sid
                polygon_infos.append(info)

        # 5. 最终的视觉空间排序（用于给图层顺次命名 L1, L2...）
        polygon_infos.sort(key=lambda p: (-p.centroid_y, p.centroid_x))
        
        # 6. 赋值与渲染更新
        for i, info in enumerate(polygon_infos):
            info.system_id = f"L{i + 1}"
            info.color = get_layer_color(i)

        self._polygon_infos = polygon_infos
        self._canvas.draw_geological_section(self._polygon_infos, self._extraction)
        # PySide6 版用 QStackedWidget.setCurrentWidget 切换；tkinter 版手动 pack 切换
        self._placeholder_lbl.pack_forget()
        self._canvas.get_tk_widget().pack(fill='both', expand=True)
        self._param_panel.generate_rows(self._polygon_infos, self._library_records)
        
        # 参数回填
        for idx, info in enumerate(self._polygon_infos):
            if info.library_id in lib_map and idx < self._param_panel.row_count:
                self._param_panel._rows[idx].fill_from_record(copy.deepcopy(lib_map[info.library_id]))

        # 绘制水位线（从主表传入）
        if self._water_level is not None:
            try:
                wl = float(self._water_level)
                if self._polygon_infos:
                    xs = [p[0] for info in self._polygon_infos for p in info.exterior_coords]
                    self._canvas.draw_water_level(wl, min(xs), max(xs))
            except (ValueError, TypeError):
                pass

    def _reload_topo_list(self):
        """列出工作簿中所有以 '地形表' 结尾的地形 sheet，并预选当前项目地形表"""
        names = []
        try:
            if self._excel_path and Path(self._excel_path).exists():
                wb = openpyxl.load_workbook(self._excel_path, read_only=True)
                names = [n for n in wb.sheetnames if n.endswith("地形表")]
                wb.close()
        except Exception:
            names = []
        self._topo_combo.configure(values=names)
        if self._sheet_name and self._sheet_name in names:
            self._topo_combo.set(self._sheet_name)

    def _on_topo_selected(self):
        """选择已有地形参数：切换到该 sheet，加载并重绘"""
        sel = self._topo_combo.get()
        if not sel or sel == self._sheet_name:
            return
        self._sheet_name = sel
        self._polygon_infos = []
        self._extraction = None
        # 先切回占位，避免旧图残留
        try:
            self._canvas.get_tk_widget().pack_forget()
        except Exception:
            pass
        self._placeholder_lbl.pack(fill='both', expand=True)
        self._param_panel.generate_rows([], self._library_records)
        self._load_from_excel()
        if not self._polygon_infos:
            print(f"[地质插件] '{sel}' 无有效地形数据，保持占位")

    def _on_import(self):
        """导入DXF文件"""
        from General.FilePath import suppress_dialog_stderr
        with suppress_dialog_stderr():
            filepath = filedialog.askopenfilename(
                title="选择DXF文件",
                filetypes=[("DXF files", "*.dxf"), ("All files", "*.*")]
            )
            if not filepath:
                return
        try:
            self._status_bar.configure(text="正在导入...")
            self.update()
            
            # 提取DXF数据
            self._extraction = DXFDataExtractor(DEFAULT_CONFIG).extract(filepath)
            if not self._extraction.soil_lines:
                messagebox.showerror("导入失败", "未识别到土层线段。")
                return

            # 生成多边形
            polygons = TopologyPolygonizer().polygonize(
                self._extraction.soil_lines, self._extraction.boundary_lines)

            if not polygons:
                messagebox.showerror("识别失败", "未能生成闭合多边形。")
                return

            # 排序并分配ID
            self._polygon_infos = SpatialSorter().sort_and_assign_ids(polygons)
            for i, info in enumerate(self._polygon_infos):
                info.color = get_layer_color(i)

            # 绘制地质剖面
            self._extraction.water_level_lines = []
            self._canvas.draw_geological_section(self._polygon_infos, self._extraction)

            # 切换显示画布
            self._placeholder_lbl.pack_forget()
            self._canvas.get_tk_widget().pack(fill='both', expand=True)

            # 生成参数行
            self._param_panel.generate_rows(self._polygon_infos, self._library_records)
            
            # 更新状态栏
            wl_str = f"水位: {float(self._water_level):.2f}m | " if self._water_level else ""
            self._status_bar.configure(text=f"已识别 {len(self._polygon_infos)} 个土层 | {wl_str}{filepath}")
            
        except Exception as e:
            messagebox.showerror("导入错误", str(e))
            self._status_bar.configure(text="导入失败")
    
    def _on_save(self):
        """保存数据"""
        # 前置校验
        if not self._param_panel.row_count:
            messagebox.showinfo("提示", "请先导入 DXF 或打开已有地形。")
            return
        
        empty = self._param_panel.get_empty_rows()
        if empty:
            messagebox.showinfo("提示", f"第 {', '.join(map(str, empty))} 行存在未填写的参数。")
            return
        
        if not self._polygon_infos or len(self._polygon_infos) == 0:
            messagebox.showinfo("提示", "无有效坐标矩阵，请先导入 DXF 文件。")
            return
        
        if not any(p.polygon.area > 0.01 for p in self._polygon_infos):
            messagebox.showinfo("提示", "坐标矩阵面积异常，请重新导入 DXF。")
            return
        
        all_records = self._param_panel.get_all_records()  # 完整列表（含重复，与 polygon_infos 等长）

        # 内部去重：同一土层名称只保留首条，用于写"土层参数库"sheet
        seen_names: set[str] = set()
        unique_records = []
        for r in all_records:
            if r.soil_name not in seen_names:
                seen_names.add(r.soil_name)
                unique_records.append(r)

        # 检查冲突（用去重后的列表比照库）
        try:
            existing_lib = {r.soil_name: r for r in self._param_library.load_library()}
        except:
            existing_lib = {}

        conflicts = [(r, existing_lib[r.soil_name]) for r in unique_records
                     if r.soil_name in existing_lib and not ExcelParamLibrary._equal(r, existing_lib[r.soil_name])]

        if conflicts:
            detail = [f"{r.soil_name}" for (r, _) in conflicts]
            messagebox.showwarning("提示", f"以下土层名称已存在: {', '.join(detail)}，请重命名后保存")
            return

        # 正式写入
        try:
            # 1) 土层参数库 sheet：去重写入
            self._param_library.save_with_dedup(unique_records)

            # 2) 回填 library_id 到 polygon_infos（按 soil_name 查表，而非索引对齐）
            fresh_lib = self._param_library.load_library()
            lib_by_name = {r.soil_name: r.library_id for r in fresh_lib}
            for ri, info in enumerate(self._polygon_infos):
                if ri < len(all_records):
                    name = all_records[ri].soil_name
                    if name in lib_by_name:
                        info.library_id = lib_by_name[name]

            # 3) 地形参数 sheet：全部 polygon_infos 逐行写入（含重复标签，正确行为）
            
            self._param_library.save_topography(
                self._polygon_infos,
                ground_line=getattr(self, "_ground_line", None),
                sheet_name=self._sheet_name
            )
            
            self._library_records = self._param_library.load_library()
            self._on_closing()
            
        except Exception as e:
            messagebox.showerror("保存错误", str(e))
    
    def _on_canvas_click(self, sid):
        """画布点击事件"""
        for i, info in enumerate(self._polygon_infos):
            if info.system_id == sid:
                self._param_panel.highlight_row(i)
                break
    
    def _on_param_click(self, idx):
        """参数行点击事件"""
        if 0 <= idx < len(self._polygon_infos):
            self._canvas.highlight_polygon(self._polygon_infos[idx].system_id)
    
    def _on_library_selected(self, idx, library_id):
        """库选择事件"""
        if 0 <= idx < len(self._polygon_infos):
            self._polygon_infos[idx].library_id = library_id
    
    def _on_closing(self):
        """关闭窗口"""
        self.destroy()



# ═══════════════════════════════════════════════════════════════════════
# 入口
# ═══════════════════════════════════════════════════════════════════════

def geology_main(root=None):
    """主函数（独立运行时创建隐藏 root；被主程序调用时传入已有 root）"""
    excel_path = sys.argv[1] if len(sys.argv) > 1 else None
    sheet_name = sys.argv[2] if len(sys.argv) > 2 else None
    water_level = sys.argv[3] if len(sys.argv) > 3 else None
    if root is None:
        root = tk.Tk()
        root.withdraw()
    app = MainWindow(root, excel_path, sheet_name, water_level)
    # 运行主循环
    app.mainloop()





