from __future__ import annotations

import math
import random
import time
from collections.abc import Iterator
from dataclasses import dataclass, field
from pathlib import Path
from threading import local
from typing import Any

from lugest_core.laser.quote_engine import analyze_dxf_geometry, merge_laser_quote_settings

try:
    from shapely import STRtree, affinity as shapely_affinity, make_valid as shapely_make_valid
    from shapely.geometry import GeometryCollection as ShapelyGeometryCollection
    from shapely.geometry import MultiPolygon as ShapelyMultiPolygon
    from shapely.geometry import Polygon as ShapelyPolygon
    from shapely.geometry import box as shapely_box
    from shapely.ops import unary_union as shapely_unary_union

    SHAPELY_AVAILABLE = True
except Exception:
    STRtree = None
    shapely_affinity = None
    shapely_make_valid = None
    ShapelyGeometryCollection = None
    ShapelyMultiPolygon = None
    ShapelyPolygon = None
    shapely_box = None
    shapely_unary_union = None
    SHAPELY_AVAILABLE = False


DEFAULT_SHEET_PROFILES: list[dict[str, Any]] = [
    {"name": "1000 x 2000", "width_mm": 1000.0, "height_mm": 2000.0},
    {"name": "1250 x 2500", "width_mm": 1250.0, "height_mm": 2500.0},
    {"name": "1500 x 3000", "width_mm": 1500.0, "height_mm": 3000.0},
    {"name": "2000 x 4000", "width_mm": 2000.0, "height_mm": 4000.0},
]

_NESTING_RUN_CONTEXT = local()


def _optimization_level() -> str:
    level = str(getattr(_NESTING_RUN_CONTEXT, "level", "tap1") or "tap1").strip().lower()
    return level if level in {"tap1", "tap2", "tap3"} else "tap1"


def _optimization_deadline_reached() -> bool:
    cancel_check = getattr(_NESTING_RUN_CONTEXT, "cancel_check", None)
    if callable(cancel_check) and bool(cancel_check()):
        return True
    deadline = float(getattr(_NESTING_RUN_CONTEXT, "deadline", 0.0) or 0.0)
    return deadline > 0.0 and time.monotonic() >= deadline


@dataclass
class NestItem:
    source_index: int
    path: str
    description: str
    ref_externa: str
    material: str
    thickness_mm: float
    qty: int
    bbox_width_mm: float
    bbox_height_mm: float
    net_area_mm2: float
    file_name: str
    geometry_warnings: tuple[str, ...] = field(default_factory=tuple)
    outer_polygons: tuple[tuple[tuple[float, float], ...], ...] = field(default_factory=tuple)
    hole_polygons: tuple[tuple[tuple[float, float], ...], ...] = field(default_factory=tuple)
    preview_paths: tuple[tuple[tuple[float, float], ...], ...] = field(default_factory=tuple)
    shape_source: str = "bbox"
    rotation_policy: str = "auto"
    priority: int = 0
    shape_cache_key: str = ""


def _as_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value or 0.0)
    except Exception:
        return float(default)


def _as_int(value: Any, default: int = 0) -> int:
    try:
        return int(round(float(value or 0)))
    except Exception:
        return int(default)


def _unique_texts(values: list[Any]) -> list[str]:
    seen: set[str] = set()
    ordered: list[str] = []
    for value in list(values or []):
        text = str(value or "").strip()
        if not text or text in seen:
            continue
        seen.add(text)
        ordered.append(text)
    return ordered


def _points_bbox(points: list[tuple[float, float]]) -> dict[str, float]:
    if not points:
        return {"min_x": 0.0, "min_y": 0.0, "max_x": 0.0, "max_y": 0.0, "width": 0.0, "height": 0.0}
    xs = [float(point[0]) for point in points]
    ys = [float(point[1]) for point in points]
    min_x = min(xs)
    min_y = min(ys)
    max_x = max(xs)
    max_y = max(ys)
    return {
        "min_x": round(min_x, 3),
        "min_y": round(min_y, 3),
        "max_x": round(max_x, 3),
        "max_y": round(max_y, 3),
        "width": round(max_x - min_x, 3),
        "height": round(max_y - min_y, 3),
    }


def _polygon_area_mm2(polygon: tuple[tuple[float, float], ...] | list[tuple[float, float]]) -> float:
    raw = list(polygon or [])
    if len(raw) < 3:
        return 0.0
    area = 0.0
    for index in range(len(raw)):
        x1, y1 = raw[index]
        x2, y2 = raw[(index + 1) % len(raw)]
        area += (float(x1) * float(y2)) - (float(x2) * float(y1))
    return abs(area) / 2.0


def _polygon_centroid_point(polygon: tuple[tuple[float, float], ...] | list[tuple[float, float]]) -> tuple[float, float]:
    raw = list(polygon or [])
    if len(raw) < 3:
        bbox = _points_bbox(raw)
        return ((bbox["min_x"] + bbox["max_x"]) / 2.0, (bbox["min_y"] + bbox["max_y"]) / 2.0)
    signed_area = 0.0
    centroid_x = 0.0
    centroid_y = 0.0
    for index in range(len(raw)):
        x1, y1 = raw[index]
        x2, y2 = raw[(index + 1) % len(raw)]
        cross = (float(x1) * float(y2)) - (float(x2) * float(y1))
        signed_area += cross
        centroid_x += (float(x1) + float(x2)) * cross
        centroid_y += (float(y1) + float(y2)) * cross
    if abs(signed_area) <= 1e-9:
        bbox = _points_bbox(raw)
        return ((bbox["min_x"] + bbox["max_x"]) / 2.0, (bbox["min_y"] + bbox["max_y"]) / 2.0)
    scale = 1.0 / (3.0 * signed_area)
    return (centroid_x * scale, centroid_y * scale)


def _explode_multi_part_shape(
    outer_polygons: tuple[tuple[tuple[float, float], ...], ...],
    hole_polygons: tuple[tuple[tuple[float, float], ...], ...],
) -> list[dict[str, Any]]:
    if len(list(outer_polygons or [])) <= 1:
        return []

    assigned_holes: dict[int, list[tuple[tuple[float, float], ...]]] = {index: [] for index in range(len(outer_polygons))}
    for hole_polygon in list(hole_polygons or []):
        probe_point = _polygon_centroid_point(hole_polygon)
        assigned_index = -1
        for outer_index, outer_polygon in enumerate(list(outer_polygons or [])):
            if _point_in_polygon(probe_point, outer_polygon):
                assigned_index = outer_index
                break
        if assigned_index < 0:
            first_point = tuple(list(hole_polygon or [probe_point])[0])
            for outer_index, outer_polygon in enumerate(list(outer_polygons or [])):
                if _point_in_polygon(first_point, outer_polygon):
                    assigned_index = outer_index
                    break
        if assigned_index >= 0:
            assigned_holes.setdefault(assigned_index, []).append(hole_polygon)

    exploded: list[dict[str, Any]] = []
    for outer_index, outer_polygon in enumerate(list(outer_polygons or []), start=1):
        component_holes = tuple(assigned_holes.get(outer_index - 1, []))
        component_points = list(outer_polygon)
        for hole_polygon in list(component_holes or []):
            component_points.extend(list(hole_polygon or []))
        component_bbox = _points_bbox(component_points)
        offset_x = component_bbox["min_x"]
        offset_y = component_bbox["min_y"]
        normalized_outer = _normalize_polygon(list(outer_polygon), offset_x=offset_x, offset_y=offset_y)
        normalized_holes = tuple(
            polygon
            for polygon in (
                _normalize_polygon(list(hole_polygon or []), offset_x=offset_x, offset_y=offset_y)
                for hole_polygon in list(component_holes or [])
            )
            if polygon
        )
        if not normalized_outer:
            continue
        net_area = max(
            0.0,
            _polygon_area_mm2(outer_polygon) - sum(_polygon_area_mm2(hole_polygon) for hole_polygon in list(component_holes or [])),
        )
        exploded.append(
            {
                "component_index": outer_index,
                "component_bbox": component_bbox,
                "bbox_width_mm": component_bbox["width"],
                "bbox_height_mm": component_bbox["height"],
                "net_area_mm2": round(net_area, 2),
                "outer_polygons": (normalized_outer,),
                "hole_polygons": normalized_holes,
            }
        )
    return exploded


def _explode_multi_part_preview_paths(
    preview_paths: tuple[tuple[tuple[float, float], ...], ...],
    components: list[dict[str, Any]],
) -> dict[int, tuple[tuple[tuple[float, float], ...], ...]]:
    assigned: dict[int, list[tuple[tuple[float, float], ...]]] = {index: [] for index in range(len(list(components or [])))}
    for path in list(preview_paths or []):
        path_bbox = _points_bbox(list(path or []))
        probe_point = (
            (path_bbox["min_x"] + path_bbox["max_x"]) / 2.0,
            (path_bbox["min_y"] + path_bbox["max_y"]) / 2.0,
        )
        selected_index = -1
        for component_index, component in enumerate(list(components or [])):
            component_bbox = dict(component.get("component_bbox", {}) or {})
            if not component_bbox:
                continue
            tol = 0.5
            if (
                path_bbox["min_x"] >= (_as_float(component_bbox.get("min_x", 0.0), 0.0) - tol)
                and path_bbox["min_y"] >= (_as_float(component_bbox.get("min_y", 0.0), 0.0) - tol)
                and path_bbox["max_x"] <= (_as_float(component_bbox.get("max_x", 0.0), 0.0) + tol)
                and path_bbox["max_y"] <= (_as_float(component_bbox.get("max_y", 0.0), 0.0) + tol)
            ):
                selected_index = component_index
                break
        if selected_index < 0:
            for component_index, component in enumerate(list(components or [])):
                outer_polygons = list(component.get("outer_polygons", ()) or ())
                if any(_point_in_polygon(probe_point, polygon) for polygon in outer_polygons):
                    selected_index = component_index
                    break
        if selected_index < 0:
            continue
        component_bbox = dict(components[selected_index].get("component_bbox", {}) or {})
        normalized = _normalize_path(
            list(path or []),
            offset_x=_as_float(component_bbox.get("min_x", 0.0), 0.0),
            offset_y=_as_float(component_bbox.get("min_y", 0.0), 0.0),
        )
        if normalized:
            assigned.setdefault(selected_index, []).append(normalized)
    return {index: tuple(paths) for index, paths in assigned.items()}


def _point_in_polygon(point: tuple[float, float], polygon: tuple[tuple[float, float], ...] | list[tuple[float, float]]) -> bool:
    raw = list(polygon or [])
    if len(raw) < 3:
        return False
    x, y = float(point[0]), float(point[1])
    bbox = _points_bbox(raw)
    if x < bbox["min_x"] or x > bbox["max_x"] or y < bbox["min_y"] or y > bbox["max_y"]:
        return False
    inside = False
    if raw[0] != raw[-1]:
        raw.append(raw[0])
    for index in range(len(raw) - 1):
        x1, y1 = raw[index]
        x2, y2 = raw[index + 1]
        if ((y1 > y) != (y2 > y)) and (x < ((x2 - x1) * (y - y1) / max(1e-12, (y2 - y1))) + x1):
            inside = not inside
    return inside


def _normalize_polygon(points: list[Any], *, offset_x: float = 0.0, offset_y: float = 0.0) -> tuple[tuple[float, float], ...]:
    out: list[tuple[float, float]] = []
    for point in list(points or []):
        if not isinstance(point, (list, tuple)) or len(point) < 2:
            continue
        x = round(_as_float(point[0], 0.0) - offset_x, 3)
        y = round(_as_float(point[1], 0.0) - offset_y, 3)
        candidate = (x, y)
        if out and out[-1] == candidate:
            continue
        out.append(candidate)
    if len(out) >= 2 and out[0] == out[-1]:
        out = out[:-1]
    if len(out) < 3:
        return ()
    return tuple(out)


def _normalize_path(points: list[Any], *, offset_x: float = 0.0, offset_y: float = 0.0) -> tuple[tuple[float, float], ...]:
    out: list[tuple[float, float]] = []
    for point in list(points or []):
        if not isinstance(point, (list, tuple)) or len(point) < 2:
            continue
        x = round(_as_float(point[0], 0.0) - offset_x, 3)
        y = round(_as_float(point[1], 0.0) - offset_y, 3)
        candidate = (x, y)
        if out and out[-1] == candidate:
            continue
        out.append(candidate)
    if len(out) < 2:
        return ()
    return tuple(out)


def _normalize_polygon_collection(value: Any) -> tuple[tuple[tuple[float, float], ...], ...]:
    if value is None or value == "":
        return ()
    raw = value
    if isinstance(raw, dict):
        raw = raw.get("points", raw.get("outer", raw.get("outer_polygons", [])))
    if not isinstance(raw, (list, tuple)):
        return ()
    raw_list = list(raw or [])
    if not raw_list:
        return ()
    if isinstance(raw_list[0], (list, tuple)) and len(raw_list[0]) >= 2 and not isinstance(raw_list[0][0], (list, tuple, dict)):
        raw_polygons = [raw_list]
    else:
        raw_polygons = raw_list
    polygons: list[tuple[tuple[float, float], ...]] = []
    for polygon_points in list(raw_polygons or []):
        polygon = _normalize_polygon(list(polygon_points or []))
        if polygon:
            polygons.append(polygon)
    if not polygons:
        return ()
    bbox = _points_bbox([point for polygon in polygons for point in polygon])
    return tuple(
        _normalize_polygon(list(polygon), offset_x=bbox["min_x"], offset_y=bbox["min_y"])
        for polygon in polygons
        if polygon
    )


def _normalize_path_collection(value: Any, *, offset_x: float = 0.0, offset_y: float = 0.0) -> tuple[tuple[tuple[float, float], ...], ...]:
    if value is None or value == "":
        return ()
    raw = value
    if isinstance(raw, dict):
        raw = raw.get("points", raw.get("paths", raw.get("cut_paths", [])))
    if not isinstance(raw, (list, tuple)):
        return ()
    paths: list[tuple[tuple[float, float], ...]] = []
    for path_points in list(raw or []):
        path = _normalize_path(list(path_points or []), offset_x=offset_x, offset_y=offset_y)
        if path:
            paths.append(path)
    return tuple(paths)


def _translate_polygons(polygons: tuple[tuple[tuple[float, float], ...], ...], dx: float, dy: float) -> list[list[tuple[float, float]]]:
    translated: list[list[tuple[float, float]]] = []
    for polygon in list(polygons or []):
        points = [(round(x + dx, 3), round(y + dy, 3)) for x, y in list(polygon or [])]
        if len(points) >= 3:
            translated.append(points)
    return translated


def _translate_paths(paths: tuple[tuple[tuple[float, float], ...], ...], dx: float, dy: float) -> list[list[tuple[float, float]]]:
    translated: list[list[tuple[float, float]]] = []
    for path in list(paths or []):
        points = [(round(x + dx, 3), round(y + dy, 3)) for x, y in list(path or [])]
        if len(points) >= 2:
            translated.append(points)
    return translated


def _rectangle_polygon(width_mm: float, height_mm: float) -> tuple[tuple[float, float], ...]:
    width = max(0.0, _as_float(width_mm, 0.0))
    height = max(0.0, _as_float(height_mm, 0.0))
    if width <= 0.0 or height <= 0.0:
        return ()
    return (
        (0.0, 0.0),
        (round(width, 3), 0.0),
        (round(width, 3), round(height, 3)),
        (0.0, round(height, 3)),
    )


def _point_on_segment(
    point: tuple[float, float],
    start: tuple[float, float],
    end: tuple[float, float],
    tol: float = 1e-3,
) -> bool:
    px, py = float(point[0]), float(point[1])
    x1, y1 = float(start[0]), float(start[1])
    x2, y2 = float(end[0]), float(end[1])
    seg_dx = x2 - x1
    seg_dy = y2 - y1
    seg_len_sq = (seg_dx * seg_dx) + (seg_dy * seg_dy)
    if seg_len_sq <= tol * tol:
        return math.hypot(px - x1, py - y1) <= tol
    projection = ((px - x1) * seg_dx + (py - y1) * seg_dy) / seg_len_sq
    if projection < -tol or projection > 1.0 + tol:
        return False
    nearest_x = x1 + (projection * seg_dx)
    nearest_y = y1 + (projection * seg_dy)
    return math.hypot(px - nearest_x, py - nearest_y) <= tol


def _point_on_polygon_boundary(
    point: tuple[float, float],
    polygon: tuple[tuple[float, float], ...] | list[tuple[float, float]],
    tol: float = 1e-3,
) -> bool:
    raw = list(polygon or [])
    if len(raw) < 2:
        return False
    for index in range(len(raw)):
        if _point_on_segment(point, raw[index], raw[(index + 1) % len(raw)], tol=tol):
            return True
    return False


def _point_in_solid_region(
    point: tuple[float, float],
    outer_polygons: tuple[tuple[tuple[float, float], ...], ...] | list[list[tuple[float, float]]],
    hole_polygons: tuple[tuple[tuple[float, float], ...], ...] | list[list[tuple[float, float]]],
    *,
    strict: bool = True,
    tol: float = 1e-3,
) -> bool:
    outer_list = list(outer_polygons or [])
    hole_list = list(hole_polygons or [])
    if strict and any(_point_on_polygon_boundary(point, polygon, tol=tol) for polygon in outer_list):
        return False
    if any(_point_in_polygon(point, polygon) for polygon in outer_list):
        if any(
            _point_in_polygon(point, polygon) or (strict and _point_on_polygon_boundary(point, polygon, tol=tol))
            for polygon in hole_list
        ):
            return False
        return True
    if not strict and any(_point_on_polygon_boundary(point, polygon, tol=tol) for polygon in outer_list):
        return not any(_point_in_polygon(point, polygon) for polygon in hole_list)
    return False


def _segment_cross(a: tuple[float, float], b: tuple[float, float], c: tuple[float, float]) -> float:
    return ((float(b[0]) - float(a[0])) * (float(c[1]) - float(a[1]))) - ((float(b[1]) - float(a[1])) * (float(c[0]) - float(a[0])))


def _segment_axis_overlap(
    a1: tuple[float, float],
    a2: tuple[float, float],
    b1: tuple[float, float],
    b2: tuple[float, float],
) -> float:
    if abs(float(a1[0]) - float(a2[0])) >= abs(float(a1[1]) - float(a2[1])):
        left = max(min(float(a1[0]), float(a2[0])), min(float(b1[0]), float(b2[0])))
        right = min(max(float(a1[0]), float(a2[0])), max(float(b1[0]), float(b2[0])))
    else:
        left = max(min(float(a1[1]), float(a2[1])), min(float(b1[1]), float(b2[1])))
        right = min(max(float(a1[1]), float(a2[1])), max(float(b1[1]), float(b2[1])))
    return max(0.0, right - left)


def _segments_overlap_interior(
    a1: tuple[float, float],
    a2: tuple[float, float],
    b1: tuple[float, float],
    b2: tuple[float, float],
    tol: float = 1e-3,
) -> bool:
    cross1 = _segment_cross(a1, a2, b1)
    cross2 = _segment_cross(a1, a2, b2)
    cross3 = _segment_cross(b1, b2, a1)
    cross4 = _segment_cross(b1, b2, a2)
    if abs(cross1) <= tol and abs(cross2) <= tol and abs(cross3) <= tol and abs(cross4) <= tol:
        return _segment_axis_overlap(a1, a2, b1, b2) > tol
    return (
        ((cross1 > tol and cross2 < -tol) or (cross1 < -tol and cross2 > tol))
        and ((cross3 > tol and cross4 < -tol) or (cross3 < -tol and cross4 > tol))
    )


def _polygon_edges_overlap(
    left_polygon: tuple[tuple[float, float], ...] | list[tuple[float, float]],
    right_polygon: tuple[tuple[float, float], ...] | list[tuple[float, float]],
    tol: float = 1e-3,
) -> bool:
    left = list(left_polygon or [])
    right = list(right_polygon or [])
    if len(left) < 2 or len(right) < 2:
        return False
    left_bbox = _points_bbox(left)
    right_bbox = _points_bbox(right)
    if (
        min(left_bbox["max_x"], right_bbox["max_x"]) < max(left_bbox["min_x"], right_bbox["min_x"]) - tol
        or min(left_bbox["max_y"], right_bbox["max_y"]) < max(left_bbox["min_y"], right_bbox["min_y"]) - tol
    ):
        return False
    for left_index in range(len(left)):
        a1 = left[left_index]
        a2 = left[(left_index + 1) % len(left)]
        for right_index in range(len(right)):
            b1 = right[right_index]
            b2 = right[(right_index + 1) % len(right)]
            if _segments_overlap_interior(a1, a2, b1, b2, tol=tol):
                return True
    return False


def _placement_geometry(
    placement: dict[str, Any],
) -> tuple[tuple[tuple[tuple[float, float], ...], ...], tuple[tuple[tuple[float, float], ...], ...]]:
    raw_outer = list(placement.get("shape_outer_polygons", []) or [])
    raw_holes = list(placement.get("shape_hole_polygons", []) or [])
    outer_polygons = tuple(
        polygon
        for polygon in (
            _normalize_polygon(list(points or []))
            for points in raw_outer
        )
        if polygon
    )
    hole_polygons = tuple(
        polygon
        for polygon in (
            _normalize_polygon(list(points or []))
            for points in raw_holes
        )
        if polygon
    )
    if outer_polygons:
        return outer_polygons, hole_polygons
    x_mm = _as_float(placement.get("x_mm", 0.0), 0.0)
    y_mm = _as_float(placement.get("y_mm", 0.0), 0.0)
    width_mm = _as_float(placement.get("width_mm", 0.0), 0.0)
    height_mm = _as_float(placement.get("height_mm", 0.0), 0.0)
    rect = _rectangle_polygon(width_mm, height_mm)
    if not rect:
        return (), ()
    translated = tuple((round(x_mm + x, 3), round(y_mm + y, 3)) for x, y in rect)
    return (translated,), ()


def _polygons_bbox(polygons: tuple[tuple[tuple[float, float], ...], ...] | list[list[tuple[float, float]]]) -> dict[str, float]:
    points = [tuple(point) for polygon in list(polygons or []) for point in list(polygon or [])]
    return _points_bbox(points)


def _bbox_overlap_mm(left_bbox: dict[str, float], right_bbox: dict[str, float]) -> tuple[float, float]:
    overlap_x = min(_as_float(left_bbox.get("max_x", 0.0), 0.0), _as_float(right_bbox.get("max_x", 0.0), 0.0)) - max(
        _as_float(left_bbox.get("min_x", 0.0), 0.0),
        _as_float(right_bbox.get("min_x", 0.0), 0.0),
    )
    overlap_y = min(_as_float(left_bbox.get("max_y", 0.0), 0.0), _as_float(right_bbox.get("max_y", 0.0), 0.0)) - max(
        _as_float(left_bbox.get("min_y", 0.0), 0.0),
        _as_float(right_bbox.get("min_y", 0.0), 0.0),
    )
    return max(0.0, overlap_x), max(0.0, overlap_y)


def _solid_regions_overlap(
    left_outer: tuple[tuple[tuple[float, float], ...], ...],
    left_holes: tuple[tuple[tuple[float, float], ...], ...],
    right_outer: tuple[tuple[tuple[float, float], ...], ...],
    right_holes: tuple[tuple[tuple[float, float], ...], ...],
    tol: float = 1e-3,
) -> bool:
    if not left_outer or not right_outer:
        return False
    left_bbox = _polygons_bbox(left_outer)
    right_bbox = _polygons_bbox(right_outer)
    overlap_x, overlap_y = _bbox_overlap_mm(left_bbox, right_bbox)
    if overlap_x <= tol or overlap_y <= tol:
        return False

    for left_polygon in list(left_outer or []) + list(left_holes or []):
        for right_polygon in list(right_outer or []) + list(right_holes or []):
            if _polygon_edges_overlap(left_polygon, right_polygon, tol=tol):
                return True

    for polygon in list(left_outer or []):
        for point in list(polygon or []):
            if _point_in_solid_region(point, right_outer, right_holes, strict=True, tol=tol):
                return True
    for polygon in list(right_outer or []):
        for point in list(polygon or []):
            if _point_in_solid_region(point, left_outer, left_holes, strict=True, tol=tol):
                return True
    return False


def _candidate_shape_conflicts(
    sheet: dict[str, Any],
    candidate_outer: tuple[tuple[tuple[float, float], ...], ...],
    candidate_holes: tuple[tuple[tuple[float, float], ...], ...],
) -> bool:
    for placement in list(sheet.get("placements", []) or []):
        placed_outer, placed_holes = _placement_geometry(dict(placement or {}))
        if _solid_regions_overlap(candidate_outer, candidate_holes, placed_outer, placed_holes):
            return True
    return False


def _placement_shapely_geometry(placement: dict[str, Any]) -> Any:
    raw_outer = tuple(placement.get("shape_outer_polygons", ()) or ())
    raw_holes = tuple(placement.get("shape_hole_polygons", ()) or ())
    if raw_outer:
        raw_points = [point for polygon in raw_outer for point in list(polygon or ())]
        geometry = _shapely_geometry_from_polygons(raw_outer, raw_holes)
        if geometry is not None and raw_points:
            min_x = min(float(point[0]) for point in raw_points)
            min_y = min(float(point[1]) for point in raw_points)
            return shapely_affinity.translate(geometry, xoff=min_x, yoff=min_y)
    x = _as_float(placement.get("x_mm", 0.0), 0.0)
    y = _as_float(placement.get("y_mm", 0.0), 0.0)
    width = max(0.0, _as_float(placement.get("width_mm", 0.0), 0.0))
    height = max(0.0, _as_float(placement.get("height_mm", 0.0), 0.0))
    if width <= 0.0 or height <= 0.0:
        return None
    return shapely_box(x, y, x + width, y + height)


def _sheet_overlap_diagnostics_shapely(sheet_row: dict[str, Any]) -> dict[str, Any]:
    placements = [dict(row or {}) for row in list(sheet_row.get("placements", []) or [])]
    geometries = [_placement_shapely_geometry(placement) for placement in placements]
    bbox_overlap_pairs = 0
    solid_overlap_pairs: list[tuple[str, str]] = []
    spacing_violation_pairs: list[tuple[str, str]] = []
    min_part_distance: float | None = None
    for index, left in enumerate(placements):
        left_geometry = geometries[index]
        if left_geometry is None:
            continue
        left_bounds = left_geometry.bounds
        for right_index in range(index + 1, len(placements)):
            right = placements[right_index]
            right_geometry = geometries[right_index]
            if right_geometry is None:
                continue
            right_bounds = right_geometry.bounds
            overlap_x = min(left_bounds[2], right_bounds[2]) - max(left_bounds[0], right_bounds[0])
            overlap_y = min(left_bounds[3], right_bounds[3]) - max(left_bounds[1], right_bounds[1])
            if overlap_x > 1e-3 and overlap_y > 1e-3:
                bbox_overlap_pairs += 1
            left_label = str(left.get("ref_externa", left.get("file_name", "")) or index + 1)
            right_label = str(right.get("ref_externa", right.get("file_name", "")) or right_index + 1)
            intersection_area = float(left_geometry.intersection(right_geometry).area)
            distance = float(left_geometry.distance(right_geometry))
            min_part_distance = distance if min_part_distance is None else min(min_part_distance, distance)
            if intersection_area > 1e-5:
                solid_overlap_pairs.append((left_label, right_label))
                continue
            required_spacing = max(
                _as_float(left.get("required_spacing_mm", 0.0), 0.0),
                _as_float(right.get("required_spacing_mm", 0.0), 0.0),
            )
            if distance < required_spacing - 1e-3:
                spacing_violation_pairs.append((left_label, right_label))

    width = max(0.0, _as_float(sheet_row.get("sheet_width_mm", 0.0), 0.0))
    height = max(0.0, _as_float(sheet_row.get("sheet_height_mm", 0.0), 0.0))
    edge_violations: list[str] = []
    min_edge_distance: float | None = None
    if width > 0.0 and height > 0.0:
        sheet_outer = tuple(sheet_row.get("sheet_outer_polygons", ()) or ())
        sheet_holes = tuple(sheet_row.get("sheet_hole_polygons", ()) or ())
        full_sheet = _shapely_geometry_from_polygons(
            sheet_outer,
            sheet_holes,
            fallback_width=width,
            fallback_height=height,
        )
        if full_sheet is None:
            full_sheet = shapely_box(0.0, 0.0, width, height)
        boundary = full_sheet.boundary
        for index, placement in enumerate(placements):
            geometry = geometries[index]
            if geometry is None:
                continue
            required_edge = max(0.0, _as_float(placement.get("required_edge_margin_mm", 0.0), 0.0))
            allowed = full_sheet.buffer(-required_edge, join_style="mitre") if required_edge > 0.0 else full_sheet
            edge_distance = float(geometry.distance(boundary))
            min_edge_distance = edge_distance if min_edge_distance is None else min(min_edge_distance, edge_distance)
            if not allowed.covers(geometry):
                edge_violations.append(str(placement.get("ref_externa", placement.get("file_name", "")) or index + 1))
    return {
        "bbox_overlap_pair_count": bbox_overlap_pairs,
        "solid_overlap_pair_count": len(solid_overlap_pairs),
        "solid_overlap_pairs": solid_overlap_pairs,
        "part_in_part_pair_count": max(0, bbox_overlap_pairs - len(solid_overlap_pairs)),
        "spacing_violation_pair_count": len(spacing_violation_pairs),
        "spacing_violation_pairs": spacing_violation_pairs,
        "edge_violation_count": len(edge_violations),
        "edge_violation_refs": edge_violations,
        "min_part_distance_mm": round(min_part_distance, 3) if min_part_distance is not None else None,
        "min_edge_distance_mm": round(min_edge_distance, 3) if min_edge_distance is not None else None,
        "validation_engine": "GEOS",
    }


def _sheet_overlap_diagnostics(sheet_row: dict[str, Any]) -> dict[str, Any]:
    if SHAPELY_AVAILABLE:
        return _sheet_overlap_diagnostics_shapely(sheet_row)
    placements = [dict(row or {}) for row in list(sheet_row.get("placements", []) or [])]
    bbox_overlap_pairs = 0
    solid_overlap_pairs: list[tuple[str, str]] = []
    for index, left in enumerate(placements):
        left_outer, left_holes = _placement_geometry(left)
        left_bbox = _polygons_bbox(left_outer)
        for right in placements[index + 1:]:
            right_outer, right_holes = _placement_geometry(right)
            right_bbox = _polygons_bbox(right_outer)
            overlap_x, overlap_y = _bbox_overlap_mm(left_bbox, right_bbox)
            if overlap_x > 1e-3 and overlap_y > 1e-3:
                bbox_overlap_pairs += 1
            if _solid_regions_overlap(left_outer, left_holes, right_outer, right_holes):
                solid_overlap_pairs.append(
                    (
                        str(left.get("ref_externa", left.get("file_name", "")) or "").strip() or str(index + 1),
                        str(right.get("ref_externa", right.get("file_name", "")) or "").strip() or str(index + 2),
                    )
                )
    return {
        "bbox_overlap_pair_count": bbox_overlap_pairs,
        "solid_overlap_pair_count": len(solid_overlap_pairs),
        "solid_overlap_pairs": solid_overlap_pairs,
        "part_in_part_pair_count": max(0, bbox_overlap_pairs - len(solid_overlap_pairs)),
    }


def _normalize_sheet_profile(row: dict[str, Any], index: int = 0) -> dict[str, Any] | None:
    payload = dict(row or {})
    width_mm = max(0.0, _as_float(payload.get("width_mm", 0.0), 0.0))
    height_mm = max(0.0, _as_float(payload.get("height_mm", 0.0), 0.0))
    outer_polygons = _normalize_polygon_collection(payload.get("outer_polygons", payload.get("contorno_points", payload.get("shape_points", ()))))
    hole_polygons = _normalize_polygon_collection(payload.get("hole_polygons", ()))
    if outer_polygons:
        contour_bbox = _points_bbox([point for polygon in outer_polygons for point in polygon])
        width_mm = max(width_mm, contour_bbox["width"])
        height_mm = max(height_mm, contour_bbox["height"])
    elif height_mm > width_mm:
        width_mm, height_mm = height_mm, width_mm
    if width_mm <= 0.0 or height_mm <= 0.0:
        return None
    name = str(payload.get("name", "") or "").strip() or f"Formato {index + 1}"
    source_kind = str(payload.get("source_kind", "purchase") or "purchase").strip().lower()
    if source_kind not in {"purchase", "stock", "retalho"}:
        source_kind = "purchase"
    source_label = str(payload.get("source_label", "") or "").strip() or name
    return {
        "name": name,
        "width_mm": round(width_mm, 3),
        "height_mm": round(height_mm, 3),
        "area_mm2": round(width_mm * height_mm, 2),
        "source_kind": source_kind,
        "source_label": source_label,
        "material_id": str(payload.get("material_id", "") or "").strip(),
        "lote": str(payload.get("lote", "") or "").strip(),
        "local": str(payload.get("local", "") or "").strip(),
        "quantity_available": max(0, _as_int(payload.get("quantity_available", 0), 0)),
        "is_retalho": bool(payload.get("is_retalho", source_kind == "retalho")),
        "p_compra": round(_as_float(payload.get("p_compra", 0.0), 0.0), 6),
        "peso_unid": round(_as_float(payload.get("peso_unid", 0.0), 0.0), 3),
        "outer_polygons": outer_polygons,
        "hole_polygons": hole_polygons,
    }


def _normalize_stock_sheet_candidate(row: dict[str, Any], index: int = 0) -> dict[str, Any] | None:
    payload = dict(row or {})
    quantity_available = max(
        0,
        _as_int(
            payload.get("quantity_available", payload.get("disponivel", payload.get("quantidade", 0))),
            0,
        ),
    )
    width_mm = max(0.0, _as_float(payload.get("width_mm", payload.get("largura", 0.0)), 0.0))
    height_mm = max(0.0, _as_float(payload.get("height_mm", payload.get("comprimento", 0.0)), 0.0))
    if quantity_available <= 0 or width_mm <= 0.0 or height_mm <= 0.0:
        return None
    source_kind = str(payload.get("source_kind", "retalho" if payload.get("is_retalho") else "stock") or "stock").strip().lower()
    if source_kind not in {"stock", "retalho"}:
        source_kind = "retalho" if bool(payload.get("is_retalho")) else "stock"
    lot_label = str(payload.get("lote", "") or payload.get("origem_lote", "") or "").strip()
    material_id = str(payload.get("material_id", "") or payload.get("id", "") or "").strip()
    dim_label = str(payload.get("dimensao", "") or "").strip() or f"{height_mm:g} x {width_mm:g}"
    name = str(payload.get("name", "") or "").strip() or f"{'Retalho' if source_kind == 'retalho' else 'Stock'} {lot_label or material_id or (index + 1)} | {dim_label}"
    source_label = str(payload.get("source_label", "") or "").strip() or f"{'Retalho' if source_kind == 'retalho' else 'Stock'} {lot_label or material_id or (index + 1)}"
    return _normalize_sheet_profile(
        {
            "name": name,
            "width_mm": width_mm,
            "height_mm": height_mm,
            "source_kind": source_kind,
            "source_label": source_label,
            "material_id": material_id,
            "lote": lot_label,
            "local": str(payload.get("local", "") or "").strip(),
            "quantity_available": quantity_available,
            "is_retalho": bool(payload.get("is_retalho", source_kind == "retalho")),
            "p_compra": payload.get("p_compra", 0.0),
            "peso_unid": payload.get("peso_unid", 0.0),
            "outer_polygons": payload.get("outer_polygons", payload.get("contorno_points", ())),
            "hole_polygons": payload.get("hole_polygons", ()),
        },
        index,
    )


def default_sheet_profiles(laser_settings: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    settings = merge_laser_quote_settings(laser_settings)
    raw_profiles = list(dict(settings.get("nesting", {}) or {}).get("sheet_profiles", []) or [])
    if not raw_profiles:
        raw_profiles = [dict(row) for row in DEFAULT_SHEET_PROFILES]
    profiles: list[dict[str, Any]] = []
    for index, row in enumerate(raw_profiles):
        profile = _normalize_sheet_profile(dict(row or {}), index)
        if profile is not None:
            profiles.append(profile)
    if profiles:
        return profiles
    fallback: list[dict[str, Any]] = []
    for index, row in enumerate(DEFAULT_SHEET_PROFILES):
        profile = _normalize_sheet_profile(row, index)
        if profile is not None:
            fallback.append(profile)
    return fallback


def default_nesting_options(laser_settings: dict[str, Any] | None = None) -> dict[str, Any]:
    settings = merge_laser_quote_settings(laser_settings)
    payload = dict(settings.get("nesting", {}) or {})
    return {
        "default_part_spacing_mm": max(0.0, _as_float(payload.get("default_part_spacing_mm", 8.0), 8.0)),
        "default_edge_margin_mm": max(0.0, _as_float(payload.get("default_edge_margin_mm", 8.0), 8.0)),
        "allow_rotate": bool(payload.get("allow_rotate", True)),
        "auto_select_sheet": bool(payload.get("auto_select_sheet", False)),
        "use_stock_first": bool(payload.get("use_stock_first", False)),
        "allow_purchase_fallback": bool(payload.get("allow_purchase_fallback", True)),
        "shape_aware": bool(payload.get("shape_aware", True)),
        "allow_mirror": bool(payload.get("allow_mirror", True)),
        # Free-angle search is expensive; keep it opt-in and let 0/90 rotation
        # cover the common production case by default.
        "free_angle_rotation": bool(payload.get("free_angle_rotation", False)),
        "shape_grid_mm": max(2.0, _as_float(payload.get("shape_grid_mm", 10.0), 10.0)),
        "common_line_estimate": bool(payload.get("common_line_estimate", True)),
        "common_line_tolerance_mm": max(0.0, _as_float(payload.get("common_line_tolerance_mm", 1.0), 1.0)),
        "lead_optimization": bool(payload.get("lead_optimization", True)),
        "lead_optimization_pct": max(0.0, min(50.0, _as_float(payload.get("lead_optimization_pct", 8.0), 8.0))),
        "optimization_level": (
            str(payload.get("optimization_level", "tap1") or "tap1").strip().lower()
            if str(payload.get("optimization_level", "tap1") or "tap1").strip().lower() in {"tap1", "tap2", "tap3"}
            else "tap1"
        ),
        "sheet_profiles": default_sheet_profiles(settings),
    }


def _row_path(row: dict[str, Any]) -> str:
    return str((row or {}).get("desenho", "") or "").strip()


def _row_description(row: dict[str, Any]) -> str:
    return str((row or {}).get("descricao", "") or "").strip() or Path(_row_path(row)).stem


def _row_ref(row: dict[str, Any]) -> str:
    return str((row or {}).get("ref_externa", "") or "").strip() or Path(_row_path(row)).stem


def _row_material(row: dict[str, Any]) -> str:
    return str((row or {}).get("material", "") or "").strip()


def _row_thickness_mm(row: dict[str, Any]) -> float:
    raw = str((row or {}).get("espessura", "") or "").strip().replace(",", ".")
    try:
        return float(raw)
    except Exception:
        return 0.0


def _row_qty(row: dict[str, Any]) -> int:
    return max(1, _as_int((row or {}).get("qtd", 1), 1))


def _row_rotation_policy(row: dict[str, Any]) -> str:
    raw = str(
        (row or {}).get(
            "nest_rotation_policy",
            (row or {}).get(
                "rotation_policy",
                (row or {}).get("rotation_mode", "auto"),
            ),
        )
        or "auto"
    ).strip().lower()
    if raw in {"0", "0°", "fixo_0", "forcar_0", "force_0", "fixed_0", "none", "sem"}:
        return "0"
    if raw in {"90", "90°", "fixo_90", "forcar_90", "force_90", "fixed_90"}:
        return "90"
    return "auto"


def _row_priority(row: dict[str, Any]) -> int:
    raw_value = (
        (row or {}).get(
            "nest_priority",
            (row or {}).get(
                "nesting_priority",
                (row or {}).get(
                    "priority",
                    (row or {}).get("prioridade", 0),
                ),
            ),
        )
    )
    if isinstance(raw_value, str):
        normalized = str(raw_value or "").strip().lower()
        if normalized in {"critica", "crítica", "critical", "urgent", "urgente"}:
            return 2
        if normalized in {"alta", "high", "prioritaria", "prioritária", "priority"}:
            return 1
        if normalized in {"baixa", "low"}:
            return -1
        if normalized in {"normal", ""}:
            return 0
    return max(-1, min(2, _as_int(raw_value, 0)))


def compatible_laser_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for row in list(rows or []):
        path = _row_path(row)
        if not path:
            continue
        if not Path(path).exists():
            continue
        if "corte laser" not in str((row or {}).get("operacao", "") or "").strip().lower():
            continue
        out.append(dict(row or {}))
    return out


def grouped_laser_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    groups: dict[tuple[str, float], list[dict[str, Any]]] = {}
    for row in compatible_laser_rows(rows):
        key = (_row_material(row), round(_row_thickness_mm(row), 3))
        groups.setdefault(key, []).append(row)
    ordered: list[dict[str, Any]] = []
    for (material, thickness_mm), group_rows in sorted(groups.items(), key=lambda item: (item[0][0], item[0][1])):
        ordered.append(
            {
                "key": f"{material}|{thickness_mm:g}",
                "label": f"{material or '-'} | {thickness_mm:g} mm",
                "material": material,
                "thickness_mm": thickness_mm,
                "rows": [dict(row) for row in group_rows],
            }
        )
    return ordered


def build_nesting_items(rows: list[dict[str, Any]], laser_settings: dict[str, Any] | None = None) -> tuple[list[NestItem], list[str]]:
    settings = merge_laser_quote_settings(laser_settings)
    layer_rules = dict(settings.get("layer_rules", {}) or {})
    items: list[NestItem] = []
    warnings: list[str] = []
    for index, row in enumerate(list(rows or [])):
        path = _row_path(row)
        if not path:
            continue
        try:
            geometry = analyze_dxf_geometry(path, layer_rules)
        except Exception as exc:
            warnings.append(f"{Path(path).name}: {exc}")
            continue
        bbox = dict(geometry.get("bbox_mm", {}) or {})
        metrics = dict(geometry.get("metrics", {}) or {})
        width = max(0.0, _as_float(bbox.get("width", 0.0), 0.0))
        height = max(0.0, _as_float(bbox.get("height", 0.0), 0.0))
        file_name = str(geometry.get("file_name", "") or Path(path).name)
        geometry_warnings = _unique_texts(list(geometry.get("warnings", []) or []))
        for warning in geometry_warnings:
            warnings.append(f"{file_name}: {warning}")
        net_area_mm2 = max(0.0, _as_float(metrics.get("net_area_mm2", 0.0), 0.0))
        bbox_min_x = _as_float(bbox.get("min_x", 0.0), 0.0)
        bbox_min_y = _as_float(bbox.get("min_y", 0.0), 0.0)
        nesting_shape = dict(geometry.get("nesting_shape", {}) or {})
        preview_shape = dict(geometry.get("preview_paths", {}) or {})
        raw_outer = list(nesting_shape.get("outer_polygons", []) or [])
        raw_holes = list(nesting_shape.get("hole_polygons", []) or [])
        preview_paths = _normalize_path_collection(
            preview_shape.get("cut_paths", ()),
            offset_x=bbox_min_x,
            offset_y=bbox_min_y,
        )
        outer_polygons = tuple(
            polygon
            for polygon in (
                _normalize_polygon(list(points or []), offset_x=bbox_min_x, offset_y=bbox_min_y)
                for points in raw_outer
            )
            if polygon
        )
        hole_polygons = tuple(
            polygon
            for polygon in (
                _normalize_polygon(list(points or []), offset_x=bbox_min_x, offset_y=bbox_min_y)
                for points in raw_holes
            )
            if polygon
        )
        if width <= 0.0 or height <= 0.0:
            warnings.append(f"{file_name}: caixa invalida para nesting.")
            continue
        if net_area_mm2 <= 0.0:
            warnings.append(f"{file_name}: area liquida indisponivel; confirma o DXF para um aproveitamento real fiavel.")
        if not outer_polygons:
            fallback_polygon = _rectangle_polygon(width, height)
            outer_polygons = (fallback_polygon,) if fallback_polygon else ()
        exploded_components = _explode_multi_part_shape(outer_polygons, hole_polygons)
        if exploded_components:
            warnings.append(
                f"{file_name}: DXF multi-peca detetado; {len(exploded_components)} subpecas independentes foram desdobradas automaticamente para nesting."
            )
            exploded_preview_paths = _explode_multi_part_preview_paths(preview_paths, exploded_components)
            base_ref = _row_ref(row)
            base_description = _row_description(row)
            total_components = len(exploded_components)
            for component in list(exploded_components or []):
                component_index = _as_int(component.get("component_index", 0), 0)
                suffix = f" [{component_index}/{total_components}]"
                items.append(
                    NestItem(
                        source_index=index,
                        path=path,
                        description=f"{base_description}{suffix}" if base_description else f"{file_name}{suffix}",
                        ref_externa=f"{base_ref}{suffix}" if base_ref else f"{Path(file_name).stem}{suffix}",
                        material=_row_material(row),
                        thickness_mm=_row_thickness_mm(row),
                        qty=_row_qty(row),
                        bbox_width_mm=max(0.0, _as_float(component.get("bbox_width_mm", 0.0), 0.0)),
                        bbox_height_mm=max(0.0, _as_float(component.get("bbox_height_mm", 0.0), 0.0)),
                        net_area_mm2=max(0.0, _as_float(component.get("net_area_mm2", 0.0), 0.0)),
                        file_name=f"{file_name}{suffix}",
                        geometry_warnings=tuple(geometry_warnings),
                        outer_polygons=tuple(component.get("outer_polygons", ()) or ()),
                        hole_polygons=tuple(component.get("hole_polygons", ()) or ()),
                        preview_paths=tuple(exploded_preview_paths.get(max(0, component_index - 1), ()) or ()),
                        shape_source="dxf",
                        rotation_policy=_row_rotation_policy(row),
                        priority=_row_priority(row),
                        shape_cache_key=f"{path}::{component_index}",
                    )
                )
            continue
        items.append(
            NestItem(
                source_index=index,
                path=path,
                description=_row_description(row),
                ref_externa=_row_ref(row),
                material=_row_material(row),
                thickness_mm=_row_thickness_mm(row),
                qty=_row_qty(row),
                bbox_width_mm=width,
                bbox_height_mm=height,
                net_area_mm2=net_area_mm2,
                file_name=file_name,
                geometry_warnings=tuple(geometry_warnings),
                outer_polygons=outer_polygons,
                hole_polygons=hole_polygons,
                preview_paths=preview_paths,
                shape_source="dxf" if list(nesting_shape.get("outer_polygons", []) or []) else "bbox",
                rotation_policy=_row_rotation_policy(row),
                priority=_row_priority(row),
                shape_cache_key=path,
            )
        )
    return items, _unique_texts(warnings)


def _expand_items(items: list[NestItem]) -> list[dict[str, Any]]:
    expanded: list[dict[str, Any]] = []
    for item in items:
        for copy_index in range(item.qty):
            expanded.append(
                {
                    "item": item,
                    "copy_index": copy_index + 1,
                    "bbox_area_mm2": max(0.0, item.bbox_width_mm * item.bbox_height_mm),
                }
            )
    return expanded


def _candidate_orientations(item: NestItem, allow_rotate: bool, allow_mirror: bool = False, free_angles: bool = False) -> list[dict[str, Any]]:
    policy = str(getattr(item, "rotation_policy", "auto") or "auto").strip().lower()
    has_rotation = bool(abs(item.bbox_width_mm - item.bbox_height_mm) > 1e-6)
    can_rotate = bool(allow_rotate and has_rotation)
    mirror_values = (False, True) if allow_mirror else (False,)

    def _variants(base: list[dict[str, Any]]) -> list[dict[str, Any]]:
        return [dict(orientation, mirrored=mirrored) for orientation in base for mirrored in mirror_values]

    if policy == "90":
        if has_rotation:
            return _variants([{"width": item.bbox_height_mm, "height": item.bbox_width_mm, "rotated": True, "angle_deg": 90.0}])
        return _variants([{"width": item.bbox_width_mm, "height": item.bbox_height_mm, "rotated": False, "angle_deg": 0.0}])
    if policy == "0" or not has_rotation:
        return _variants([{"width": item.bbox_width_mm, "height": item.bbox_height_mm, "rotated": False, "angle_deg": 0.0}])
    base = [
        {"width": item.bbox_width_mm, "height": item.bbox_height_mm, "rotated": False, "angle_deg": 0.0},
        {"width": item.bbox_height_mm, "height": item.bbox_width_mm, "rotated": True, "angle_deg": 90.0},
    ]
    if free_angles and policy == "auto":
        for angle in (15.0, 30.0, 45.0, 60.0, 75.0, 105.0, 120.0, 135.0, 150.0, 165.0):
            base.append({"width": item.bbox_width_mm, "height": item.bbox_height_mm, "rotated": abs(angle % 180.0 - 90.0) < 1e-6, "angle_deg": angle})
    return _variants(base)


def _transform_item_points(item: NestItem, points: list[tuple[float, float]], rotated: bool, mirrored: bool, angle_deg: float | None = None) -> list[tuple[float, float]]:
    transformed: list[tuple[float, float]] = []
    angle = 90.0 if angle_deg is None and rotated else float(angle_deg or 0.0)
    rad = math.radians(angle)
    cos_a = math.cos(rad)
    sin_a = math.sin(rad)
    for x, y in points:
        tx = max(0.0, float(item.bbox_width_mm) - float(x)) if mirrored else float(x)
        ty = float(y)
        if abs(angle) < 1e-9:
            transformed.append((round(tx, 3), round(ty, 3)))
        else:
            transformed.append((round((tx * cos_a) - (ty * sin_a), 3), round((tx * sin_a) + (ty * cos_a), 3)))
    return transformed


def _item_shape_polygons(item: NestItem, rotated: bool, mirrored: bool = False, angle_deg: float | None = None) -> tuple[tuple[tuple[tuple[float, float], ...], ...], tuple[tuple[tuple[float, float], ...], ...], float, float]:
    outer = tuple(tuple((float(x), float(y)) for x, y in list(polygon or [])) for polygon in list(item.outer_polygons or []))
    holes = tuple(tuple((float(x), float(y)) for x, y in list(polygon or [])) for polygon in list(item.hole_polygons or []))
    angle = 90.0 if angle_deg is None and rotated else float(angle_deg or 0.0)
    if abs(angle) < 1e-9 and not mirrored:
        return outer, holes, item.bbox_width_mm, item.bbox_height_mm

    transformed_outer: list[list[tuple[float, float]]] = []
    transformed_holes: list[list[tuple[float, float]]] = []
    all_points: list[tuple[float, float]] = []

    for polygon in outer:
        transformed_points = _transform_item_points(item, list(polygon), rotated, mirrored, angle)
        transformed_outer.append(transformed_points)
        all_points.extend(transformed_points)
    for polygon in holes:
        transformed_points = _transform_item_points(item, list(polygon), rotated, mirrored, angle)
        transformed_holes.append(transformed_points)
        all_points.extend(transformed_points)
    bbox = _points_bbox(all_points)
    offset_x = bbox["min_x"]
    offset_y = bbox["min_y"]
    normalized_outer = tuple(
        _normalize_polygon(points, offset_x=offset_x, offset_y=offset_y)
        for points in transformed_outer
    )
    normalized_holes = tuple(
        _normalize_polygon(points, offset_x=offset_x, offset_y=offset_y)
        for points in transformed_holes
    )
    return (
        tuple(polygon for polygon in normalized_outer if polygon),
        tuple(polygon for polygon in normalized_holes if polygon),
        bbox["width"],
        bbox["height"],
    )


def _item_preview_paths(item: NestItem, rotated: bool, mirrored: bool = False, angle_deg: float | None = None) -> tuple[tuple[tuple[float, float], ...], ...]:
    paths = tuple(tuple((float(x), float(y)) for x, y in list(path or [])) for path in list(item.preview_paths or []))
    if not paths:
        return ()
    angle = 90.0 if angle_deg is None and rotated else float(angle_deg or 0.0)
    if abs(angle) < 1e-9 and not mirrored:
        return tuple(path for path in paths if len(path) >= 2)

    transformed_paths: list[list[tuple[float, float]]] = []
    all_points: list[tuple[float, float]] = []
    for path in list(paths or []):
        transformed_path = _transform_item_points(item, list(path), rotated, mirrored, angle)
        if len(transformed_path) < 2:
            continue
        transformed_paths.append(transformed_path)
        all_points.extend(transformed_path)
    if not all_points:
        return ()
    bbox = _points_bbox(all_points)
    return tuple(
        path
        for path in (
            _normalize_path(points, offset_x=bbox["min_x"], offset_y=bbox["min_y"])
            for points in transformed_paths
        )
        if path
    )


def _cell_hits_shape(
    cell_x_mm: float,
    cell_y_mm: float,
    grid_mm: float,
    outer_polygons: tuple[tuple[tuple[float, float], ...], ...],
    hole_polygons: tuple[tuple[tuple[float, float], ...], ...],
    outer_bboxes: tuple[dict[str, float], ...] | None = None,
    hole_bboxes: tuple[dict[str, float], ...] | None = None,
) -> bool:
    outer_list = tuple(outer_polygons or ())
    hole_list = tuple(hole_polygons or ())
    outer_boxes = outer_bboxes or tuple(_points_bbox(list(polygon or ())) for polygon in outer_list)
    hole_boxes = hole_bboxes or tuple(_points_bbox(list(polygon or ())) for polygon in hole_list)

    def inside_any(
        point: tuple[float, float],
        polygons: tuple[tuple[tuple[float, float], ...], ...],
        bboxes: tuple[dict[str, float], ...],
    ) -> bool:
        px, py = point
        return any(
            bbox["min_x"] <= px <= bbox["max_x"]
            and bbox["min_y"] <= py <= bbox["max_y"]
            and _point_in_polygon(point, polygon)
            for polygon, bbox in zip(polygons, bboxes)
        )

    sample_points = [
        (cell_x_mm + (grid_mm * 0.50), cell_y_mm + (grid_mm * 0.50)),
        (cell_x_mm + (grid_mm * 0.20), cell_y_mm + (grid_mm * 0.20)),
        (cell_x_mm + (grid_mm * 0.80), cell_y_mm + (grid_mm * 0.20)),
        (cell_x_mm + (grid_mm * 0.80), cell_y_mm + (grid_mm * 0.80)),
        (cell_x_mm + (grid_mm * 0.20), cell_y_mm + (grid_mm * 0.80)),
    ]
    for point in sample_points:
        if inside_any(point, outer_list, outer_boxes) and not inside_any(point, hole_list, hole_boxes):
            return True
    for polygon in outer_list:
        for px, py in list(polygon or []):
            if cell_x_mm <= px <= (cell_x_mm + grid_mm) and cell_y_mm <= py <= (cell_y_mm + grid_mm):
                if not inside_any((px, py), hole_list, hole_boxes):
                    return True
    return False


def _expand_cells(cells: set[tuple[int, int]], radius_cells: int) -> set[tuple[int, int]]:
    if radius_cells <= 0:
        return set(cells)
    out: set[tuple[int, int]] = set()
    limit = max(0, int(radius_cells))
    for cell_x, cell_y in list(cells or []):
        for dy in range(-limit, limit + 1):
            for dx in range(-limit, limit + 1):
                if (dx * dx) + (dy * dy) > (limit * limit):
                    continue
                out.add((cell_x + dx, cell_y + dy))
    return out


def _shape_mask(
    item: NestItem,
    *,
    rotated: bool,
    mirrored: bool = False,
    angle_deg: float = 0.0,
    grid_mm: float,
    part_spacing_mm: float,
    cache: dict[tuple, dict[str, Any]],
) -> dict[str, Any]:
    cache_key = str(getattr(item, "shape_cache_key", "") or item.path)
    key = (cache_key, round(float(angle_deg or 0.0), 4), bool(mirrored), round(grid_mm, 4), round(part_spacing_mm, 4))
    cached = cache.get(key)
    if cached is not None:
        return cached

    outer_polygons, hole_polygons, width_mm, height_mm = _item_shape_polygons(item, rotated, mirrored, angle_deg)
    if not outer_polygons:
        fallback_polygon = _rectangle_polygon(width_mm, height_mm)
        outer_polygons = (fallback_polygon,) if fallback_polygon else ()
    cols = max(1, int(math.ceil(max(width_mm, 0.0) / max(grid_mm, 1.0))))
    rows = max(1, int(math.ceil(max(height_mm, 0.0) / max(grid_mm, 1.0))))
    outer_bboxes = tuple(_points_bbox(list(polygon or ())) for polygon in outer_polygons)
    hole_bboxes = tuple(_points_bbox(list(polygon or ())) for polygon in hole_polygons)
    base_cells: set[tuple[int, int]] = set()
    for row_index in range(rows):
        cell_y_mm = row_index * grid_mm
        for col_index in range(cols):
            cell_x_mm = col_index * grid_mm
            if _cell_hits_shape(
                cell_x_mm,
                cell_y_mm,
                grid_mm,
                outer_polygons,
                hole_polygons,
                outer_bboxes,
                hole_bboxes,
            ):
                base_cells.add((col_index, row_index))
    if not base_cells:
        base_cells = {(col_index, row_index) for row_index in range(rows) for col_index in range(cols)}

    spacing_radius_cells = max(0, int(math.ceil(max(0.0, part_spacing_mm) / max(grid_mm, 1.0))))
    min_cell_x = min((cell[0] for cell in base_cells), default=0)
    min_cell_y = min((cell[1] for cell in base_cells), default=0)
    max_cell_x = max((cell[0] for cell in base_cells), default=cols - 1)
    max_cell_y = max((cell[1] for cell in base_cells), default=rows - 1)
    normalized_cells = {
        (cell_x - min_cell_x, cell_y - min_cell_y)
        for cell_x, cell_y in list(base_cells or [])
    }
    payload = {
        "cells": tuple(sorted(normalized_cells, key=lambda cell: (cell[1], cell[0]))),
        "width_cells": max(1, (max_cell_x - min_cell_x) + 1),
        "height_cells": max(1, (max_cell_y - min_cell_y) + 1),
        "draw_offset_x_mm": round(max(0.0, -min_cell_x * grid_mm), 3),
        "draw_offset_y_mm": round(max(0.0, -min_cell_y * grid_mm), 3),
        "occupied_area_mm2": round(len(normalized_cells) * grid_mm * grid_mm, 2),
        "spacing_radius_cells": spacing_radius_cells,
        "shape_outer_polygons": outer_polygons,
        "shape_hole_polygons": hole_polygons,
        "bbox_width_mm": round(width_mm, 3),
        "bbox_height_mm": round(height_mm, 3),
    }
    cache[key] = payload
    return payload


def _sheet_layout_bounds(sheet: dict[str, Any], *, edge_margin_mm: float = 0.0) -> tuple[float, float, float, float]:
    placements = list(sheet.get("placements", []) or [])
    if not placements:
        return 0.0, 0.0, 0.0, 0.0
    min_x = min(max(0.0, _as_float(placement.get("x_mm", 0.0), 0.0) - max(0.0, edge_margin_mm)) for placement in placements)
    min_y = min(max(0.0, _as_float(placement.get("y_mm", 0.0), 0.0) - max(0.0, edge_margin_mm)) for placement in placements)
    max_x = max(max(0.0, _as_float(placement.get("x_mm", 0.0), 0.0) - max(0.0, edge_margin_mm)) + _as_float(placement.get("width_mm", 0.0), 0.0) for placement in placements)
    max_y = max(max(0.0, _as_float(placement.get("y_mm", 0.0), 0.0) - max(0.0, edge_margin_mm)) + _as_float(placement.get("height_mm", 0.0), 0.0) for placement in placements)
    return min_x, min_y, max_x, max_y


def _projected_layout_metrics(
    sheet: dict[str, Any],
    candidate: dict[str, Any],
    *,
    edge_margin_mm: float = 0.0,
) -> dict[str, float]:
    placements = list(sheet.get("placements", []) or [])
    cand_x = _as_float(candidate.get("x", 0.0), 0.0)
    cand_y = _as_float(candidate.get("y", 0.0), 0.0)
    cand_w = _as_float(candidate.get("place_w", candidate.get("width", 0.0)), 0.0)
    cand_h = _as_float(candidate.get("place_h", candidate.get("height", 0.0)), 0.0)
    if placements:
        min_x, min_y, max_x, max_y = _sheet_layout_bounds(sheet, edge_margin_mm=edge_margin_mm)
        proj_min_x = min(min_x, cand_x)
        proj_min_y = min(min_y, cand_y)
        proj_max_x = max(max_x, cand_x + cand_w)
        proj_max_y = max(max_y, cand_y + cand_h)
    else:
        proj_min_x = cand_x
        proj_min_y = cand_y
        proj_max_x = cand_x + cand_w
        proj_max_y = cand_y + cand_h
    span_width = max(0.0, proj_max_x - proj_min_x)
    span_height = max(0.0, proj_max_y - proj_min_y)
    profile = dict(sheet.get("profile", {}) or {})
    sheet_width = max(0.0, _as_float(profile.get("width_mm", 0.0), 0.0))
    sheet_height = max(0.0, _as_float(profile.get("height_mm", 0.0), 0.0))
    largest_edge_remnant = 0.0
    fragmented_remainder = 0.0
    if sheet_width > 0.0 and sheet_height > 0.0:
        largest_edge_remnant = _largest_edge_remnant_area(
            sheet_width,
            sheet_height,
            proj_min_x,
            proj_min_y,
            proj_max_x,
            proj_max_y,
        )
        projected_remaining_bbox = max(0.0, (sheet_width * sheet_height) - (span_width * span_height))
        fragmented_remainder = max(0.0, projected_remaining_bbox - largest_edge_remnant)
    return {
        "min_x": round(proj_min_x, 3),
        "min_y": round(proj_min_y, 3),
        "max_x": round(proj_max_x, 3),
        "max_y": round(proj_max_y, 3),
        "span_width": round(span_width, 3),
        "span_height": round(span_height, 3),
        "largest_edge_remnant": round(largest_edge_remnant, 3),
        "fragmented_remainder": round(fragmented_remainder, 3),
    }


def _shape_candidate_score(
    candidate: dict[str, Any],
    *,
    strategy_name: str = "",
    sheet: dict[str, Any] | None = None,
    edge_margin_mm: float = 0.0,
) -> tuple[float, ...]:
    normalized = str(strategy_name or "").strip().lower()
    projected = _projected_layout_metrics(sheet or {}, candidate, edge_margin_mm=edge_margin_mm)
    if "retalho" in normalized:
        return (
            -_as_float(projected.get("largest_edge_remnant", 0.0), 0.0),
            _as_float(projected.get("fragmented_remainder", 0.0), 0.0),
            _as_float(projected.get("span_width", 0.0), 0.0),
            _as_float(projected.get("span_height", 0.0), 0.0),
            _as_float(candidate.get("x", 0.0), 0.0) + _as_float(candidate.get("place_w", 0.0), 0.0),
            _as_float(candidate.get("y", 0.0), 0.0) + _as_float(candidate.get("place_h", 0.0), 0.0),
            _as_float(candidate.get("x", 0.0), 0.0),
            _as_float(candidate.get("y", 0.0), 0.0),
            _as_float(candidate.get("occupied_area_mm2", 0.0), 0.0),
        )
    if "width-first" in normalized:
        return (
            _as_float(projected.get("span_width", 0.0), 0.0),
            _as_float(projected.get("span_height", 0.0), 0.0),
            _as_float(candidate.get("x", 0.0), 0.0) + _as_float(candidate.get("place_w", 0.0), 0.0),
            _as_float(candidate.get("y", 0.0), 0.0) + _as_float(candidate.get("place_h", 0.0), 0.0),
            _as_float(candidate.get("x", 0.0), 0.0),
            _as_float(candidate.get("y", 0.0), 0.0),
            _as_float(candidate.get("occupied_area_mm2", 0.0), 0.0),
        )
    return (
        _as_float(candidate.get("y", 0.0), 0.0) + _as_float(candidate.get("place_h", 0.0), 0.0),
        _as_float(candidate.get("x", 0.0), 0.0) + _as_float(candidate.get("place_w", 0.0), 0.0),
        _as_float(candidate.get("y", 0.0), 0.0),
        _as_float(candidate.get("x", 0.0), 0.0),
        _as_float(candidate.get("occupied_area_mm2", 0.0), 0.0),
    )


def _sheet_allowed_cells(profile: dict[str, Any], *, edge_margin_mm: float, grid_mm: float, width_cells: int, height_cells: int) -> tuple[set[int] | None, list[list[tuple[float, float]]], list[list[tuple[float, float]]]]:
    outer_polygons = tuple(profile.get("outer_polygons", ()) or ())
    hole_polygons = tuple(profile.get("hole_polygons", ()) or ())
    if not outer_polygons:
        return None, [], []
    outer_bboxes = tuple(_points_bbox(list(polygon or ())) for polygon in outer_polygons)
    hole_bboxes = tuple(_points_bbox(list(polygon or ())) for polygon in hole_polygons)
    allowed_indices: set[int] = set()
    for row_index in range(height_cells):
        cell_y_mm = edge_margin_mm + (row_index * grid_mm)
        for col_index in range(width_cells):
            cell_x_mm = edge_margin_mm + (col_index * grid_mm)
            if _cell_hits_shape(
                cell_x_mm,
                cell_y_mm,
                grid_mm,
                outer_polygons,
                hole_polygons,
                outer_bboxes,
                hole_bboxes,
            ):
                allowed_indices.add((row_index * width_cells) + col_index)
    if not allowed_indices:
        return None, _translate_polygons(outer_polygons, 0.0, 0.0), _translate_polygons(hole_polygons, 0.0, 0.0)
    return allowed_indices, _translate_polygons(outer_polygons, 0.0, 0.0), _translate_polygons(hole_polygons, 0.0, 0.0)


def _new_shape_sheet(profile: dict[str, Any], *, edge_margin_mm: float, grid_mm: float) -> dict[str, Any]:
    usable_width, usable_height = _profile_usable_dimensions(profile, edge_margin_mm)
    width_cells = max(1, int(math.floor(usable_width / max(grid_mm, 1.0))))
    height_cells = max(1, int(math.floor(usable_height / max(grid_mm, 1.0))))
    allowed_cells, sheet_outer_polygons, sheet_hole_polygons = _sheet_allowed_cells(
        profile,
        edge_margin_mm=edge_margin_mm,
        grid_mm=grid_mm,
        width_cells=width_cells,
        height_cells=height_cells,
    )
    return {
        "profile": dict(profile or {}),
        "placements": [],
        "occupied_cells": set(),
        "allowed_cells": allowed_cells,
        "grid_mm": float(grid_mm),
        "grid_width_cells": width_cells,
        "grid_height_cells": height_cells,
        "sheet_outer_polygons": sheet_outer_polygons,
        "sheet_hole_polygons": sheet_hole_polygons,
        "used_net_area_mm2": 0.0,
        "used_bbox_area_mm2": 0.0,
    }


def _try_place_on_shape_sheet(
    sheet: dict[str, Any],
    item: NestItem,
    *,
    allow_rotate: bool,
    allow_mirror: bool = False,
    free_angles: bool = False,
    grid_mm: float,
    part_spacing_mm: float,
    edge_margin_mm: float,
    cache: dict[tuple, dict[str, Any]],
    strategy_name: str = "",
) -> dict[str, Any] | None:
    occupied = set(sheet.get("occupied_cells", set()) or set())
    allowed = sheet.get("allowed_cells", None)
    sheet_width_cells = max(1, int(sheet.get("grid_width_cells", 1) or 1))
    sheet_height_cells = max(1, int(sheet.get("grid_height_cells", 1) or 1))
    best: dict[str, Any] | None = None

    for orientation in _candidate_orientations(item, allow_rotate, allow_mirror=allow_mirror, free_angles=free_angles):
        angle_deg = float(orientation.get("angle_deg", 90.0 if bool(orientation.get("rotated")) else 0.0) or 0.0)
        mask = _shape_mask(
            item,
            rotated=bool(orientation.get("rotated")),
            mirrored=bool(orientation.get("mirrored")),
            angle_deg=angle_deg,
            grid_mm=grid_mm,
            part_spacing_mm=part_spacing_mm,
            cache=cache,
        )
        mask_width_cells = int(mask.get("width_cells", 0) or 0)
        mask_height_cells = int(mask.get("height_cells", 0) or 0)
        if mask_width_cells <= 0 or mask_height_cells <= 0:
            continue
        if mask_width_cells > sheet_width_cells or mask_height_cells > sheet_height_cells:
            continue
        mask_cells = list(mask.get("cells", []) or [])
        spacing_radius_cells = max(0, int(mask.get("spacing_radius_cells", 0) or 0))
        def try_position(grid_x: int, grid_y: int) -> dict[str, Any] | None:
            blocked = False
            for cell_x, cell_y in mask_cells:
                index = ((grid_y + cell_y) * sheet_width_cells) + grid_x + cell_x
                if (allowed is not None and index not in allowed) or index in occupied:
                    blocked = True
                    break
                if spacing_radius_cells > 0:
                    base_x = grid_x + cell_x
                    base_y = grid_y + cell_y
                    for delta_y in range(-spacing_radius_cells, spacing_radius_cells + 1):
                        for delta_x in range(-spacing_radius_cells, spacing_radius_cells + 1):
                            if delta_x == 0 and delta_y == 0:
                                continue
                            if (delta_x * delta_x) + (delta_y * delta_y) > (spacing_radius_cells * spacing_radius_cells):
                                continue
                            probe_x = base_x + delta_x
                            probe_y = base_y + delta_y
                            if probe_x < 0 or probe_y < 0 or probe_x >= sheet_width_cells or probe_y >= sheet_height_cells:
                                continue
                            probe_index = (probe_y * sheet_width_cells) + probe_x
                            if probe_index in occupied:
                                blocked = True
                                break
                        if blocked:
                            break
                if blocked:
                    break
            if blocked:
                return None
            draw_x = (grid_x * grid_mm) + _as_float(mask.get("draw_offset_x_mm", 0.0), 0.0) + max(0.0, edge_margin_mm)
            draw_y = (grid_y * grid_mm) + _as_float(mask.get("draw_offset_y_mm", 0.0), 0.0) + max(0.0, edge_margin_mm)
            candidate_outer = tuple(
                tuple((round(draw_x + x, 3), round(draw_y + y, 3)) for x, y in list(points or []))
                for points in list(mask.get("shape_outer_polygons", ()) or [])
                if points
            )
            candidate_holes = tuple(
                tuple((round(draw_x + x, 3), round(draw_y + y, 3)) for x, y in list(points or []))
                for points in list(mask.get("shape_hole_polygons", ()) or [])
                if points
            )
            if _candidate_shape_conflicts(sheet, candidate_outer, candidate_holes):
                return None
            return {
                "x": round(grid_x * grid_mm, 3),
                "y": round(grid_y * grid_mm, 3),
                "grid_x": grid_x,
                "grid_y": grid_y,
                "width": round(_as_float(mask.get("bbox_width_mm", orientation.get("width", 0.0)), 0.0), 3),
                "height": round(_as_float(mask.get("bbox_height_mm", orientation.get("height", 0.0)), 0.0), 3),
                "rotated": bool(orientation.get("rotated")),
                "mirrored": bool(orientation.get("mirrored")),
                "rotation_deg": round(angle_deg, 3),
                "place_w": round(mask_width_cells * grid_mm, 3),
                "place_h": round(mask_height_cells * grid_mm, 3),
                "mask_cells": mask_cells,
                "draw_offset_x_mm": round(_as_float(mask.get("draw_offset_x_mm", 0.0), 0.0), 3),
                "draw_offset_y_mm": round(_as_float(mask.get("draw_offset_y_mm", 0.0), 0.0), 3),
                "occupied_area_mm2": round(_as_float(mask.get("occupied_area_mm2", 0.0), 0.0), 2),
                "shape_outer_polygons": tuple(mask.get("shape_outer_polygons", ()) or ()),
                "shape_hole_polygons": tuple(mask.get("shape_hole_polygons", ()) or ()),
            }

        candidate_positions = _shape_anchor_positions(
            sheet,
            sheet_width_cells=sheet_width_cells,
            sheet_height_cells=sheet_height_cells,
            mask_width_cells=mask_width_cells,
            mask_height_cells=mask_height_cells,
            grid_mm=grid_mm,
            edge_margin_mm=edge_margin_mm,
        )
        for grid_x, grid_y in candidate_positions:
            candidate = try_position(grid_x, grid_y)
            if candidate is None:
                continue
            if best is None or _shape_candidate_score(candidate, strategy_name=strategy_name, sheet=sheet, edge_margin_mm=edge_margin_mm) < _shape_candidate_score(best, strategy_name=strategy_name, sheet=sheet, edge_margin_mm=edge_margin_mm):
                best = candidate
        if best is not None:
            continue
        for grid_y in range(sheet_height_cells - mask_height_cells + 1):
            for grid_x in range(sheet_width_cells - mask_width_cells + 1):
                candidate = try_position(grid_x, grid_y)
                if candidate is None:
                    continue
                if best is None or _shape_candidate_score(candidate, strategy_name=strategy_name, sheet=sheet, edge_margin_mm=edge_margin_mm) < _shape_candidate_score(best, strategy_name=strategy_name, sheet=sheet, edge_margin_mm=edge_margin_mm):
                    best = candidate
    return best


def _placement_score(
    candidate: dict[str, Any],
    *,
    strategy_name: str = "",
    sheet: dict[str, Any] | None = None,
    edge_margin_mm: float = 0.0,
) -> tuple[float, ...]:
    normalized = str(strategy_name or "").strip().lower()
    projected = _projected_layout_metrics(sheet or {}, candidate, edge_margin_mm=edge_margin_mm)
    if "retalho" in normalized:
        return (
            -_as_float(projected.get("largest_edge_remnant", 0.0), 0.0),
            _as_float(projected.get("fragmented_remainder", 0.0), 0.0),
            _as_float(projected.get("span_width", 0.0), 0.0),
            _as_float(projected.get("span_height", 0.0), 0.0),
            0.0 if not bool(candidate.get("new_shelf")) else 1.0,
            _as_float(candidate.get("x", 0.0), 0.0) + _as_float(candidate.get("place_w", 0.0), 0.0),
            _as_float(candidate.get("y", 0.0), 0.0) + _as_float(candidate.get("place_h", 0.0), 0.0),
            _as_float(candidate.get("waste", 0.0), 0.0),
            _as_float(candidate.get("height_gap", 0.0), 0.0),
        )
    if "width-first" in normalized:
        return (
            _as_float(projected.get("span_width", 0.0), 0.0),
            _as_float(projected.get("span_height", 0.0), 0.0),
            0.0 if not bool(candidate.get("new_shelf")) else 1.0,
            _as_float(candidate.get("x", 0.0), 0.0) + _as_float(candidate.get("place_w", 0.0), 0.0),
            _as_float(candidate.get("y", 0.0), 0.0) + _as_float(candidate.get("place_h", 0.0), 0.0),
            _as_float(candidate.get("waste", 0.0), 0.0),
            _as_float(candidate.get("height_gap", 0.0), 0.0),
        )
    return (
        0.0 if not bool(candidate.get("new_shelf")) else 1.0,
        _as_float(candidate.get("y", 0.0), 0.0) + _as_float(candidate.get("place_h", 0.0), 0.0),
        _as_float(candidate.get("waste", 0.0), 0.0),
        _as_float(candidate.get("height_gap", 0.0), 0.0),
        _as_float(candidate.get("place_h", 0.0), 0.0),
    )


def _largest_edge_remnant_area(width_mm: float, height_mm: float, min_x: float, min_y: float, max_x: float, max_y: float) -> float:
    left_area = max(0.0, min_x) * max(0.0, height_mm)
    right_area = max(0.0, width_mm - max_x) * max(0.0, height_mm)
    top_area = max(0.0, min_y) * max(0.0, width_mm)
    bottom_area = max(0.0, height_mm - max_y) * max(0.0, width_mm)
    return max(left_area, right_area, top_area, bottom_area)


def _profile_usable_dimensions(profile: dict[str, Any], edge_margin_mm: float) -> tuple[float, float]:
    width_mm = max(0.0, _as_float(profile.get("width_mm", 0.0), 0.0))
    height_mm = max(0.0, _as_float(profile.get("height_mm", 0.0), 0.0))
    usable_width = width_mm - (2.0 * max(0.0, edge_margin_mm))
    usable_height = height_mm - (2.0 * max(0.0, edge_margin_mm))
    if usable_width <= 0.0 or usable_height <= 0.0:
        raise ValueError("A margem a borda e maior do que a chapa util disponivel.")
    return usable_width, usable_height


def _profile_can_fit_all_items(
    profile: dict[str, Any],
    items: list[NestItem],
    *,
    edge_margin_mm: float,
    allow_rotate: bool,
) -> bool:
    try:
        usable_width, usable_height = _profile_usable_dimensions(profile, edge_margin_mm)
    except Exception:
        return False
    for item in list(items or []):
        if not any(
            float(orientation.get("width", 0.0) or 0.0) <= usable_width + 1e-6
            and float(orientation.get("height", 0.0) or 0.0) <= usable_height + 1e-6
            for orientation in _candidate_orientations(item, allow_rotate, allow_mirror=False, free_angles=False)
        ):
            return False
    return True


def _try_place_on_sheet(
    sheet: dict[str, Any],
    item: NestItem,
    *,
    usable_width: float,
    usable_height: float,
    part_spacing_mm: float,
    allow_rotate: bool,
    strategy_name: str = "",
    edge_margin_mm: float = 0.0,
) -> dict[str, Any] | None:
    best: dict[str, Any] | None = None
    for orientation in _candidate_orientations(item, allow_rotate):
        place_w = orientation["width"]
        place_h = orientation["height"]
        for shelf_index, shelf in enumerate(list(sheet.get("shelves", []) or [])):
            if place_h > _as_float(shelf.get("height", 0.0), 0.0) + 1e-6:
                continue
            shelf_x = _as_float(shelf.get("x", 0.0), 0.0)
            if shelf_x + place_w > usable_width + 1e-6:
                continue
            candidate = {
                "shelf_index": shelf_index,
                "new_shelf": False,
                "x": shelf_x,
                "y": _as_float(shelf.get("y", 0.0), 0.0),
                "width": orientation["width"],
                "height": orientation["height"],
                "rotated": orientation["rotated"],
                "place_w": place_w,
                "place_h": place_h,
                "waste": usable_width - (shelf_x + place_w),
                "height_gap": max(0.0, _as_float(shelf.get("height", 0.0), 0.0) - place_h),
            }
            if best is None or _placement_score(candidate, strategy_name=strategy_name, sheet=sheet, edge_margin_mm=edge_margin_mm) < _placement_score(best, strategy_name=strategy_name, sheet=sheet, edge_margin_mm=edge_margin_mm):
                best = candidate
        cursor_y = _as_float(sheet.get("cursor_y", 0.0), 0.0)
        if place_w <= usable_width + 1e-6 and cursor_y + place_h <= usable_height + 1e-6:
            candidate = {
                "shelf_index": len(list(sheet.get("shelves", []) or [])),
                "new_shelf": True,
                "x": 0.0,
                "y": cursor_y,
                "width": orientation["width"],
                "height": orientation["height"],
                "rotated": orientation["rotated"],
                "place_w": place_w,
                "place_h": place_h,
                "waste": usable_width - place_w,
                "height_gap": 0.0,
            }
            if best is None or _placement_score(candidate, strategy_name=strategy_name, sheet=sheet, edge_margin_mm=edge_margin_mm) < _placement_score(best, strategy_name=strategy_name, sheet=sheet, edge_margin_mm=edge_margin_mm):
                best = candidate
    return best


def _new_sheet(profile: dict[str, Any]) -> dict[str, Any]:
    return {
        "profile": dict(profile or {}),
        "placements": [],
        "shelves": [],
        "cursor_y": 0.0,
        "used_net_area_mm2": 0.0,
        "used_bbox_area_mm2": 0.0,
    }


def _apply_placement(
    sheet: dict[str, Any],
    placement: dict[str, Any],
    row: dict[str, Any],
    *,
    edge_margin_mm: float,
    part_spacing_mm: float,
) -> None:
    item: NestItem = row["item"]
    if "mask_cells" in placement:
        occupied = set(sheet.get("occupied_cells", set()) or set())
        sheet_width_cells = max(1, int(sheet.get("grid_width_cells", 1) or 1))
        grid_x = int(placement.get("grid_x", 0) or 0)
        grid_y = int(placement.get("grid_y", 0) or 0)
        for cell_x, cell_y in list(placement.get("mask_cells", []) or []):
            occupied.add(((grid_y + int(cell_y)) * sheet_width_cells) + grid_x + int(cell_x))
        sheet["occupied_cells"] = occupied
    elif bool(placement.get("new_shelf")):
        shelves = list(sheet.get("shelves", []) or [])
        shelves.append(
            {
                "y": _as_float(placement.get("y", 0.0), 0.0),
                "height": _as_float(placement.get("height", 0.0), 0.0),
                "x": _as_float(placement.get("x", 0.0), 0.0) + _as_float(placement.get("width", 0.0), 0.0) + max(0.0, part_spacing_mm),
            }
        )
        sheet["shelves"] = shelves
        sheet["cursor_y"] = _as_float(placement.get("y", 0.0), 0.0) + _as_float(placement.get("height", 0.0), 0.0) + max(0.0, part_spacing_mm)
    else:
        shelf_index = int(placement.get("shelf_index", 0) or 0)
        shelves = list(sheet.get("shelves", []) or [])
        if 0 <= shelf_index < len(shelves):
            shelves[shelf_index]["x"] = _as_float(placement.get("x", 0.0), 0.0) + _as_float(placement.get("width", 0.0), 0.0) + max(0.0, part_spacing_mm)
            sheet["shelves"] = shelves
    draw_x = edge_margin_mm + _as_float(placement.get("x", 0.0), 0.0)
    draw_y = edge_margin_mm + _as_float(placement.get("y", 0.0), 0.0)
    draw_x += _as_float(placement.get("draw_offset_x_mm", 0.0), 0.0)
    draw_y += _as_float(placement.get("draw_offset_y_mm", 0.0), 0.0)
    layout_area_mm2 = round(
        _as_float(
            placement.get("occupied_area_mm2", item.bbox_width_mm * item.bbox_height_mm),
            item.bbox_width_mm * item.bbox_height_mm,
        ),
        2,
    )
    placements = list(sheet.get("placements", []) or [])
    placements.append(
        {
            "path": item.path,
            "file_name": item.file_name,
            "description": item.description,
            "ref_externa": item.ref_externa,
            "material": item.material,
            "thickness_mm": item.thickness_mm,
            "rotated": bool(placement.get("rotated")),
            "mirrored": bool(placement.get("mirrored")),
            "rotation_deg": round(_as_float(placement.get("rotation_deg", 90.0 if bool(placement.get("rotated")) else 0.0), 0.0), 3),
            "x_mm": round(draw_x, 3),
            "y_mm": round(draw_y, 3),
            "width_mm": round(_as_float(placement.get("width", 0.0), 0.0), 3),
            "height_mm": round(_as_float(placement.get("height", 0.0), 0.0), 3),
            "net_area_mm2": round(item.net_area_mm2, 2),
            "bbox_area_mm2": round(item.bbox_width_mm * item.bbox_height_mm, 2),
            "layout_area_mm2": layout_area_mm2,
            "copy_index": int(row.get("copy_index", 0) or 0),
            "shape_mode": str(placement.get("shape_mode", "grid" if "mask_cells" in placement else "bbox") or "bbox"),
            "required_spacing_mm": round(max(0.0, float(part_spacing_mm or 0.0)), 3),
            "required_edge_margin_mm": round(max(0.0, float(edge_margin_mm or 0.0)), 3),
            "shape_outer_polygons": _translate_polygons(tuple(placement.get("shape_outer_polygons", ()) or ()), draw_x, draw_y),
            "shape_hole_polygons": _translate_polygons(tuple(placement.get("shape_hole_polygons", ()) or ()), draw_x, draw_y),
            "preview_paths": _translate_paths(
                _item_preview_paths(
                    item,
                    bool(placement.get("rotated")),
                    bool(placement.get("mirrored")),
                    _as_float(placement.get("rotation_deg", 90.0 if bool(placement.get("rotated")) else 0.0), 0.0),
                ),
                draw_x,
                draw_y,
            ),
        }
    )
    sheet["placements"] = placements
    sheet["used_net_area_mm2"] = _as_float(sheet.get("used_net_area_mm2", 0.0), 0.0) + item.net_area_mm2
    sheet["used_bbox_area_mm2"] = _as_float(sheet.get("used_bbox_area_mm2", 0.0), 0.0) + layout_area_mm2


def _strategy_sort_key(name: str):
    normalized = str(name or "").strip().lower()
    if normalized.startswith("shape-"):
        normalized = normalized[6:]
    if normalized == "compact":
        return lambda row: (
            max(-1, min(2, _as_int(getattr(row.get("item"), "priority", row.get("priority", 0)), 0))),
            _as_float(row.get("bbox_area_mm2", 0.0), 0.0),
            min(row["item"].bbox_width_mm, row["item"].bbox_height_mm),
            _as_float(row["item"].net_area_mm2, 0.0) / max(1.0, _as_float(row.get("bbox_area_mm2", 0.0), 0.0)),
            max(row["item"].bbox_width_mm, row["item"].bbox_height_mm),
        )
    if normalized in {"retalho", "retalho-side", "retalho-useful"}:
        return lambda row: (
            max(-1, min(2, _as_int(getattr(row.get("item"), "priority", row.get("priority", 0)), 0))),
            max(row["item"].bbox_width_mm, row["item"].bbox_height_mm),
            min(row["item"].bbox_width_mm, row["item"].bbox_height_mm),
            _as_float(row.get("bbox_area_mm2", 0.0), 0.0),
            _as_float(row["item"].net_area_mm2, 0.0) / max(1.0, _as_float(row.get("bbox_area_mm2", 0.0), 0.0)),
        )
    if normalized == "area":
        return lambda row: (
            max(-1, min(2, _as_int(getattr(row.get("item"), "priority", row.get("priority", 0)), 0))),
            _as_float(row.get("bbox_area_mm2", 0.0), 0.0),
            max(row["item"].bbox_width_mm, row["item"].bbox_height_mm),
            min(row["item"].bbox_width_mm, row["item"].bbox_height_mm),
        )
    if normalized == "height-first":
        return lambda row: (
            max(-1, min(2, _as_int(getattr(row.get("item"), "priority", row.get("priority", 0)), 0))),
            max(row["item"].bbox_height_mm, row["item"].bbox_width_mm),
            row["item"].bbox_height_mm,
            row["item"].bbox_width_mm,
            _as_float(row.get("bbox_area_mm2", 0.0), 0.0),
        )
    if normalized == "width-first":
        return lambda row: (
            max(-1, min(2, _as_int(getattr(row.get("item"), "priority", row.get("priority", 0)), 0))),
            max(row["item"].bbox_width_mm, row["item"].bbox_height_mm),
            row["item"].bbox_width_mm,
            row["item"].bbox_height_mm,
            _as_float(row.get("bbox_area_mm2", 0.0), 0.0),
        )
    return lambda row: (
        max(-1, min(2, _as_int(getattr(row.get("item"), "priority", row.get("priority", 0)), 0))),
        max(row["item"].bbox_width_mm, row["item"].bbox_height_mm),
        _as_float(row.get("bbox_area_mm2", 0.0), 0.0),
        min(row["item"].bbox_width_mm, row["item"].bbox_height_mm),
    )


def _material_lookup(settings: dict[str, Any], material_name: str) -> tuple[dict[str, Any], dict[str, Any]]:
    machine_profiles = dict(settings.get("machine_profiles", {}) or {})
    machine_name = str(settings.get("active_machine", "") or "").strip()
    machine_profile = dict(machine_profiles.get(machine_name, {}) or {})
    if not machine_profile and machine_profiles:
        machine_profile = dict(next(iter(machine_profiles.values())) or {})
    machine_materials = dict(machine_profile.get("materials", {}) or {})
    machine_material = dict(machine_materials.get(material_name, {}) or {})

    commercial_profiles = dict(settings.get("commercial_profiles", {}) or {})
    commercial_name = str(settings.get("active_commercial", "") or "").strip()
    commercial_profile = dict(commercial_profiles.get(commercial_name, {}) or {})
    if not commercial_profile and commercial_profiles:
        commercial_profile = dict(next(iter(commercial_profiles.values())) or {})
    commercial_materials = dict(commercial_profile.get("materials", {}) or {})
    commercial_material = dict(commercial_materials.get(material_name, {}) or {})
    return machine_material, {**commercial_profile, "_material": commercial_material}


def _material_estimate(items: list[NestItem], summary: dict[str, Any], settings: dict[str, Any]) -> dict[str, Any]:
    if not items:
        return {
            "gross_sheet_mass_kg": 0.0,
            "purchase_sheet_mass_kg": 0.0,
            "stock_sheet_mass_kg": 0.0,
            "net_part_mass_kg": 0.0,
            "scrap_mass_kg": 0.0,
            "material_purchase_cost_eur": 0.0,
            "material_purchase_requirement_eur": 0.0,
            "material_scrap_credit_eur": 0.0,
            "material_net_cost_eur": 0.0,
        }
    primary = items[0]
    machine_material, commercial_profile = _material_lookup(settings, primary.material)
    commercial_material = dict(commercial_profile.get("_material", {}) or {})
    density_kg_m3 = max(
        1.0,
        _as_float(
            commercial_material.get("density_kg_m3", machine_material.get("density_kg_m3", 7800.0)),
            7800.0,
        ),
    )
    price_per_kg = max(0.0, _as_float(commercial_material.get("price_per_kg", 0.0), 0.0))
    scrap_credit_per_kg = max(0.0, _as_float(commercial_material.get("scrap_credit_per_kg", 0.0), 0.0))
    use_scrap_credit = bool(commercial_profile.get("use_scrap_credit", True))
    thickness_m = max(0.0, primary.thickness_mm / 1000.0)
    gross_area_m2 = max(0.0, _as_float(summary.get("total_sheet_area_mm2", 0.0), 0.0) / 1_000_000.0)
    purchase_area_m2 = max(0.0, _as_float(summary.get("purchase_sheet_area_mm2", 0.0), 0.0) / 1_000_000.0)
    stock_area_m2 = max(0.0, _as_float(summary.get("stock_sheet_area_mm2", 0.0), 0.0) / 1_000_000.0)
    net_area_m2 = max(0.0, _as_float(summary.get("used_net_area_mm2", 0.0), 0.0) / 1_000_000.0)
    gross_mass_kg = gross_area_m2 * thickness_m * density_kg_m3
    purchase_mass_kg = purchase_area_m2 * thickness_m * density_kg_m3
    stock_mass_kg = stock_area_m2 * thickness_m * density_kg_m3
    net_mass_kg = net_area_m2 * thickness_m * density_kg_m3
    scrap_mass_kg = max(0.0, gross_mass_kg - net_mass_kg)
    total_material_cost = gross_mass_kg * price_per_kg
    purchase_requirement_cost = purchase_mass_kg * price_per_kg
    scrap_credit = scrap_mass_kg * scrap_credit_per_kg if use_scrap_credit else 0.0
    return {
        "gross_sheet_mass_kg": round(gross_mass_kg, 4),
        "purchase_sheet_mass_kg": round(purchase_mass_kg, 4),
        "stock_sheet_mass_kg": round(stock_mass_kg, 4),
        "net_part_mass_kg": round(net_mass_kg, 4),
        "scrap_mass_kg": round(scrap_mass_kg, 4),
        "material_purchase_cost_eur": round(total_material_cost, 2),
        "material_purchase_requirement_eur": round(purchase_requirement_cost, 2),
        "material_scrap_credit_eur": round(scrap_credit, 2),
        "material_net_cost_eur": round(max(0.0, total_material_cost - scrap_credit), 2),
    }


def _result_score(result: dict[str, Any]) -> tuple[Any, ...]:
    summary = dict(result.get("summary", {}) or {})
    sheet_rows = [dict(row or {}) for row in list(result.get("sheets", []) or [])]
    purchase_area = _as_float(summary.get("purchase_sheet_area_mm2", 0.0), 0.0)
    total_area = _as_float(summary.get("total_sheet_area_mm2", 0.0), 0.0)
    span_area = _as_float(summary.get("layout_span_area_mm2", 0.0), 0.0)
    utilization_net = _as_float(summary.get("utilization_net_pct", 0.0), 0.0)
    compactness = _as_float(summary.get("layout_compactness_pct", 0.0), 0.0)
    wasted_area = max(0.0, total_area - _as_float(summary.get("used_net_area_mm2", 0.0), 0.0))
    largest_edge_remnant = _as_float(summary.get("largest_edge_remnant_area_mm2", 0.0), 0.0)
    fragmented_remainder = _as_float(summary.get("fragmented_remainder_area_mm2", 0.0), 0.0)
    sequential_fill_reward = sum(
        (
            _as_float(row.get("used_net_area_mm2", 0.0), 0.0)
            / max(1.0, _as_float(row.get("sheet_area_mm2", 0.0), 0.0))
        )
        ** 2
        for row in sheet_rows
    )
    sequential_fill_profile = tuple(
        -value
        for value in sorted(
            (
                _as_float(row.get("used_net_area_mm2", 0.0), 0.0)
                / max(1.0, _as_float(row.get("sheet_area_mm2", 0.0), 0.0))
                for row in sheet_rows
            ),
            reverse=True,
        )
    )
    return (
        _as_int(summary.get("part_count_unplaced", 0), 0),
        purchase_area,
        _as_int(summary.get("sheet_count", 0), 0),
        sequential_fill_profile,
        -sequential_fill_reward,
        fragmented_remainder,
        -largest_edge_remnant,
        wasted_area,
        total_area,
        span_area,
        -utilization_net,
        -compactness,
        -_as_float(summary.get("utilization_bbox_pct", 0.0), 0.0),
    )


def _engine_mode_from_summary(summary: dict[str, Any]) -> str:
    raw = str(summary.get("engine_used", "") or "").strip().lower()
    if raw in {"shape", "bbox"}:
        return raw
    return "shape" if bool(summary.get("shape_aware")) else "bbox"


def _engine_method_label(summary: dict[str, Any]) -> str:
    requested = str(summary.get("engine_requested", "") or "").strip().lower()
    used = _engine_mode_from_summary(summary)
    label = "Contorno DXF" if used == "shape" else "Caixa DXF"
    if requested == "shape" and used == "bbox":
        label += " (fallback)"
    return label


def _engine_comparison_note(
    *,
    chosen_mode: str,
    chosen_result: dict[str, Any],
    other_mode: str,
    other_result: dict[str, Any],
    grid_mm: float,
) -> str:
    chosen_summary = dict(chosen_result.get("summary", {}) or {})
    other_summary = dict(other_result.get("summary", {}) or {})
    chosen_req = max(
        _as_int(chosen_summary.get("part_count_requested", 0), 0),
        _as_int(other_summary.get("part_count_requested", 0), 0),
    )
    chosen_placed = _as_int(chosen_summary.get("part_count_placed", 0), 0)
    other_placed = _as_int(other_summary.get("part_count_placed", 0), 0)
    chosen_purchase = _as_float(chosen_summary.get("purchase_sheet_area_mm2", 0.0), 0.0)
    other_purchase = _as_float(other_summary.get("purchase_sheet_area_mm2", 0.0), 0.0)
    chosen_total = _as_float(chosen_summary.get("total_sheet_area_mm2", 0.0), 0.0)
    other_total = _as_float(other_summary.get("total_sheet_area_mm2", 0.0), 0.0)
    chosen_sheets = _as_int(chosen_summary.get("sheet_count", 0), 0)
    other_sheets = _as_int(other_summary.get("sheet_count", 0), 0)
    chosen_label = "contorno DXF" if chosen_mode == "shape" else "caixa real do DXF"
    other_label = "contorno DXF" if other_mode == "shape" else "caixa real do DXF"
    if chosen_placed != other_placed:
        reason = f"coloca {chosen_placed}/{chosen_req} peça(s) contra {other_placed}/{chosen_req}"
    elif abs(chosen_purchase - other_purchase) > 0.5:
        reason = (
            "consome menos chapa de compra "
            f"({_as_float(chosen_purchase / 1_000_000.0, 0.0):.4f} m2 vs {_as_float(other_purchase / 1_000_000.0, 0.0):.4f} m2)"
        )
    elif abs(chosen_total - other_total) > 0.5:
        reason = (
            "consome menos chapa total "
            f"({_as_float(chosen_total / 1_000_000.0, 0.0):.4f} m2 vs {_as_float(other_total / 1_000_000.0, 0.0):.4f} m2)"
        )
    elif chosen_sheets != other_sheets:
        reason = f"usa {chosen_sheets} chapa(s) contra {other_sheets}"
    else:
        reason = "tem melhor pontuação global para este cenário"
    if other_mode == "shape":
        return (
            f"Motor DXF: foi escolhida a {chosen_label} em vez do {other_label} "
            f"(grelha {_as_float(grid_mm, 0.0):g} mm) porque {reason}."
        )
    return f"Motor DXF: foi escolhido o {chosen_label} em vez da {other_label} porque {reason}."


def _choose_best_engine_result(
    *,
    bbox_result: dict[str, Any] | None,
    shape_result: dict[str, Any] | None,
    requested_mode: str,
    grid_mm: float,
) -> dict[str, Any]:
    candidates: list[tuple[str, dict[str, Any]]] = []
    if bbox_result is not None:
        candidates.append(("bbox", dict(bbox_result or {})))
    if shape_result is not None:
        candidates.append(("shape", dict(shape_result or {})))
    if not candidates:
        raise ValueError("Sem resultados validos para comparar no motor de nesting.")
    chosen_mode, chosen_result = min(candidates, key=lambda item: _result_score(item[1]))
    summary = dict(chosen_result.get("summary", {}) or {})
    summary["engine_requested"] = str(requested_mode or "bbox").strip().lower() or "bbox"
    summary["engine_used"] = chosen_mode
    summary["engine_modes_tested"] = [mode for mode, _ in candidates]
    chosen_result["summary"] = summary
    warnings = list(chosen_result.get("warnings", []) or [])
    other_candidates = [(mode, result) for mode, result in candidates if mode != chosen_mode]
    if other_candidates:
        other_mode, other_result = min(other_candidates, key=lambda item: _result_score(item[1]))
        warnings.insert(
            0,
            _engine_comparison_note(
                chosen_mode=chosen_mode,
                chosen_result=chosen_result,
                other_mode=other_mode,
                other_result=other_result,
                grid_mm=grid_mm,
            ),
        )
    chosen_result["warnings"] = _unique_texts(warnings)
    return chosen_result


def _build_sheet_row(sheet: dict[str, Any], index: int) -> dict[str, Any]:
    profile = dict(sheet.get("profile", {}) or {})
    placements = list(sheet.get("placements", []) or [])
    width_mm = _as_float(profile.get("width_mm", 0.0), 0.0)
    height_mm = _as_float(profile.get("height_mm", 0.0), 0.0)
    area_mm2 = width_mm * height_mm
    used_net = _as_float(sheet.get("used_net_area_mm2", 0.0), 0.0)
    used_bbox = _as_float(sheet.get("used_bbox_area_mm2", 0.0), 0.0)
    if placements:
        min_x = min(_as_float(placement.get("x_mm", 0.0), 0.0) for placement in placements)
        min_y = min(_as_float(placement.get("y_mm", 0.0), 0.0) for placement in placements)
        max_x = max(
            _as_float(placement.get("x_mm", 0.0), 0.0) + _as_float(placement.get("width_mm", 0.0), 0.0)
            for placement in placements
        )
        max_y = max(
            _as_float(placement.get("y_mm", 0.0), 0.0) + _as_float(placement.get("height_mm", 0.0), 0.0)
            for placement in placements
        )
        layout_span_width = max(0.0, max_x - min_x)
        layout_span_height = max(0.0, max_y - min_y)
    else:
        layout_span_width = 0.0
        layout_span_height = 0.0
        min_x = 0.0
        min_y = 0.0
        max_x = 0.0
        max_y = 0.0
    layout_span_area = layout_span_width * layout_span_height
    largest_edge_remnant_area = _largest_edge_remnant_area(width_mm, height_mm, min_x, min_y, max_x, max_y) if placements else area_mm2
    remaining_bbox_area = max(0.0, area_mm2 - used_bbox)
    return {
        "index": index + 1,
        "profile_name": str(profile.get("name", "") or "").strip(),
        "sheet_width_mm": round(width_mm, 3),
        "sheet_height_mm": round(height_mm, 3),
        "sheet_area_mm2": round(area_mm2, 2),
        "source_kind": str(profile.get("source_kind", "purchase") or "purchase").strip().lower(),
        "source_label": str(profile.get("source_label", profile.get("name", "")) or "").strip(),
        "source_material_id": str(profile.get("material_id", "") or "").strip(),
        "source_lote": str(profile.get("lote", "") or "").strip(),
        "source_local": str(profile.get("local", "") or "").strip(),
        "opening_reason": str(sheet.get("opening_reason", "") or "").strip(),
        "sheet_outer_polygons": [list(polygon) for polygon in list(sheet.get("sheet_outer_polygons", []) or [])],
        "sheet_hole_polygons": [list(polygon) for polygon in list(sheet.get("sheet_hole_polygons", []) or [])],
        "placements": placements,
        "part_count": len(placements),
        "used_net_area_mm2": round(used_net, 2),
        "used_bbox_area_mm2": round(used_bbox, 2),
        "utilization_net_pct": round((used_net / area_mm2 * 100.0) if area_mm2 > 0.0 else 0.0, 2),
        "utilization_bbox_pct": round((used_bbox / area_mm2 * 100.0) if area_mm2 > 0.0 else 0.0, 2),
        "layout_span_width_mm": round(layout_span_width, 2),
        "layout_span_height_mm": round(layout_span_height, 2),
        "layout_span_area_mm2": round(layout_span_area, 2),
        "layout_compactness_pct": round((used_bbox / layout_span_area * 100.0) if layout_span_area > 0.0 else 0.0, 2),
        "remaining_net_area_mm2": round(max(0.0, area_mm2 - used_net), 2),
        "remaining_bbox_area_mm2": round(remaining_bbox_area, 2),
        "largest_edge_remnant_area_mm2": round(max(0.0, largest_edge_remnant_area), 2),
        "fragmented_remainder_area_mm2": round(max(0.0, remaining_bbox_area - largest_edge_remnant_area), 2),
        "geometry_validation": _sheet_overlap_diagnostics(
            {
                "placements": placements,
                "sheet_width_mm": width_mm,
                "sheet_height_mm": height_mm,
                "sheet_outer_polygons": list(sheet.get("sheet_outer_polygons", []) or []),
                "sheet_hole_polygons": list(sheet.get("sheet_hole_polygons", []) or []),
            }
        ),
    }


def _build_summary_base(
    expanded: list[dict[str, Any]],
    *,
    selected_profile: dict[str, Any],
    selection_mode: str,
    strategy_name: str,
    shape_grid_mm: float = 0.0,
) -> dict[str, Any]:
    return {
        "sheet_width_mm": round(_as_float(selected_profile.get("width_mm", 0.0), 0.0), 3),
        "sheet_height_mm": round(_as_float(selected_profile.get("height_mm", 0.0), 0.0), 3),
        "sheet_area_mm2": round(_as_float(selected_profile.get("area_mm2", 0.0), 0.0), 2),
        "selected_sheet_profile": dict(selected_profile or {}),
        "selection_mode": selection_mode,
        "strategy_name": strategy_name,
        "shape_aware": bool(str(strategy_name or "").strip().lower().startswith("shape")),
        "shape_grid_mm": round(_as_float(shape_grid_mm, 0.0), 3),
        "sheet_count": 0,
        "stock_sheet_count": 0,
        "remnant_sheet_count": 0,
        "purchased_sheet_count": 0,
        "part_count_requested": len(expanded),
        "part_count_placed": 0,
        "part_count_unplaced": 0,
        "used_net_area_mm2": 0.0,
        "used_bbox_area_mm2": 0.0,
        "layout_span_area_mm2": 0.0,
        "layout_compactness_pct": 0.0,
        "stock_sheet_area_mm2": 0.0,
        "purchase_sheet_area_mm2": 0.0,
        "total_sheet_area_mm2": 0.0,
        "utilization_net_pct": 0.0,
        "utilization_bbox_pct": 0.0,
        "waste_net_pct": 0.0,
        "waste_bbox_pct": 0.0,
        "remaining_net_area_mm2": 0.0,
        "remaining_bbox_area_mm2": 0.0,
        "largest_edge_remnant_area_mm2": 0.0,
        "fragmented_remainder_area_mm2": 0.0,
    }


def _finalize_result(
    items: list[NestItem],
    expanded: list[dict[str, Any]],
    sheets: list[dict[str, Any]],
    unplaced: list[dict[str, Any]],
    *,
    settings: dict[str, Any],
    warnings: list[str],
    selected_profile: dict[str, Any],
    selection_mode: str,
    strategy_name: str,
    shape_grid_mm: float = 0.0,
) -> dict[str, Any]:
    def _sheet_sequence_key(sheet: dict[str, Any]) -> tuple[float, ...]:
        profile = dict(sheet.get("profile", {}) or {})
        source_kind = str(profile.get("source_kind", "purchase") or "purchase").strip().lower()
        source_rank = 0.0 if source_kind in {"stock", "retalho"} else 1.0
        area = max(
            1.0,
            _as_float(profile.get("width_mm", 0.0), 0.0)
            * _as_float(profile.get("height_mm", 0.0), 0.0),
        )
        utilization = _as_float(sheet.get("used_net_area_mm2", 0.0), 0.0) / area
        return (source_rank, -utilization, -_as_float(sheet.get("used_bbox_area_mm2", 0.0), 0.0))

    sheets = sorted(list(sheets or []), key=_sheet_sequence_key)
    for sheet_index, sheet in enumerate(sheets):
        if sheet_index == 0:
            sheet["opening_reason"] = "Chapa prioritária: recebe primeiro todas as peças compatíveis."
        else:
            sheet["opening_reason"] = (
                "Chapa complementar: as peças restantes já não tinham posição válida nas chapas mais preenchidas "
                "com as margens, rotações e colisões atuais."
            )
    sheet_rows: list[dict[str, Any]] = []
    summary = _build_summary_base(
        expanded,
        selected_profile=selected_profile,
        selection_mode=selection_mode,
        strategy_name=strategy_name,
        shape_grid_mm=shape_grid_mm,
    )

    for index, sheet in enumerate(sheets):
        row = _build_sheet_row(sheet, index)
        if not row["placements"]:
            continue
        geometry_validation = dict(row.get("geometry_validation", {}) or {})
        solid_overlap_pair_count = int(geometry_validation.get("solid_overlap_pair_count", 0) or 0)
        part_in_part_pair_count = int(geometry_validation.get("part_in_part_pair_count", 0) or 0)
        spacing_violation_count = int(geometry_validation.get("spacing_violation_pair_count", 0) or 0)
        edge_violation_count = int(geometry_validation.get("edge_violation_count", 0) or 0)
        if solid_overlap_pair_count > 0:
            pair_labels = ", ".join(f"{left}/{right}" for left, right in list(geometry_validation.get("solid_overlap_pairs", []) or [])[:6])
            warnings.append(
                f"Chapa {index + 1}: foram detetadas {solid_overlap_pair_count} colisoes geometricas reais entre pecas ({pair_labels})."
            )
        elif part_in_part_pair_count > 0:
            warnings.append(
                f"Chapa {index + 1}: foram detetados {part_in_part_pair_count} encaixes internos por contorno (part-in-part), sem sobreposicao real de geometria."
            )
        if spacing_violation_count > 0:
            warnings.append(f"Chapa {index + 1}: foram detetadas {spacing_violation_count} violacoes da distancia minima entre pecas.")
        if edge_violation_count > 0:
            warnings.append(f"Chapa {index + 1}: foram detetadas {edge_violation_count} pecas fora da margem minima da chapa.")
        sheet_rows.append(row)
        summary["sheet_count"] += 1
        summary["part_count_placed"] += int(row.get("part_count", 0) or 0)
        summary["used_net_area_mm2"] += _as_float(row.get("used_net_area_mm2", 0.0), 0.0)
        summary["used_bbox_area_mm2"] += _as_float(row.get("used_bbox_area_mm2", 0.0), 0.0)
        summary["layout_span_area_mm2"] += _as_float(row.get("layout_span_area_mm2", 0.0), 0.0)
        summary["largest_edge_remnant_area_mm2"] += _as_float(row.get("largest_edge_remnant_area_mm2", 0.0), 0.0)
        summary["fragmented_remainder_area_mm2"] += _as_float(row.get("fragmented_remainder_area_mm2", 0.0), 0.0)
        summary["geometry_solid_overlap_pair_count"] = int(summary.get("geometry_solid_overlap_pair_count", 0) or 0) + solid_overlap_pair_count
        summary["geometry_part_in_part_pair_count"] = int(summary.get("geometry_part_in_part_pair_count", 0) or 0) + part_in_part_pair_count
        summary["geometry_spacing_violation_pair_count"] = int(summary.get("geometry_spacing_violation_pair_count", 0) or 0) + spacing_violation_count
        summary["geometry_edge_violation_count"] = int(summary.get("geometry_edge_violation_count", 0) or 0) + edge_violation_count
        sheet_min_part_distance = geometry_validation.get("min_part_distance_mm")
        if sheet_min_part_distance is not None:
            current_min = summary.get("geometry_min_part_distance_mm")
            summary["geometry_min_part_distance_mm"] = round(
                float(sheet_min_part_distance) if current_min is None else min(float(current_min), float(sheet_min_part_distance)),
                3,
            )
        sheet_min_edge_distance = geometry_validation.get("min_edge_distance_mm")
        if sheet_min_edge_distance is not None:
            current_min = summary.get("geometry_min_edge_distance_mm")
            summary["geometry_min_edge_distance_mm"] = round(
                float(sheet_min_edge_distance) if current_min is None else min(float(current_min), float(sheet_min_edge_distance)),
                3,
            )
        area_mm2 = _as_float(row.get("sheet_area_mm2", 0.0), 0.0)
        summary["total_sheet_area_mm2"] += area_mm2
        source_kind = str(row.get("source_kind", "") or "").strip().lower()
        if source_kind == "retalho":
            summary["remnant_sheet_count"] += 1
            summary["stock_sheet_count"] += 1
            summary["stock_sheet_area_mm2"] += area_mm2
        elif source_kind == "stock":
            summary["stock_sheet_count"] += 1
            summary["stock_sheet_area_mm2"] += area_mm2
        else:
            summary["purchased_sheet_count"] += 1
            summary["purchase_sheet_area_mm2"] += area_mm2

    summary["part_count_unplaced"] = len(unplaced)
    total_area = _as_float(summary.get("total_sheet_area_mm2", 0.0), 0.0)
    if total_area > 0.0:
        summary["utilization_net_pct"] = round(summary["used_net_area_mm2"] / total_area * 100.0, 2)
        summary["utilization_bbox_pct"] = round(summary["used_bbox_area_mm2"] / total_area * 100.0, 2)
    summary["used_net_area_mm2"] = round(summary["used_net_area_mm2"], 2)
    summary["used_bbox_area_mm2"] = round(summary["used_bbox_area_mm2"], 2)
    summary["layout_span_area_mm2"] = round(summary["layout_span_area_mm2"], 2)
    summary["largest_edge_remnant_area_mm2"] = round(summary["largest_edge_remnant_area_mm2"], 2)
    summary["fragmented_remainder_area_mm2"] = round(summary["fragmented_remainder_area_mm2"], 2)
    summary["stock_sheet_area_mm2"] = round(summary["stock_sheet_area_mm2"], 2)
    summary["purchase_sheet_area_mm2"] = round(summary["purchase_sheet_area_mm2"], 2)
    summary["total_sheet_area_mm2"] = round(total_area, 2)
    span_area = _as_float(summary.get("layout_span_area_mm2", 0.0), 0.0)
    summary["layout_compactness_pct"] = round(
        (_as_float(summary.get("used_bbox_area_mm2", 0.0), 0.0) / span_area * 100.0) if span_area > 0.0 else 0.0,
        2,
    )
    summary["waste_net_pct"] = round(max(0.0, 100.0 - _as_float(summary.get("utilization_net_pct", 0.0), 0.0)), 2)
    summary["waste_bbox_pct"] = round(max(0.0, 100.0 - _as_float(summary.get("utilization_bbox_pct", 0.0), 0.0)), 2)
    summary["remaining_net_area_mm2"] = round(max(0.0, total_area - _as_float(summary.get("used_net_area_mm2", 0.0), 0.0)), 2)
    summary["remaining_bbox_area_mm2"] = round(max(0.0, total_area - _as_float(summary.get("used_bbox_area_mm2", 0.0), 0.0)), 2)
    summary["optimization_level"] = _optimization_level()
    summary["optimization_time_limit_s"] = round(
        max(0.0, float(getattr(_NESTING_RUN_CONTEXT, "time_limit_s", 0.0) or 0.0)),
        1,
    )
    summary["optimization_elapsed_s"] = round(
        max(0.0, time.monotonic() - float(getattr(_NESTING_RUN_CONTEXT, "started_at", time.monotonic()) or time.monotonic())),
        3,
    )
    summary.update(_material_estimate(items, summary, settings))
    return {
        "sheets": sheet_rows,
        "summary": summary,
        "warnings": _unique_texts(warnings),
        "unplaced": list(unplaced),
    }


def _pack_profile(
    items: list[NestItem],
    expanded: list[dict[str, Any]],
    *,
    profile: dict[str, Any],
    part_spacing_mm: float,
    edge_margin_mm: float,
    allow_rotate: bool,
    settings: dict[str, Any],
    base_warnings: list[str],
    selection_mode: str,
) -> dict[str, Any]:
    normalized_profile = _normalize_sheet_profile(profile, 0)
    if normalized_profile is None:
        raise ValueError("Seleciona um formato de chapa valido.")
    usable_width, usable_height = _profile_usable_dimensions(normalized_profile, edge_margin_mm)
    best_result: dict[str, Any] | None = None
    for strategy_name, ordered_rows in _strategy_order_variants(
        expanded,
        shape_mode=False,
        free_angles=False,
    ):
        if best_result is not None and _optimization_deadline_reached():
            break
        warnings = list(base_warnings or [])
        sheets: list[dict[str, Any]] = []
        unplaced: list[dict[str, Any]] = []

        for row in ordered_rows:
            item: NestItem = row["item"]
            best_candidate: dict[str, Any] | None = None
            target_sheet: dict[str, Any] | None = None
            for sheet_index, sheet in enumerate(sheets):
                candidate = _try_place_on_sheet(
                    sheet,
                    item,
                    usable_width=usable_width,
                    usable_height=usable_height,
                    part_spacing_mm=part_spacing_mm,
                    allow_rotate=allow_rotate,
                )
                if candidate is None:
                    continue
                candidate_score = (
                    0.0,
                    float(sheet_index),
                    *_placement_score(candidate, strategy_name=strategy_name, sheet=sheet, edge_margin_mm=edge_margin_mm),
                )
                if best_candidate is None or candidate_score < best_candidate["score"]:
                    best_candidate = {"placement": candidate, "score": candidate_score}
                    target_sheet = sheet
            if best_candidate is None:
                candidate_sheet = _new_sheet(normalized_profile)
                candidate = _try_place_on_sheet(
                    candidate_sheet,
                    item,
                    usable_width=usable_width,
                    usable_height=usable_height,
                    part_spacing_mm=part_spacing_mm,
                    allow_rotate=allow_rotate,
                )
                if candidate is None:
                    warnings.append(f"{item.file_name}: nao foi possivel posicionar na chapa configurada.")
                    unplaced.append(
                        {
                            "ref_externa": item.ref_externa,
                            "description": item.description,
                            "file_name": item.file_name,
                            "copy_index": int(row.get("copy_index", 0) or 0),
                        }
                    )
                    continue
                target_sheet = candidate_sheet
                sheets.append(candidate_sheet)
                best_candidate = {
                    "placement": candidate,
                    "score": (
                        1.0,
                        *_placement_score(candidate, strategy_name=strategy_name, sheet=candidate_sheet, edge_margin_mm=edge_margin_mm),
                        float(len(sheets) - 1),
                    ),
                }
            if target_sheet is not None and best_candidate is not None:
                _apply_placement(target_sheet, best_candidate["placement"], row, edge_margin_mm=edge_margin_mm, part_spacing_mm=part_spacing_mm)

        strategy_result = _finalize_result(
            items,
            expanded,
            sheets,
            unplaced,
            settings=settings,
            warnings=warnings,
            selected_profile=normalized_profile,
            selection_mode=selection_mode,
            strategy_name=strategy_name,
        )
        if best_result is None or _result_score(strategy_result) < _result_score(best_result):
            best_result = strategy_result

    return best_result or _finalize_result(
        items,
        expanded,
        [],
        [],
        settings=settings,
        warnings=list(base_warnings or []),
        selected_profile=normalized_profile,
        selection_mode=selection_mode,
        strategy_name="",
    )


def _shape_engine_feasible(
    *,
    items: list[NestItem],
    profiles: list[dict[str, Any]],
    stock_candidates: list[dict[str, Any]],
    edge_margin_mm: float,
    grid_mm: float,
) -> tuple[bool, str]:
    safe_grid = max(2.0, _as_float(grid_mm, 10.0), 2.0)
    candidates = list(profiles or []) + list(stock_candidates or [])
    if not candidates:
        return False, "Sem formatos de chapa para avaliar."
    part_count = sum(max(1, int(item.qty or 0)) for item in list(items or []))
    if SHAPELY_AVAILABLE:
        if part_count > 1_000:
            return False, "Quantidade de pecas acima do limite de calculo poligonal interativo."
        return True, ""
    max_cells = 0
    for profile in candidates:
        try:
            usable_width, usable_height = _profile_usable_dimensions(profile, edge_margin_mm)
        except Exception:
            continue
        cells = int(math.ceil(usable_width / safe_grid) * math.ceil(usable_height / safe_grid))
        max_cells = max(max_cells, cells)
    if max_cells > 360_000:
        return False, f"Grelha de {safe_grid:g} mm demasiado fina para a chapa configurada."
    if part_count > 400:
        return False, "Quantidade de pecas demasiado elevada para o modo por contorno nesta fase."
    geometry_rows = []
    for item in list(items or []):
        point_count = sum(len(polygon) for polygon in list(item.outer_polygons or ()))
        point_count += sum(len(polygon) for polygon in list(item.hole_polygons or ()))
        geometry_rows.append(
            {
                "points": point_count,
                "holes": len(list(item.hole_polygons or ())),
                "qty": max(1, int(item.qty or 0)),
            }
        )
    max_item_points = max((int(row["points"]) for row in geometry_rows), default=0)
    hole_count = sum(int(row["holes"]) for row in geometry_rows)
    weighted_points = sum(int(row["points"]) * min(int(row["qty"]), 20) for row in geometry_rows)
    if part_count > 24 and (max_item_points > 1_500 or hole_count > 48 or weighted_points > 30_000):
        return False, (
            "Geometria detalhada e repetitiva acima do limite de calculo interativo; "
            "sera usado o modo retangular conservador."
        )
    return True, ""


def _effective_shape_grid_mm(raw_grid_mm: float, part_spacing_mm: float, edge_margin_mm: float) -> float:
    safe_grid = max(2.0, _as_float(raw_grid_mm, 10.0), 2.0)
    positive_limits = [
        max(2.0, _as_float(candidate_mm, 0.0), 2.0)
        for candidate_mm in (part_spacing_mm, edge_margin_mm)
        if _as_float(candidate_mm, 0.0) > 0
    ]
    if not positive_limits:
        return safe_grid
    return round(min(safe_grid, min(positive_limits)), 4)


def _strategy_names(*, shape_mode: bool, free_angles: bool, part_count: int) -> tuple[str, ...]:
    normalized_count = max(0, int(part_count or 0))
    level = _optimization_level()
    if shape_mode:
        if free_angles:
            base = ("shape-retalho", "shape-compact")
        elif normalized_count <= 8:
            base = ("shape-retalho", "shape-compact", "shape-area")
        else:
            base = ("shape-retalho", "shape-area", "shape-height-first")
        if level in {"tap2", "tap3"}:
            base += ("shape-width-first", "shape-compact")
        return tuple(dict.fromkeys(base))
    if normalized_count <= 8:
        base = ("retalho", "compact", "area")
    else:
        base = ("retalho", "area", "height-first")
    if level in {"tap2", "tap3"}:
        base += ("width-first", "compact", "longest-side")
    return tuple(dict.fromkeys(base))


def _strategy_order_variants(
    expanded: list[dict[str, Any]],
    *,
    shape_mode: bool,
    free_angles: bool,
) -> Iterator[tuple[str, list[dict[str, Any]]]]:
    """Build deterministic multi-start orders without losing user priorities."""
    rows = list(expanded or [])
    for name in _strategy_names(shape_mode=shape_mode, free_angles=free_angles, part_count=len(rows)):
        yield name, sorted(rows, key=_strategy_sort_key(name), reverse=True)
    level = _optimization_level()
    time_limit_s = float(getattr(_NESTING_RUN_CONTEXT, "time_limit_s", 0.0) or 0.0)
    if level == "tap1" or time_limit_s <= 0.0 or len(rows) <= 1:
        return

    seed_parts = [
        (
            str(getattr(row.get("item"), "ref_externa", "") or ""),
            round(_as_float(row.get("bbox_area_mm2", 0.0), 0.0), 3),
            int(row.get("copy_index", 0) or 0),
        )
        for row in rows
    ]
    seed = sum(
        (index + 1) * (sum(ord(char) for char in ref) + int(area) + copy_index)
        for index, (ref, area, copy_index) in enumerate(seed_parts)
    )
    rng = random.Random(seed)
    prefix = "shape-" if shape_mode else ""
    attempt = 0
    while attempt < 1_000_000:
        if _optimization_deadline_reached():
            break
        attempt += 1
        randomized = list(rows)
        rng.shuffle(randomized)
        # Prioridades críticas continuam sempre antes das normais/baixas;
        # a exploração altera apenas a sequência dentro do mesmo nível.
        randomized.sort(
            key=lambda row: max(
                -1,
                min(2, _as_int(getattr(row.get("item"), "priority", row.get("priority", 0)), 0)),
            ),
            reverse=True,
        )
        yield f"{prefix}explore-{attempt}", randomized


def _shape_anchor_positions(
    sheet: dict[str, Any],
    *,
    sheet_width_cells: int,
    sheet_height_cells: int,
    mask_width_cells: int,
    mask_height_cells: int,
    grid_mm: float,
    edge_margin_mm: float,
) -> list[tuple[int, int]]:
    placements = [dict(row or {}) for row in list(sheet.get("placements", []) or []) if isinstance(row, dict)]
    if not placements:
        return [(0, 0)]
    x_positions: set[int] = {0}
    y_positions: set[int] = {0}
    for placement in placements:
        start_x = max(0, int(math.floor((_as_float(placement.get("x_mm", 0.0), 0.0) - max(0.0, edge_margin_mm)) / grid_mm)))
        start_y = max(0, int(math.floor((_as_float(placement.get("y_mm", 0.0), 0.0) - max(0.0, edge_margin_mm)) / grid_mm)))
        end_x = max(0, int(math.ceil((_as_float(placement.get("x_mm", 0.0), 0.0) - max(0.0, edge_margin_mm) + _as_float(placement.get("width_mm", 0.0), 0.0)) / grid_mm)))
        end_y = max(0, int(math.ceil((_as_float(placement.get("y_mm", 0.0), 0.0) - max(0.0, edge_margin_mm) + _as_float(placement.get("height_mm", 0.0), 0.0)) / grid_mm)))
        for candidate_x in (start_x, start_x + 1, max(0, end_x - 1), end_x, end_x + 1):
            x_positions.add(candidate_x)
        for candidate_y in (start_y, start_y + 1, max(0, end_y - 1), end_y, end_y + 1):
            y_positions.add(candidate_y)
    max_grid_x = max(0, sheet_width_cells - mask_width_cells)
    max_grid_y = max(0, sheet_height_cells - mask_height_cells)
    filtered_x = sorted(value for value in x_positions if 0 <= value <= max_grid_x)
    filtered_y = sorted(value for value in y_positions if 0 <= value <= max_grid_y)
    return [(grid_x, grid_y) for grid_y in filtered_y for grid_x in filtered_x]


def _polygon_has_non_orthogonal_edges(polygon: tuple[tuple[float, float], ...] | list[tuple[float, float]], tol: float = 1e-3) -> bool:
    raw = list(polygon or [])
    if len(raw) < 2:
        return False
    for index in range(len(raw)):
        x1, y1 = raw[index]
        x2, y2 = raw[(index + 1) % len(raw)]
        if abs(float(x1) - float(x2)) > tol and abs(float(y1) - float(y2)) > tol:
            return True
    return False


def _item_benefits_from_free_angles(item: NestItem) -> bool:
    if not list(item.outer_polygons or []):
        return False
    return any(_polygon_has_non_orthogonal_edges(polygon) for polygon in list(item.outer_polygons or []))


def _effective_free_angle_rotation(items: list[NestItem], expanded: list[dict[str, Any]], requested: bool) -> tuple[bool, str]:
    if not requested:
        return False, ""
    total_parts = max(0, len(list(expanded or [])))
    if total_parts > 2:
        return False, "Rotações livres desativadas automaticamente neste estudo para evitar tempos excessivos; mantidas rotações 0°/90°."
    if not any(_item_benefits_from_free_angles(item) for item in list(items or [])):
        return False, "Rotações livres desativadas automaticamente porque as peças deste grupo não beneficiam de ângulos intermédios."
    return True, ""


def _shapely_polygon_parts(geometry: Any) -> list[Any]:
    if geometry is None or bool(getattr(geometry, "is_empty", True)):
        return []
    geom_type = str(getattr(geometry, "geom_type", "") or "")
    if geom_type == "Polygon":
        return [geometry]
    if geom_type in {"MultiPolygon", "GeometryCollection"}:
        return [part for child in list(getattr(geometry, "geoms", ()) or ()) for part in _shapely_polygon_parts(child)]
    return []


def _shapely_normalize_geometry(geometry: Any) -> Any:
    if geometry is None or bool(getattr(geometry, "is_empty", True)):
        return geometry
    min_x, min_y, _, _ = geometry.bounds
    return shapely_affinity.translate(geometry, xoff=-float(min_x), yoff=-float(min_y))


def _shapely_geometry_from_polygons(
    outer_polygons: Any,
    hole_polygons: Any,
    *,
    fallback_width: float = 0.0,
    fallback_height: float = 0.0,
) -> Any:
    if not SHAPELY_AVAILABLE:
        return None
    holes = [tuple((float(x), float(y)) for x, y in list(raw or ())) for raw in list(hole_polygons or ()) if len(list(raw or ())) >= 3]
    parts: list[Any] = []
    for raw_outer in list(outer_polygons or ()):
        shell = tuple((float(x), float(y)) for x, y in list(raw_outer or ()))
        if len(shell) < 3:
            continue
        shell_only = ShapelyPolygon(shell)
        assigned_holes = []
        for hole in holes:
            try:
                if shell_only.covers(ShapelyPolygon(hole).representative_point()):
                    assigned_holes.append(hole)
            except Exception:
                continue
        polygon = ShapelyPolygon(shell, assigned_holes)
        if not polygon.is_valid:
            polygon = shapely_make_valid(polygon)
        parts.extend(_shapely_polygon_parts(polygon))
    if not parts and fallback_width > 0.0 and fallback_height > 0.0:
        parts = [shapely_box(0.0, 0.0, float(fallback_width), float(fallback_height))]
    if not parts:
        return None
    geometry = shapely_unary_union(parts)
    if not geometry.is_valid:
        geometry = shapely_make_valid(geometry)
    polygon_parts = _shapely_polygon_parts(geometry)
    if not polygon_parts:
        return None
    return _shapely_normalize_geometry(shapely_unary_union(polygon_parts))


def _shapely_item_geometry(item: NestItem) -> Any:
    return _shapely_geometry_from_polygons(
        item.outer_polygons,
        item.hole_polygons,
        fallback_width=item.bbox_width_mm,
        fallback_height=item.bbox_height_mm,
    )


def _shapely_geometry_payload(geometry: Any) -> tuple[tuple[tuple[tuple[float, float], ...], ...], tuple[tuple[tuple[float, float], ...], ...]]:
    outer: list[tuple[tuple[float, float], ...]] = []
    holes: list[tuple[tuple[float, float], ...]] = []
    for polygon in _shapely_polygon_parts(geometry):
        shell = tuple((round(float(x), 3), round(float(y), 3)) for x, y in list(polygon.exterior.coords)[:-1])
        if len(shell) >= 3:
            outer.append(shell)
        for ring in list(polygon.interiors or ()):
            hole = tuple((round(float(x), 3), round(float(y), 3)) for x, y in list(ring.coords)[:-1])
            if len(hole) >= 3:
                holes.append(hole)
    return tuple(outer), tuple(holes)


def _shapely_orientation_variants(
    item: NestItem,
    *,
    allow_rotate: bool,
    allow_mirror: bool,
    free_angles: bool,
    base_cache: dict[str, Any],
    variant_cache: dict[tuple, list[dict[str, Any]]],
) -> list[dict[str, Any]]:
    cache_key = str(item.shape_cache_key or item.path or id(item))
    key = (cache_key, bool(allow_rotate), bool(allow_mirror), bool(free_angles), str(item.rotation_policy or "auto"))
    cached = variant_cache.get(key)
    if cached is not None:
        return cached
    geometry = base_cache.get(cache_key)
    if geometry is None:
        geometry = _shapely_item_geometry(item)
        base_cache[cache_key] = geometry
    if geometry is None:
        variant_cache[key] = []
        return []
    variants: list[dict[str, Any]] = []
    seen: set[tuple[float, float, float, bool]] = set()
    for orientation in _candidate_orientations(item, allow_rotate, allow_mirror=allow_mirror, free_angles=free_angles):
        angle = float(orientation.get("angle_deg", 90.0 if bool(orientation.get("rotated")) else 0.0) or 0.0)
        mirrored = bool(orientation.get("mirrored"))
        transformed = geometry
        if mirrored:
            transformed = shapely_affinity.scale(transformed, xfact=-1.0, yfact=1.0, origin=(0.0, 0.0))
        if abs(angle) > 1e-9:
            transformed = shapely_affinity.rotate(transformed, angle, origin=(0.0, 0.0), use_radians=False)
        transformed = _shapely_normalize_geometry(transformed)
        min_x, min_y, max_x, max_y = transformed.bounds
        width = max(0.0, float(max_x) - float(min_x))
        height = max(0.0, float(max_y) - float(min_y))
        signature = (round(width, 3), round(height, 3), round(angle % 180.0, 3), mirrored)
        if signature in seen:
            continue
        if any(bool(transformed.equals_exact(existing["geometry"], tolerance=1e-4)) for existing in variants):
            continue
        seen.add(signature)
        outer, holes = _shapely_geometry_payload(transformed)
        variants.append(
            {
                "geometry": transformed,
                "width": width,
                "height": height,
                "rotated": bool(orientation.get("rotated")),
                "mirrored": mirrored,
                "rotation_deg": angle,
                "shape_outer_polygons": outer,
                "shape_hole_polygons": holes,
            }
        )
    variant_cache[key] = variants
    return variants


def _shapely_sheet_allowed_geometry(profile: dict[str, Any], edge_margin_mm: float) -> Any:
    width = max(0.0, _as_float(profile.get("width_mm", 0.0), 0.0))
    height = max(0.0, _as_float(profile.get("height_mm", 0.0), 0.0))
    outer = tuple(profile.get("outer_polygons", ()) or ())
    holes = tuple(profile.get("hole_polygons", ()) or ())
    if outer:
        geometry = _shapely_geometry_from_polygons(outer, holes, fallback_width=width, fallback_height=height)
        if geometry is not None and edge_margin_mm > 0.0:
            geometry = geometry.buffer(-float(edge_margin_mm), join_style="mitre")
        return geometry
    if width <= 2.0 * edge_margin_mm or height <= 2.0 * edge_margin_mm:
        return None
    return shapely_box(edge_margin_mm, edge_margin_mm, width - edge_margin_mm, height - edge_margin_mm)


def _new_shapely_sheet(profile: dict[str, Any], edge_margin_mm: float) -> dict[str, Any]:
    sheet = _new_sheet(profile)
    sheet["_shapely_allowed"] = _shapely_sheet_allowed_geometry(profile, edge_margin_mm)
    sheet["_shapely_actual"] = []
    return sheet


def _shapely_candidate_positions(sheet: dict[str, Any], variant: dict[str, Any], part_spacing_mm: float) -> list[tuple[float, float]]:
    allowed = sheet.get("_shapely_allowed")
    if allowed is None or bool(getattr(allowed, "is_empty", True)):
        return []
    min_x, min_y, max_x, max_y = allowed.bounds
    width = float(variant.get("width", 0.0) or 0.0)
    height = float(variant.get("height", 0.0) or 0.0)
    spacing = max(0.0, float(part_spacing_mm or 0.0))
    xs: set[float] = {round(float(min_x), 4), round(float(max_x) - width, 4)}
    ys: set[float] = {round(float(min_y), 4), round(float(max_y) - height, 4)}
    for geometry in list(sheet.get("_shapely_actual", []) or []):
        left, bottom, right, top = geometry.bounds
        xs.update((round(float(right) + spacing, 4), round(float(left) - spacing - width, 4)))
        ys.update((round(float(top) + spacing, 4), round(float(bottom) - spacing - height, 4)))
        for polygon in _shapely_polygon_parts(geometry):
            for ring in list(polygon.interiors or ()):
                hole_left, hole_bottom, hole_right, hole_top = ring.bounds
                if (
                    float(hole_right) - float(hole_left) < width + (2.0 * spacing) - 1e-6
                    or float(hole_top) - float(hole_bottom) < height + (2.0 * spacing) - 1e-6
                ):
                    continue
                xs.update(
                    (
                        round(float(hole_left) + spacing, 4),
                        round(float(hole_right) - spacing - width, 4),
                        round(((float(hole_left) + float(hole_right) - width) / 2.0), 4),
                    )
                )
                ys.update(
                    (
                        round(float(hole_bottom) + spacing, 4),
                        round(float(hole_top) - spacing - height, 4),
                        round(((float(hole_bottom) + float(hole_top) - height) / 2.0), 4),
                    )
                )
    valid_x = sorted(value for value in xs if value >= min_x - 1e-6 and value + width <= max_x + 1e-6)
    valid_y = sorted(value for value in ys if value >= min_y - 1e-6 and value + height <= max_y + 1e-6)
    return [(x, y) for y in valid_y for x in valid_x]


def _try_place_on_shapely_sheet(
    sheet: dict[str, Any],
    item: NestItem,
    *,
    allow_rotate: bool,
    allow_mirror: bool,
    free_angles: bool,
    part_spacing_mm: float,
    edge_margin_mm: float,
    strategy_name: str,
    base_cache: dict[str, Any],
    variant_cache: dict[tuple, list[dict[str, Any]]],
) -> dict[str, Any] | None:
    allowed = sheet.get("_shapely_allowed")
    if allowed is None or bool(getattr(allowed, "is_empty", True)):
        return None
    actual_geometries = list(sheet.get("_shapely_actual", []) or [])
    tree = STRtree(actual_geometries) if actual_geometries else None
    spacing = max(0.0, float(part_spacing_mm or 0.0))
    placed_max_x = max((float(geometry.bounds[2]) for geometry in actual_geometries), default=float(allowed.bounds[0]))
    placed_max_y = max((float(geometry.bounds[3]) for geometry in actual_geometries), default=float(allowed.bounds[1]))
    best: dict[str, Any] | None = None
    normalized_strategy = str(strategy_name or "").lower()
    for variant in _shapely_orientation_variants(
        item,
        allow_rotate=allow_rotate,
        allow_mirror=allow_mirror,
        free_angles=free_angles,
        base_cache=base_cache,
        variant_cache=variant_cache,
    ):
        positions = _shapely_candidate_positions(sheet, variant, spacing)
        if "width" in normalized_strategy:
            positions.sort(key=lambda point: (point[0], point[1]))
        for x, y in positions:
            candidate = shapely_affinity.translate(variant["geometry"], xoff=x, yoff=y)
            if not allowed.covers(candidate):
                continue
            probe = candidate.buffer(spacing + 1e-6, quad_segs=2, join_style="mitre") if spacing > 0.0 else candidate
            nearby = list(tree.query(probe)) if tree is not None else []
            collision = False
            for raw_index in nearby:
                other = actual_geometries[int(raw_index)]
                if candidate.intersection(other).area > 1e-5 or candidate.distance(other) < spacing - 1e-4:
                    collision = True
                    break
            if collision:
                continue
            projected_max_x = max(placed_max_x, x + float(variant["width"]))
            projected_max_y = max(placed_max_y, y + float(variant["height"]))
            score = (
                projected_max_x if "width" in normalized_strategy else projected_max_y,
                projected_max_y if "width" in normalized_strategy else projected_max_x,
                y,
                x,
                1.0 if bool(variant.get("mirrored")) else 0.0,
                abs(float(variant.get("rotation_deg", 0.0) or 0.0)),
            )
            placement = {
                **{key: value for key, value in variant.items() if key != "geometry"},
                "x": round(x - float(edge_margin_mm), 4),
                "y": round(y - float(edge_margin_mm), 4),
                "place_w": float(variant["width"]),
                "place_h": float(variant["height"]),
                "shape_mode": "geos",
                "occupied_area_mm2": round(float(candidate.area), 2),
                "_actual_geometry": candidate,
                "score": score,
            }
            if best is None or score < best["score"]:
                best = placement
            if "compact" not in normalized_strategy:
                break
    return best


def _pack_profile_shapely(
    items: list[NestItem],
    expanded: list[dict[str, Any]],
    *,
    profile: dict[str, Any],
    part_spacing_mm: float,
    edge_margin_mm: float,
    allow_rotate: bool,
    allow_mirror: bool,
    free_angles: bool,
    settings: dict[str, Any],
    base_warnings: list[str],
    selection_mode: str,
) -> dict[str, Any] | None:
    if not SHAPELY_AVAILABLE:
        return None
    normalized_profile = _normalize_sheet_profile(profile, 0)
    if normalized_profile is None:
        raise ValueError("Seleciona um formato de chapa valido.")
    best_result: dict[str, Any] | None = None
    base_cache: dict[str, Any] = {}
    variant_cache: dict[tuple, list[dict[str, Any]]] = {}
    part_count = len(expanded)
    if part_count <= 16:
        strategies = ("shape-geos-area", "shape-geos-height", "shape-geos-width")
    else:
        # GEOS is the expensive exact validator. Large batches get a fast
        # multi-strategy bounding-box companion below, so one exact pass keeps
        # the interactive calculation responsive.
        strategies = ("shape-geos-area",)
    usable_width, usable_height = _profile_usable_dimensions(normalized_profile, edge_margin_mm)
    usable_area = max(1.0, usable_width * usable_height)
    theoretical_sheet_floor = max(
        1,
        int(math.ceil(sum(max(0.0, row["item"].net_area_mm2) for row in list(expanded or [])) / usable_area)),
    )
    for strategy_name in strategies:
        order_name = (
            "area"
            if strategy_name.endswith("area")
            else "width-first"
            if strategy_name.endswith("width")
            else "height-first"
        )
        ordered_rows = sorted(list(expanded or []), key=_strategy_sort_key(order_name), reverse=True)
        sheets: list[dict[str, Any]] = []
        unplaced: list[dict[str, Any]] = []
        warnings = list(base_warnings or [])
        for row in ordered_rows:
            item: NestItem = row["item"]
            best_candidate: dict[str, Any] | None = None
            target_sheet: dict[str, Any] | None = None
            for sheet_index, sheet in enumerate(sheets):
                candidate = _try_place_on_shapely_sheet(
                    sheet,
                    item,
                    allow_rotate=allow_rotate,
                    allow_mirror=allow_mirror,
                    free_angles=free_angles,
                    part_spacing_mm=part_spacing_mm,
                    edge_margin_mm=edge_margin_mm,
                    strategy_name=strategy_name,
                    base_cache=base_cache,
                    variant_cache=variant_cache,
                )
                if candidate is None:
                    continue
                score = (0.0, float(sheet_index), *tuple(candidate.get("score", ())))
                if best_candidate is None or score < best_candidate["global_score"]:
                    best_candidate = {**candidate, "global_score": score}
                    target_sheet = sheet
            if best_candidate is None:
                candidate_sheet = _new_shapely_sheet(normalized_profile, edge_margin_mm)
                candidate = _try_place_on_shapely_sheet(
                    candidate_sheet,
                    item,
                    allow_rotate=allow_rotate,
                    allow_mirror=allow_mirror,
                    free_angles=free_angles,
                    part_spacing_mm=part_spacing_mm,
                    edge_margin_mm=edge_margin_mm,
                    strategy_name=strategy_name,
                    base_cache=base_cache,
                    variant_cache=variant_cache,
                )
                if candidate is None:
                    unplaced.append(
                        {
                            "ref_externa": item.ref_externa,
                            "description": item.description,
                            "file_name": item.file_name,
                            "copy_index": int(row.get("copy_index", 0) or 0),
                        }
                    )
                    continue
                target_sheet = candidate_sheet
                sheets.append(candidate_sheet)
                best_candidate = candidate
            if target_sheet is not None and best_candidate is not None:
                actual_geometry = best_candidate.pop("_actual_geometry")
                best_candidate.pop("score", None)
                best_candidate.pop("global_score", None)
                _apply_placement(
                    target_sheet,
                    best_candidate,
                    row,
                    edge_margin_mm=edge_margin_mm,
                    part_spacing_mm=part_spacing_mm,
                )
                target_sheet.setdefault("_shapely_actual", []).append(actual_geometry)
        warnings.append(
            "Plano calculado por geometria poligonal GEOS, com validacao real de contornos, furos, margens e distancia entre pecas."
        )
        result = _finalize_result(
            items,
            expanded,
            sheets,
            unplaced,
            settings=settings,
            warnings=warnings,
            selected_profile=normalized_profile,
            selection_mode=selection_mode,
            strategy_name=strategy_name,
            shape_grid_mm=0.0,
        )
        if best_result is None or _result_score(result) < _result_score(best_result):
            best_result = result
        if _as_int(dict(result.get("summary", {}) or {}).get("sheet_count", 0), 0) <= theoretical_sheet_floor:
            break
    return best_result


def _pack_profile_shape(
    items: list[NestItem],
    expanded: list[dict[str, Any]],
    *,
    profile: dict[str, Any],
    part_spacing_mm: float,
    edge_margin_mm: float,
    allow_rotate: bool,
    allow_mirror: bool,
    free_angles: bool,
    settings: dict[str, Any],
    base_warnings: list[str],
    selection_mode: str,
    grid_mm: float,
) -> dict[str, Any]:
    geos_result = _pack_profile_shapely(
        items,
        expanded,
        profile=profile,
        part_spacing_mm=part_spacing_mm,
        edge_margin_mm=edge_margin_mm,
        allow_rotate=allow_rotate,
        allow_mirror=allow_mirror,
        free_angles=free_angles,
        settings=settings,
        base_warnings=base_warnings,
        selection_mode=selection_mode,
    )
    if geos_result is not None:
        return geos_result
    normalized_profile = _normalize_sheet_profile(profile, 0)
    if normalized_profile is None:
        raise ValueError("Seleciona um formato de chapa valido.")
    best_result: dict[str, Any] | None = None
    shape_cache: dict[tuple, dict[str, Any]] = {}
    for strategy_name, ordered_rows in _strategy_order_variants(
        expanded,
        shape_mode=True,
        free_angles=free_angles,
    ):
        if best_result is not None and _optimization_deadline_reached():
            break
        warnings = list(base_warnings or [])
        sheets: list[dict[str, Any]] = []
        unplaced: list[dict[str, Any]] = []

        for row in ordered_rows:
            item: NestItem = row["item"]
            best_candidate: dict[str, Any] | None = None
            target_sheet: dict[str, Any] | None = None
            for sheet_index, sheet in enumerate(sheets):
                candidate = _try_place_on_shape_sheet(
                    sheet,
                    item,
                    allow_rotate=allow_rotate,
                    allow_mirror=allow_mirror,
                    free_angles=free_angles,
                    grid_mm=grid_mm,
                    part_spacing_mm=part_spacing_mm,
                    edge_margin_mm=edge_margin_mm,
                    cache=shape_cache,
                    strategy_name=strategy_name,
                )
                if candidate is None:
                    continue
                candidate_score = (
                    0.0,
                    float(sheet_index),
                    *_shape_candidate_score(candidate, strategy_name=strategy_name, sheet=sheet, edge_margin_mm=edge_margin_mm),
                )
                if best_candidate is None or candidate_score < best_candidate["score"]:
                    best_candidate = {"placement": candidate, "score": candidate_score}
                    target_sheet = sheet
            if best_candidate is None:
                candidate_sheet = _new_shape_sheet(normalized_profile, edge_margin_mm=edge_margin_mm, grid_mm=grid_mm)
                candidate = _try_place_on_shape_sheet(
                    candidate_sheet,
                    item,
                    allow_rotate=allow_rotate,
                    allow_mirror=allow_mirror,
                    free_angles=free_angles,
                    grid_mm=grid_mm,
                    part_spacing_mm=part_spacing_mm,
                    edge_margin_mm=edge_margin_mm,
                    cache=shape_cache,
                    strategy_name=strategy_name,
                )
                if candidate is None:
                    warnings.append(f"{item.file_name}: nao foi possivel posicionar por contorno na chapa configurada.")
                    unplaced.append(
                        {
                            "ref_externa": item.ref_externa,
                            "description": item.description,
                            "file_name": item.file_name,
                            "copy_index": int(row.get("copy_index", 0) or 0),
                        }
                    )
                    continue
                target_sheet = candidate_sheet
                sheets.append(candidate_sheet)
                best_candidate = {
                    "placement": candidate,
                    "score": (
                        1.0,
                        *_shape_candidate_score(candidate, strategy_name=strategy_name, sheet=candidate_sheet, edge_margin_mm=edge_margin_mm),
                        float(len(sheets) - 1),
                    ),
                }
            if target_sheet is not None and best_candidate is not None:
                _apply_placement(target_sheet, best_candidate["placement"], row, edge_margin_mm=edge_margin_mm, part_spacing_mm=part_spacing_mm)

        strategy_result = _finalize_result(
            items,
            expanded,
            sheets,
            unplaced,
            settings=settings,
            warnings=warnings,
            selected_profile=normalized_profile,
            selection_mode=selection_mode,
            strategy_name=strategy_name,
            shape_grid_mm=grid_mm,
        )
        if best_result is None or _result_score(strategy_result) < _result_score(best_result):
            best_result = strategy_result

    return best_result or _finalize_result(
        items,
        expanded,
        [],
        [],
        settings=settings,
        warnings=list(base_warnings or []),
        selected_profile=normalized_profile,
        selection_mode=selection_mode,
        strategy_name="shape",
        shape_grid_mm=grid_mm,
    )


def _expand_stock_units(stock_candidates: list[dict[str, Any]]) -> list[dict[str, Any]]:
    units: list[dict[str, Any]] = []
    for candidate in list(stock_candidates or []):
        qty = max(0, _as_int(candidate.get("quantity_available", 0), 0))
        if qty <= 0:
            continue
        for copy_index in range(qty):
            unit = dict(candidate or {})
            unit["quantity_available"] = 1
            unit["unit_index"] = copy_index + 1
            if qty > 1:
                unit["source_label"] = f"{str(candidate.get('source_label', candidate.get('name', 'Stock')) or '').strip()} #{copy_index + 1}"
                unit["name"] = f"{str(candidate.get('name', candidate.get('source_label', 'Stock')) or '').strip()} #{copy_index + 1}"
            units.append(unit)
    units.sort(
        key=lambda row: (
            0 if str(row.get("source_kind", "") or "").strip().lower() == "retalho" else 1,
            _as_float(row.get("area_mm2", 0.0), 0.0),
            str(row.get("lote", "") or ""),
            str(row.get("material_id", "") or ""),
            _as_int(row.get("unit_index", 0), 0),
        )
    )
    return units


def _pack_with_stock(
    items: list[NestItem],
    expanded: list[dict[str, Any]],
    *,
    stock_candidates: list[dict[str, Any]],
    purchase_profile: dict[str, Any] | None,
    part_spacing_mm: float,
    edge_margin_mm: float,
    allow_rotate: bool,
    allow_purchase_fallback: bool,
    settings: dict[str, Any],
    base_warnings: list[str],
    selection_mode: str,
) -> dict[str, Any]:
    normalized_purchase = _normalize_sheet_profile(purchase_profile or {}, 0) if purchase_profile else None
    purchase_dims = _profile_usable_dimensions(normalized_purchase, edge_margin_mm) if normalized_purchase and allow_purchase_fallback else None
    normalized_stock = [profile for index, row in enumerate(list(stock_candidates or [])) if (profile := _normalize_stock_sheet_candidate(row, index)) is not None]
    selected_profile = normalized_purchase or {
        "name": "Apenas stock",
        "width_mm": 0.0,
        "height_mm": 0.0,
        "area_mm2": 0.0,
        "source_kind": "stock",
        "source_label": "Apenas stock",
    }
    best_result: dict[str, Any] | None = None
    for strategy_name, ordered_rows in _strategy_order_variants(
        expanded,
        shape_mode=False,
        free_angles=False,
    ):
        if best_result is not None and _optimization_deadline_reached():
            break
        warnings = list(base_warnings or [])
        sheets: list[dict[str, Any]] = []
        unplaced: list[dict[str, Any]] = []
        available_stock_units = _expand_stock_units(normalized_stock)

        for row in ordered_rows:
            item: NestItem = row["item"]
            best_candidate: dict[str, Any] | None = None
            target_sheet: dict[str, Any] | None = None

            for sheet_index, sheet in enumerate(sheets):
                profile = dict(sheet.get("profile", {}) or {})
                try:
                    usable_width, usable_height = _profile_usable_dimensions(profile, edge_margin_mm)
                except Exception:
                    continue
                candidate = _try_place_on_sheet(
                    sheet,
                    item,
                    usable_width=usable_width,
                    usable_height=usable_height,
                    part_spacing_mm=part_spacing_mm,
                    allow_rotate=allow_rotate,
                    strategy_name=strategy_name,
                    edge_margin_mm=edge_margin_mm,
                )
                if candidate is None:
                    continue
                candidate_score = (
                    0.0,
                    float(sheet_index),
                    *_placement_score(candidate, strategy_name=strategy_name, sheet=sheet, edge_margin_mm=edge_margin_mm),
                )
                if best_candidate is None or candidate_score < best_candidate["score"]:
                    best_candidate = {"placement": candidate, "score": candidate_score}
                    target_sheet = sheet

            stock_choice: dict[str, Any] | None = None
            if best_candidate is None:
                for stock_index, stock_profile in enumerate(available_stock_units):
                    try:
                        usable_width, usable_height = _profile_usable_dimensions(stock_profile, edge_margin_mm)
                    except Exception:
                        continue
                    candidate_sheet = _new_sheet(stock_profile)
                    candidate = _try_place_on_sheet(
                        candidate_sheet,
                        item,
                        usable_width=usable_width,
                        usable_height=usable_height,
                        part_spacing_mm=part_spacing_mm,
                        allow_rotate=allow_rotate,
                        strategy_name=strategy_name,
                        edge_margin_mm=edge_margin_mm,
                    )
                    if candidate is None:
                        continue
                    source_priority = 0.0 if str(stock_profile.get("source_kind", "") or "").strip().lower() == "retalho" else 1.0
                    candidate_score = (
                        1.0,
                        source_priority,
                        _as_float(stock_profile.get("area_mm2", 0.0), 0.0),
                        *_placement_score(candidate, strategy_name=strategy_name, sheet=candidate_sheet, edge_margin_mm=edge_margin_mm),
                        float(stock_index),
                    )
                    if best_candidate is None or candidate_score < best_candidate["score"]:
                        best_candidate = {"placement": candidate, "score": candidate_score}
                        stock_choice = {"index": stock_index, "profile": stock_profile}

            if best_candidate is None and normalized_purchase is not None and allow_purchase_fallback and purchase_dims is not None:
                candidate_sheet = _new_sheet(normalized_purchase)
                candidate = _try_place_on_sheet(
                    candidate_sheet,
                    item,
                    usable_width=purchase_dims[0],
                    usable_height=purchase_dims[1],
                    part_spacing_mm=part_spacing_mm,
                    allow_rotate=allow_rotate,
                    strategy_name=strategy_name,
                    edge_margin_mm=edge_margin_mm,
                )
                if candidate is not None:
                    best_candidate = {
                        "placement": candidate,
                        "score": (
                            2.0,
                            _as_float(normalized_purchase.get("area_mm2", 0.0), 0.0),
                            *_placement_score(candidate, strategy_name=strategy_name, sheet=candidate_sheet, edge_margin_mm=edge_margin_mm),
                        ),
                    }
                    target_sheet = candidate_sheet
                    sheets.append(candidate_sheet)

            if target_sheet is None and stock_choice is not None:
                target_sheet = _new_sheet(stock_choice["profile"])
                sheets.append(target_sheet)
                available_stock_units.pop(stock_choice["index"])

            if target_sheet is None or best_candidate is None:
                warnings.append(f"{item.file_name}: nao foi possivel posicionar com o stock/formato atual.")
                unplaced.append(
                    {
                        "ref_externa": item.ref_externa,
                        "description": item.description,
                        "file_name": item.file_name,
                        "copy_index": int(row.get("copy_index", 0) or 0),
                    }
                )
                continue

            _apply_placement(target_sheet, best_candidate["placement"], row, edge_margin_mm=edge_margin_mm, part_spacing_mm=part_spacing_mm)

        strategy_result = _finalize_result(
            items,
            expanded,
            sheets,
            unplaced,
            settings=settings,
            warnings=warnings,
            selected_profile=selected_profile,
            selection_mode=selection_mode,
            strategy_name=strategy_name,
            shape_grid_mm=grid_mm,
        )
        if best_result is None or _result_score(strategy_result) < _result_score(best_result):
            best_result = strategy_result

    return best_result or _finalize_result(
        items,
        expanded,
        [],
        [],
        settings=settings,
        warnings=list(base_warnings or []),
        selected_profile=selected_profile,
        selection_mode=selection_mode,
        strategy_name="",
    )


def _pack_with_stock_shapely(
    items: list[NestItem],
    expanded: list[dict[str, Any]],
    *,
    stock_candidates: list[dict[str, Any]],
    purchase_profile: dict[str, Any] | None,
    part_spacing_mm: float,
    edge_margin_mm: float,
    allow_rotate: bool,
    allow_mirror: bool,
    free_angles: bool,
    allow_purchase_fallback: bool,
    settings: dict[str, Any],
    base_warnings: list[str],
    selection_mode: str,
) -> dict[str, Any] | None:
    if not SHAPELY_AVAILABLE:
        return None
    normalized_purchase = _normalize_sheet_profile(purchase_profile or {}, 0) if purchase_profile else None
    normalized_stock = [
        profile
        for index, row in enumerate(list(stock_candidates or []))
        if (profile := _normalize_stock_sheet_candidate(row, index)) is not None
    ]
    selected_profile = normalized_purchase or {
        "name": "Apenas stock",
        "width_mm": 0.0,
        "height_mm": 0.0,
        "area_mm2": 0.0,
        "source_kind": "stock",
        "source_label": "Apenas stock",
    }
    sheets = [_new_shapely_sheet(profile, edge_margin_mm) for profile in _expand_stock_units(normalized_stock)]
    ordered_rows = sorted(list(expanded or []), key=_strategy_sort_key("area"), reverse=True)
    unplaced: list[dict[str, Any]] = []
    base_cache: dict[str, Any] = {}
    variant_cache: dict[tuple, list[dict[str, Any]]] = {}
    strategy_name = "shape-geos-stock"
    for row in ordered_rows:
        item: NestItem = row["item"]
        best_candidate: dict[str, Any] | None = None
        target_sheet: dict[str, Any] | None = None
        for sheet_index, sheet in enumerate(sheets):
            candidate = _try_place_on_shapely_sheet(
                sheet,
                item,
                allow_rotate=allow_rotate,
                allow_mirror=allow_mirror,
                free_angles=free_angles,
                part_spacing_mm=part_spacing_mm,
                edge_margin_mm=edge_margin_mm,
                strategy_name=strategy_name,
                base_cache=base_cache,
                variant_cache=variant_cache,
            )
            if candidate is None:
                continue
            source_kind = str(dict(sheet.get("profile", {}) or {}).get("source_kind", "purchase") or "purchase").lower()
            source_rank = 0.0 if source_kind in {"stock", "retalho"} else 1.0
            score = (source_rank, float(sheet_index), *tuple(candidate.get("score", ())))
            if best_candidate is None or score < best_candidate["global_score"]:
                best_candidate = {**candidate, "global_score": score}
                target_sheet = sheet
        if best_candidate is None and normalized_purchase is not None and allow_purchase_fallback:
            candidate_sheet = _new_shapely_sheet(normalized_purchase, edge_margin_mm)
            candidate = _try_place_on_shapely_sheet(
                candidate_sheet,
                item,
                allow_rotate=allow_rotate,
                allow_mirror=allow_mirror,
                free_angles=free_angles,
                part_spacing_mm=part_spacing_mm,
                edge_margin_mm=edge_margin_mm,
                strategy_name=strategy_name,
                base_cache=base_cache,
                variant_cache=variant_cache,
            )
            if candidate is not None:
                target_sheet = candidate_sheet
                sheets.append(candidate_sheet)
                best_candidate = candidate
        if best_candidate is None:
            unplaced.append(
                {
                    "ref_externa": item.ref_externa,
                    "description": item.description,
                    "file_name": item.file_name,
                    "copy_index": int(row.get("copy_index", 0) or 0),
                }
            )
            continue
        actual_geometry = best_candidate.pop("_actual_geometry")
        best_candidate.pop("score", None)
        best_candidate.pop("global_score", None)
        _apply_placement(
            target_sheet,
            best_candidate,
            row,
            edge_margin_mm=edge_margin_mm,
            part_spacing_mm=part_spacing_mm,
        )
        target_sheet.setdefault("_shapely_actual", []).append(actual_geometry)
    warnings = list(base_warnings or [])
    warnings.append("Stock e retalhos validados pelo motor poligonal GEOS antes de qualquer complemento de compra.")
    return _finalize_result(
        items,
        expanded,
        sheets,
        unplaced,
        settings=settings,
        warnings=warnings,
        selected_profile=selected_profile,
        selection_mode=selection_mode,
        strategy_name=strategy_name,
        shape_grid_mm=0.0,
    )


def _pack_with_stock_shape(
    items: list[NestItem],
    expanded: list[dict[str, Any]],
    *,
    stock_candidates: list[dict[str, Any]],
    purchase_profile: dict[str, Any] | None,
    part_spacing_mm: float,
    edge_margin_mm: float,
    allow_rotate: bool,
    allow_mirror: bool,
    free_angles: bool,
    allow_purchase_fallback: bool,
    settings: dict[str, Any],
    base_warnings: list[str],
    selection_mode: str,
    grid_mm: float,
) -> dict[str, Any]:
    geos_result = _pack_with_stock_shapely(
        items,
        expanded,
        stock_candidates=stock_candidates,
        purchase_profile=purchase_profile,
        part_spacing_mm=part_spacing_mm,
        edge_margin_mm=edge_margin_mm,
        allow_rotate=allow_rotate,
        allow_mirror=allow_mirror,
        free_angles=free_angles,
        allow_purchase_fallback=allow_purchase_fallback,
        settings=settings,
        base_warnings=base_warnings,
        selection_mode=selection_mode,
    )
    if geos_result is not None:
        return geos_result
    normalized_purchase = _normalize_sheet_profile(purchase_profile or {}, 0) if purchase_profile else None
    normalized_stock = [profile for index, row in enumerate(list(stock_candidates or [])) if (profile := _normalize_stock_sheet_candidate(row, index)) is not None]
    selected_profile = normalized_purchase or {
        "name": "Apenas stock",
        "width_mm": 0.0,
        "height_mm": 0.0,
        "area_mm2": 0.0,
        "source_kind": "stock",
        "source_label": "Apenas stock",
    }
    best_result: dict[str, Any] | None = None
    shape_cache: dict[tuple, dict[str, Any]] = {}
    for strategy_name, ordered_rows in _strategy_order_variants(
        expanded,
        shape_mode=True,
        free_angles=free_angles,
    ):
        if best_result is not None and _optimization_deadline_reached():
            break
        warnings = list(base_warnings or [])
        sheets: list[dict[str, Any]] = []
        unplaced: list[dict[str, Any]] = []
        available_stock_units = _expand_stock_units(normalized_stock)

        for row in ordered_rows:
            item: NestItem = row["item"]
            best_candidate: dict[str, Any] | None = None
            target_sheet: dict[str, Any] | None = None

            for sheet_index, sheet in enumerate(sheets):
                candidate = _try_place_on_shape_sheet(
                    sheet,
                    item,
                    allow_rotate=allow_rotate,
                    allow_mirror=allow_mirror,
                    free_angles=free_angles,
                    grid_mm=grid_mm,
                    part_spacing_mm=part_spacing_mm,
                    edge_margin_mm=edge_margin_mm,
                    cache=shape_cache,
                    strategy_name=strategy_name,
                )
                if candidate is None:
                    continue
                candidate_score = (
                    0.0,
                    float(sheet_index),
                    *_shape_candidate_score(candidate, strategy_name=strategy_name, sheet=sheet, edge_margin_mm=edge_margin_mm),
                )
                if best_candidate is None or candidate_score < best_candidate["score"]:
                    best_candidate = {"placement": candidate, "score": candidate_score}
                    target_sheet = sheet

            stock_choice: dict[str, Any] | None = None
            if best_candidate is None:
                for stock_index, stock_profile in enumerate(available_stock_units):
                    try:
                        candidate_sheet = _new_shape_sheet(stock_profile, edge_margin_mm=edge_margin_mm, grid_mm=grid_mm)
                    except Exception:
                        continue
                    candidate = _try_place_on_shape_sheet(
                        candidate_sheet,
                        item,
                        allow_rotate=allow_rotate,
                        allow_mirror=allow_mirror,
                        free_angles=free_angles,
                        grid_mm=grid_mm,
                        part_spacing_mm=part_spacing_mm,
                        edge_margin_mm=edge_margin_mm,
                        cache=shape_cache,
                        strategy_name=strategy_name,
                    )
                    if candidate is None:
                        continue
                    source_priority = 0.0 if str(stock_profile.get("source_kind", "") or "").strip().lower() == "retalho" else 1.0
                    candidate_score = (
                        1.0,
                        source_priority,
                        _as_float(stock_profile.get("area_mm2", 0.0), 0.0),
                        *_shape_candidate_score(candidate, strategy_name=strategy_name, sheet=candidate_sheet, edge_margin_mm=edge_margin_mm),
                        float(stock_index),
                    )
                    if best_candidate is None or candidate_score < best_candidate["score"]:
                        best_candidate = {"placement": candidate, "score": candidate_score}
                        stock_choice = {"index": stock_index, "profile": stock_profile}

            if best_candidate is None and normalized_purchase is not None and allow_purchase_fallback:
                try:
                    candidate_sheet = _new_shape_sheet(normalized_purchase, edge_margin_mm=edge_margin_mm, grid_mm=grid_mm)
                except Exception:
                    candidate_sheet = None
                if candidate_sheet is not None:
                    candidate = _try_place_on_shape_sheet(
                        candidate_sheet,
                        item,
                        allow_rotate=allow_rotate,
                        allow_mirror=allow_mirror,
                        free_angles=free_angles,
                        grid_mm=grid_mm,
                        part_spacing_mm=part_spacing_mm,
                        edge_margin_mm=edge_margin_mm,
                        cache=shape_cache,
                        strategy_name=strategy_name,
                    )
                    if candidate is not None:
                        best_candidate = {
                            "placement": candidate,
                            "score": (
                                2.0,
                                _as_float(normalized_purchase.get("area_mm2", 0.0), 0.0),
                                *_shape_candidate_score(candidate, strategy_name=strategy_name, sheet=candidate_sheet, edge_margin_mm=edge_margin_mm),
                            ),
                        }
                        target_sheet = candidate_sheet
                        sheets.append(candidate_sheet)

            if target_sheet is None and stock_choice is not None:
                target_sheet = _new_shape_sheet(stock_choice["profile"], edge_margin_mm=edge_margin_mm, grid_mm=grid_mm)
                sheets.append(target_sheet)
                available_stock_units.pop(stock_choice["index"])

            if target_sheet is None or best_candidate is None:
                warnings.append(f"{item.file_name}: nao foi possivel posicionar por contorno com o stock/formato atual.")
                unplaced.append(
                    {
                        "ref_externa": item.ref_externa,
                        "description": item.description,
                        "file_name": item.file_name,
                        "copy_index": int(row.get("copy_index", 0) or 0),
                    }
                )
                continue

            _apply_placement(target_sheet, best_candidate["placement"], row, edge_margin_mm=edge_margin_mm, part_spacing_mm=part_spacing_mm)

        strategy_result = _finalize_result(
            items,
            expanded,
            sheets,
            unplaced,
            settings=settings,
            warnings=warnings,
            selected_profile=selected_profile,
            selection_mode=selection_mode,
            strategy_name=strategy_name,
        )
        if best_result is None or _result_score(strategy_result) < _result_score(best_result):
            best_result = strategy_result

    return best_result or _finalize_result(
        items,
        expanded,
        [],
        [],
        settings=settings,
        warnings=list(base_warnings or []),
        selected_profile=selected_profile,
        selection_mode=selection_mode,
        strategy_name="shape",
        shape_grid_mm=grid_mm,
    )


def _candidate_row_from_result(name: str, result: dict[str, Any]) -> dict[str, Any]:
    summary = dict(result.get("summary", {}) or {})
    return {
        "name": str(name or "").strip(),
        "method": _engine_method_label(summary),
        "shape_aware": bool(summary.get("shape_aware", False)),
        "engine_used": _engine_mode_from_summary(summary),
        "sheet_count": int(summary.get("sheet_count", 0) or 0),
        "stock_sheet_count": int(summary.get("stock_sheet_count", 0) or 0),
        "purchased_sheet_count": int(summary.get("purchased_sheet_count", 0) or 0),
        "part_count_unplaced": int(summary.get("part_count_unplaced", 0) or 0),
        "purchase_sheet_area_mm2": round(_as_float(summary.get("purchase_sheet_area_mm2", 0.0), 0.0), 2),
        "total_sheet_area_mm2": round(_as_float(summary.get("total_sheet_area_mm2", 0.0), 0.0), 2),
        "utilization_net_pct": round(_as_float(summary.get("utilization_net_pct", 0.0), 0.0), 2),
        "utilization_bbox_pct": round(_as_float(summary.get("utilization_bbox_pct", 0.0), 0.0), 2),
        "layout_compactness_pct": round(_as_float(summary.get("layout_compactness_pct", 0.0), 0.0), 2),
    }


def nest_parts(
    rows: list[dict[str, Any]],
    *,
    sheet_width_mm: float | None = None,
    sheet_height_mm: float | None = None,
    part_spacing_mm: float,
    edge_margin_mm: float,
    allow_rotate: bool,
    laser_settings: dict[str, Any] | None = None,
    sheet_name: str = "",
    sheet_profiles: list[dict[str, Any]] | None = None,
    auto_select_sheet: bool = False,
    stock_sheet_candidates: list[dict[str, Any]] | None = None,
    use_stock_first: bool = False,
    allow_purchase_fallback: bool = True,
    shape_aware: bool | None = None,
    allow_mirror: bool | None = None,
    free_angle_rotation: bool | None = None,
    strict_shape_only: bool = False,
    shape_grid_mm: float | None = None,
    optimization_level: str | None = None,
    time_limit_s: float | None = None,
    cancel_check: Any = None,
) -> dict[str, Any]:
    settings = merge_laser_quote_settings(laser_settings)
    nesting_options = default_nesting_options(settings)
    requested_level = str(optimization_level or nesting_options.get("optimization_level", "tap1") or "tap1").strip().lower()
    level = requested_level if requested_level in {"tap1", "tap2", "tap3"} else "tap1"
    default_limits = {"tap1": 30.0, "tap2": 120.0, "tap3": 240.0}
    limit_s = max(0.0, _as_float(time_limit_s, default_limits[level])) if time_limit_s is not None else default_limits[level]
    started_at = time.monotonic()
    _NESTING_RUN_CONTEXT.level = level
    _NESTING_RUN_CONTEXT.time_limit_s = limit_s
    _NESTING_RUN_CONTEXT.started_at = started_at
    _NESTING_RUN_CONTEXT.deadline = started_at + limit_s if limit_s > 0.0 else 0.0
    _NESTING_RUN_CONTEXT.cancel_check = cancel_check
    items, warnings = build_nesting_items(rows, settings)
    expanded = _expand_items(items)
    normalized_stock = [profile for index, row in enumerate(list(stock_sheet_candidates or [])) if (profile := _normalize_stock_sheet_candidate(row, index)) is not None]
    use_shape_engine = bool(nesting_options.get("shape_aware", True) if shape_aware is None else shape_aware)
    allow_shape_mirror = bool(nesting_options.get("allow_mirror", True) if allow_mirror is None else allow_mirror)
    allow_free_angles, free_angle_note = _effective_free_angle_rotation(
        items,
        expanded,
        bool(nesting_options.get("free_angle_rotation", False) if free_angle_rotation is None else free_angle_rotation),
    )
    if free_angle_note:
        warnings.append(free_angle_note)
    requested_grid_mm = max(2.0, _as_float(nesting_options.get("shape_grid_mm", 10.0) if shape_grid_mm is None else shape_grid_mm, 10.0), 2.0)
    grid_mm = _effective_shape_grid_mm(requested_grid_mm, part_spacing_mm, edge_margin_mm)
    requested_engine = "shape" if use_shape_engine else "bbox"

    def _choose_engine_variant(*, bbox_result: dict[str, Any] | None, shape_result: dict[str, Any] | None) -> dict[str, Any]:
        if requested_engine == "shape":
            if shape_result is not None:
                if bbox_result is not None and _result_score(bbox_result) < _result_score(shape_result):
                    chosen = dict(bbox_result or {})
                    chosen_summary = dict(chosen.get("summary", {}) or {})
                    chosen_summary["engine_requested"] = "shape"
                    # A non-overlapping bounding-box plan is also geometrically
                    # valid for the enclosed real contours; retain shape mode
                    # while recording which conservative optimizer won.
                    chosen_summary["engine_used"] = "shape"
                    chosen_summary["shape_optimizer"] = "bbox-conservative"
                    chosen_summary["engine_modes_tested"] = ["shape", "bbox"]
                    chosen["summary"] = chosen_summary
                    chosen["warnings"] = _unique_texts(
                        [
                            "O otimizador híbrido escolheu o plano conservador por caixas porque usa menos chapa ou deixa um remanescente mais útil."
                        ]
                        + list(chosen.get("warnings", []) or [])
                    )
                    return chosen
                chosen = dict(shape_result or {})
                chosen_summary = dict(chosen.get("summary", {}) or {})
                chosen_summary["engine_requested"] = "shape"
                chosen_summary["engine_used"] = "shape"
                chosen_summary["engine_modes_tested"] = ["shape"] + (["bbox"] if bbox_result is not None else [])
                chosen["summary"] = chosen_summary
                return chosen
            if strict_shape_only:
                raise RuntimeError("Nao foi possivel calcular o nesting apenas por contorno real com a geometria e configuracao atuais.")
            if bbox_result is not None:
                fallback = dict(bbox_result or {})
                fallback_summary = dict(fallback.get("summary", {}) or {})
                fallback_summary["engine_requested"] = "shape"
                fallback_summary["engine_used"] = "bbox"
                fallback_summary["engine_modes_tested"] = ["bbox"]
                fallback["summary"] = fallback_summary
                fallback["warnings"] = _unique_texts(
                    ["Nesting por contorno indisponível neste cenário; foi usado fallback por bounding box."]
                    + list(fallback.get("warnings", []) or [])
                )
                return fallback
        return _choose_best_engine_result(
            bbox_result=bbox_result,
            shape_result=shape_result,
            requested_mode=requested_engine,
            grid_mm=grid_mm,
        )

    if auto_select_sheet:
        profiles: list[dict[str, Any]] = []
        raw_profiles = list(sheet_profiles or default_sheet_profiles(settings))
        for index, row in enumerate(raw_profiles):
            profile = _normalize_sheet_profile(dict(row or {}), index)
            if profile is not None:
                profiles.append(profile)
        shape_ok, shape_reason = _shape_engine_feasible(
            items=items,
            profiles=profiles,
            stock_candidates=normalized_stock,
            edge_margin_mm=edge_margin_mm,
            grid_mm=grid_mm,
        )
        shape_active = bool(use_shape_engine and shape_ok)
        if use_shape_engine and strict_shape_only and not shape_active:
            raise RuntimeError(shape_reason or "O motor de contorno real nao conseguiu validar este cenario.")
        if use_shape_engine and not shape_active and shape_reason:
            warnings.append(f"Nesting por contorno desativado automaticamente: {shape_reason}")
        best_result: dict[str, Any] | None = None
        candidate_rows: list[dict[str, Any]] = []
        candidate_errors: list[str] = []

        if use_stock_first and normalized_stock:
            stock_only_bbox: dict[str, Any] | None = None
            stock_only_shape: dict[str, Any] | None = None
            if shape_active and requested_engine == "shape":
                stock_only_shape = _pack_with_stock_shape(
                    items,
                    expanded,
                    stock_candidates=normalized_stock,
                    purchase_profile=None,
                    part_spacing_mm=part_spacing_mm,
                    edge_margin_mm=edge_margin_mm,
                    allow_rotate=allow_rotate,
                    allow_mirror=allow_shape_mirror,
                    free_angles=allow_free_angles,
                    allow_purchase_fallback=False,
                    settings=settings,
                    base_warnings=warnings,
                    selection_mode="auto_stock",
                    grid_mm=grid_mm,
                )
                if stock_only_shape is None and not strict_shape_only:
                    stock_only_bbox = _pack_with_stock(
                        items,
                        expanded,
                        stock_candidates=normalized_stock,
                        purchase_profile=None,
                        part_spacing_mm=part_spacing_mm,
                        edge_margin_mm=edge_margin_mm,
                        allow_rotate=allow_rotate,
                        allow_purchase_fallback=False,
                        settings=settings,
                        base_warnings=warnings,
                        selection_mode="auto_stock",
                    )
            else:
                stock_only_bbox = _pack_with_stock(
                    items,
                    expanded,
                    stock_candidates=normalized_stock,
                    purchase_profile=None,
                    part_spacing_mm=part_spacing_mm,
                    edge_margin_mm=edge_margin_mm,
                    allow_rotate=allow_rotate,
                    allow_purchase_fallback=False,
                    settings=settings,
                    base_warnings=warnings,
                    selection_mode="auto_stock",
                )
            stock_only_result = _choose_engine_variant(bbox_result=stock_only_bbox, shape_result=stock_only_shape)
            candidate_rows.append(_candidate_row_from_result("Apenas stock", stock_only_result))
            best_result = stock_only_result

        if not profiles and not (use_stock_first and normalized_stock and not allow_purchase_fallback):
            raise ValueError("Define pelo menos um formato de chapa valido para a escolha automatica.")

        for profile in profiles:
            try:
                if use_stock_first and normalized_stock:
                    bbox_result: dict[str, Any] | None = None
                    shape_result: dict[str, Any] | None = None
                    if shape_active and requested_engine == "shape":
                        shape_result = _pack_with_stock_shape(
                            items,
                            expanded,
                            stock_candidates=normalized_stock,
                            purchase_profile=profile,
                            part_spacing_mm=part_spacing_mm,
                            edge_margin_mm=edge_margin_mm,
                            allow_rotate=allow_rotate,
                            allow_mirror=allow_shape_mirror,
                            free_angles=allow_free_angles,
                            allow_purchase_fallback=allow_purchase_fallback,
                            settings=settings,
                            base_warnings=warnings,
                            selection_mode="auto_stock",
                            grid_mm=grid_mm,
                        )
                        if shape_result is None and not strict_shape_only:
                            bbox_result = _pack_with_stock(
                                items,
                                expanded,
                                stock_candidates=normalized_stock,
                                purchase_profile=profile,
                                part_spacing_mm=part_spacing_mm,
                                edge_margin_mm=edge_margin_mm,
                                allow_rotate=allow_rotate,
                                allow_purchase_fallback=allow_purchase_fallback,
                                settings=settings,
                                base_warnings=warnings,
                                selection_mode="auto_stock",
                            )
                    else:
                        bbox_result = _pack_with_stock(
                            items,
                            expanded,
                            stock_candidates=normalized_stock,
                            purchase_profile=profile,
                            part_spacing_mm=part_spacing_mm,
                            edge_margin_mm=edge_margin_mm,
                            allow_rotate=allow_rotate,
                            allow_purchase_fallback=allow_purchase_fallback,
                            settings=settings,
                            base_warnings=warnings,
                            selection_mode="auto_stock",
                        )
                    result = _choose_engine_variant(bbox_result=bbox_result, shape_result=shape_result)
                else:
                    bbox_result = None
                    shape_result = None
                    profile_supports_all = _profile_can_fit_all_items(
                        profile,
                        items,
                        edge_margin_mm=edge_margin_mm,
                        allow_rotate=allow_rotate,
                    )
                    if shape_active and requested_engine == "shape" and (profile_supports_all or strict_shape_only or len(expanded) <= 2):
                        shape_result = _pack_profile_shape(
                            items,
                            expanded,
                            profile=profile,
                            part_spacing_mm=part_spacing_mm,
                            edge_margin_mm=edge_margin_mm,
                            allow_rotate=allow_rotate,
                            allow_mirror=allow_shape_mirror,
                            free_angles=allow_free_angles,
                            settings=settings,
                            base_warnings=warnings,
                            selection_mode="auto",
                            grid_mm=grid_mm,
                        )
                        if not strict_shape_only and (shape_result is None or len(expanded) > 16):
                            bbox_result = _pack_profile(
                                items,
                                expanded,
                                profile=profile,
                                part_spacing_mm=part_spacing_mm,
                                edge_margin_mm=edge_margin_mm,
                                allow_rotate=allow_rotate,
                                settings=settings,
                                base_warnings=warnings,
                                selection_mode="auto",
                            )
                    else:
                        bbox_result = _pack_profile(
                            items,
                            expanded,
                            profile=profile,
                            part_spacing_mm=part_spacing_mm,
                            edge_margin_mm=edge_margin_mm,
                            allow_rotate=allow_rotate,
                            settings=settings,
                            base_warnings=warnings,
                            selection_mode="auto",
                        )
                    result = _choose_engine_variant(bbox_result=bbox_result, shape_result=shape_result)
            except Exception as exc:
                candidate_errors.append(f"{profile.get('name', 'Chapa')}: {exc}")
                continue
            candidate_rows.append(_candidate_row_from_result(str(profile.get("name", "") or "").strip(), result))
            if best_result is None or _result_score(result) < _result_score(best_result):
                best_result = result

        if best_result is None:
            raise ValueError("Nao foi possivel analisar os formatos de chapa disponiveis.")
        best_result["sheet_candidates"] = candidate_rows
        if candidate_errors:
            best_result["warnings"] = _unique_texts(list(best_result.get("warnings", []) or []) + candidate_errors)
        return best_result

    profile = _normalize_sheet_profile(
        {
            "name": sheet_name or f"{_as_float(sheet_width_mm, 0.0):g} x {_as_float(sheet_height_mm, 0.0):g}",
            "width_mm": sheet_width_mm,
            "height_mm": sheet_height_mm,
        },
        0,
    )
    shape_ok, shape_reason = _shape_engine_feasible(
        items=items,
        profiles=[profile] if profile else [],
        stock_candidates=normalized_stock,
        edge_margin_mm=edge_margin_mm,
        grid_mm=grid_mm,
    )
    shape_active = bool(use_shape_engine and shape_ok)
    if use_shape_engine and strict_shape_only and not shape_active:
        raise RuntimeError(shape_reason or "O motor de contorno real nao conseguiu validar este cenario.")
    if use_shape_engine and not shape_active and shape_reason:
        warnings.append(f"Nesting por contorno desativado automaticamente: {shape_reason}")
    if use_stock_first and normalized_stock:
        bbox_result: dict[str, Any] | None = None
        shape_result: dict[str, Any] | None = None
        if shape_active and requested_engine == "shape":
            shape_result = _pack_with_stock_shape(
                items,
                expanded,
                stock_candidates=normalized_stock,
                purchase_profile=profile,
                part_spacing_mm=part_spacing_mm,
                edge_margin_mm=edge_margin_mm,
                allow_rotate=allow_rotate,
                allow_mirror=allow_shape_mirror,
                free_angles=allow_free_angles,
                allow_purchase_fallback=allow_purchase_fallback,
                settings=settings,
                base_warnings=warnings,
                selection_mode="manual_stock",
                grid_mm=grid_mm,
            )
            if shape_result is None and not strict_shape_only:
                bbox_result = _pack_with_stock(
                    items,
                    expanded,
                    stock_candidates=normalized_stock,
                    purchase_profile=profile,
                    part_spacing_mm=part_spacing_mm,
                    edge_margin_mm=edge_margin_mm,
                    allow_rotate=allow_rotate,
                    allow_purchase_fallback=allow_purchase_fallback,
                    settings=settings,
                    base_warnings=warnings,
                    selection_mode="manual_stock",
                )
        else:
            bbox_result = _pack_with_stock(
                items,
                expanded,
                stock_candidates=normalized_stock,
                purchase_profile=profile,
                part_spacing_mm=part_spacing_mm,
                edge_margin_mm=edge_margin_mm,
                allow_rotate=allow_rotate,
                allow_purchase_fallback=allow_purchase_fallback,
                settings=settings,
                base_warnings=warnings,
                selection_mode="manual_stock",
            )
        return _choose_engine_variant(bbox_result=bbox_result, shape_result=shape_result)
    if profile is None:
        raise ValueError("Seleciona um formato de chapa valido.")
    bbox_result = None
    shape_result = None
    if shape_active and requested_engine == "shape":
        shape_result = _pack_profile_shape(
            items,
            expanded,
            profile=profile,
            part_spacing_mm=part_spacing_mm,
            edge_margin_mm=edge_margin_mm,
            allow_rotate=allow_rotate,
            allow_mirror=allow_shape_mirror,
            free_angles=allow_free_angles,
            settings=settings,
            base_warnings=warnings,
            selection_mode="manual",
            grid_mm=grid_mm,
        )
        if not strict_shape_only and (shape_result is None or len(expanded) > 16):
            bbox_result = _pack_profile(
                items,
                expanded,
                profile=profile,
                part_spacing_mm=part_spacing_mm,
                edge_margin_mm=edge_margin_mm,
                allow_rotate=allow_rotate,
                settings=settings,
                base_warnings=warnings,
                selection_mode="manual",
            )
    else:
        bbox_result = _pack_profile(
            items,
            expanded,
            profile=profile,
            part_spacing_mm=part_spacing_mm,
            edge_margin_mm=edge_margin_mm,
            allow_rotate=allow_rotate,
            settings=settings,
            base_warnings=warnings,
            selection_mode="manual",
        )
    return _choose_engine_variant(bbox_result=bbox_result, shape_result=shape_result)
