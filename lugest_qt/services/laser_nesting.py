from __future__ import annotations

import math
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .laser_quote_engine import analyze_dxf_geometry, merge_laser_quote_settings


DEFAULT_SHEET_PROFILES: list[dict[str, Any]] = [
    {"name": "1000 x 2000", "width_mm": 1000.0, "height_mm": 2000.0},
    {"name": "1250 x 2500", "width_mm": 1250.0, "height_mm": 2500.0},
    {"name": "1500 x 3000", "width_mm": 1500.0, "height_mm": 3000.0},
    {"name": "2000 x 4000", "width_mm": 2000.0, "height_mm": 4000.0},
]


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


def _sheet_overlap_diagnostics(sheet_row: dict[str, Any]) -> dict[str, Any]:
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
        "shape_grid_mm": max(2.0, _as_float(payload.get("shape_grid_mm", 10.0), 10.0)),
        "common_line_estimate": bool(payload.get("common_line_estimate", True)),
        "common_line_tolerance_mm": max(0.0, _as_float(payload.get("common_line_tolerance_mm", 1.0), 1.0)),
        "lead_optimization": bool(payload.get("lead_optimization", True)),
        "lead_optimization_pct": max(0.0, min(50.0, _as_float(payload.get("lead_optimization_pct", 8.0), 8.0))),
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


def _candidate_orientations(item: NestItem, allow_rotate: bool) -> list[dict[str, Any]]:
    policy = str(getattr(item, "rotation_policy", "auto") or "auto").strip().lower()
    has_rotation = bool(abs(item.bbox_width_mm - item.bbox_height_mm) > 1e-6)
    can_rotate = bool(allow_rotate and has_rotation)
    if policy == "90":
        if has_rotation:
            return [{"width": item.bbox_height_mm, "height": item.bbox_width_mm, "rotated": True}]
        return [{"width": item.bbox_width_mm, "height": item.bbox_height_mm, "rotated": False}]
    if policy == "0" or not has_rotation:
        return [{"width": item.bbox_width_mm, "height": item.bbox_height_mm, "rotated": False}]
    return [
        {"width": item.bbox_width_mm, "height": item.bbox_height_mm, "rotated": False},
        {"width": item.bbox_height_mm, "height": item.bbox_width_mm, "rotated": True},
    ]


def _item_shape_polygons(item: NestItem, rotated: bool) -> tuple[tuple[tuple[tuple[float, float], ...], ...], tuple[tuple[tuple[float, float], ...], ...], float, float]:
    outer = tuple(tuple((float(x), float(y)) for x, y in list(polygon or [])) for polygon in list(item.outer_polygons or []))
    holes = tuple(tuple((float(x), float(y)) for x, y in list(polygon or [])) for polygon in list(item.hole_polygons or []))
    if not rotated:
        return outer, holes, item.bbox_width_mm, item.bbox_height_mm

    transformed_outer: list[list[tuple[float, float]]] = []
    transformed_holes: list[list[tuple[float, float]]] = []
    all_points: list[tuple[float, float]] = []

    def _rotate(points: tuple[tuple[float, float], ...]) -> list[tuple[float, float]]:
        return [(round(item.bbox_height_mm - y, 3), round(x, 3)) for x, y in list(points or [])]

    for polygon in outer:
        rotated_points = _rotate(polygon)
        transformed_outer.append(rotated_points)
        all_points.extend(rotated_points)
    for polygon in holes:
        rotated_points = _rotate(polygon)
        transformed_holes.append(rotated_points)
        all_points.extend(rotated_points)
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


def _item_preview_paths(item: NestItem, rotated: bool) -> tuple[tuple[tuple[float, float], ...], ...]:
    paths = tuple(tuple((float(x), float(y)) for x, y in list(path or [])) for path in list(item.preview_paths or []))
    if not paths:
        return ()
    if not rotated:
        return tuple(path for path in paths if len(path) >= 2)

    transformed_paths: list[list[tuple[float, float]]] = []
    all_points: list[tuple[float, float]] = []
    for path in list(paths or []):
        rotated_path = [(round(item.bbox_height_mm - y, 3), round(x, 3)) for x, y in list(path or [])]
        if len(rotated_path) < 2:
            continue
        transformed_paths.append(rotated_path)
        all_points.extend(rotated_path)
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
) -> bool:
    sample_points = [
        (cell_x_mm + (grid_mm * 0.50), cell_y_mm + (grid_mm * 0.50)),
        (cell_x_mm + (grid_mm * 0.20), cell_y_mm + (grid_mm * 0.20)),
        (cell_x_mm + (grid_mm * 0.80), cell_y_mm + (grid_mm * 0.20)),
        (cell_x_mm + (grid_mm * 0.80), cell_y_mm + (grid_mm * 0.80)),
        (cell_x_mm + (grid_mm * 0.20), cell_y_mm + (grid_mm * 0.80)),
    ]
    for point in sample_points:
        if any(_point_in_polygon(point, polygon) for polygon in list(outer_polygons or [])) and not any(
            _point_in_polygon(point, polygon) for polygon in list(hole_polygons or [])
        ):
            return True
    for polygon in list(outer_polygons or []):
        for px, py in list(polygon or []):
            if cell_x_mm <= px <= (cell_x_mm + grid_mm) and cell_y_mm <= py <= (cell_y_mm + grid_mm):
                if not any(_point_in_polygon((px, py), hole) for hole in list(hole_polygons or [])):
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
    grid_mm: float,
    part_spacing_mm: float,
    cache: dict[tuple[str, bool, float, float], dict[str, Any]],
) -> dict[str, Any]:
    cache_key = str(getattr(item, "shape_cache_key", "") or item.path)
    key = (cache_key, bool(rotated), round(grid_mm, 4), round(part_spacing_mm, 4))
    cached = cache.get(key)
    if cached is not None:
        return cached

    outer_polygons, hole_polygons, width_mm, height_mm = _item_shape_polygons(item, rotated)
    if not outer_polygons:
        fallback_polygon = _rectangle_polygon(width_mm, height_mm)
        outer_polygons = (fallback_polygon,) if fallback_polygon else ()
    cols = max(1, int(math.ceil(max(width_mm, 0.0) / max(grid_mm, 1.0))))
    rows = max(1, int(math.ceil(max(height_mm, 0.0) / max(grid_mm, 1.0))))
    base_cells: set[tuple[int, int]] = set()
    for row_index in range(rows):
        cell_y_mm = row_index * grid_mm
        for col_index in range(cols):
            cell_x_mm = col_index * grid_mm
            if _cell_hits_shape(cell_x_mm, cell_y_mm, grid_mm, outer_polygons, hole_polygons):
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


def _shape_candidate_score(candidate: dict[str, Any]) -> tuple[float, ...]:
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
    allowed_indices: set[int] = set()
    for row_index in range(height_cells):
        cell_y_mm = edge_margin_mm + (row_index * grid_mm)
        for col_index in range(width_cells):
            cell_x_mm = edge_margin_mm + (col_index * grid_mm)
            if _cell_hits_shape(cell_x_mm, cell_y_mm, grid_mm, outer_polygons, hole_polygons):
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
    grid_mm: float,
    part_spacing_mm: float,
    edge_margin_mm: float,
    cache: dict[tuple[str, bool, float, float], dict[str, Any]],
) -> dict[str, Any] | None:
    occupied = set(sheet.get("occupied_cells", set()) or set())
    allowed = sheet.get("allowed_cells", None)
    sheet_width_cells = max(1, int(sheet.get("grid_width_cells", 1) or 1))
    sheet_height_cells = max(1, int(sheet.get("grid_height_cells", 1) or 1))
    best: dict[str, Any] | None = None

    for orientation in _candidate_orientations(item, allow_rotate):
        mask = _shape_mask(
            item,
            rotated=bool(orientation.get("rotated")),
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
        placed_for_orientation = False
        for grid_y in range(sheet_height_cells - mask_height_cells + 1):
            if placed_for_orientation:
                break
            for grid_x in range(sheet_width_cells - mask_width_cells + 1):
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
                    continue
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
                    continue
                candidate = {
                    "x": round(grid_x * grid_mm, 3),
                    "y": round(grid_y * grid_mm, 3),
                    "grid_x": grid_x,
                    "grid_y": grid_y,
                    "width": round(_as_float(mask.get("bbox_width_mm", orientation.get("width", 0.0)), 0.0), 3),
                    "height": round(_as_float(mask.get("bbox_height_mm", orientation.get("height", 0.0)), 0.0), 3),
                    "rotated": bool(orientation.get("rotated")),
                    "place_w": round(mask_width_cells * grid_mm, 3),
                    "place_h": round(mask_height_cells * grid_mm, 3),
                    "mask_cells": mask_cells,
                    "draw_offset_x_mm": round(_as_float(mask.get("draw_offset_x_mm", 0.0), 0.0), 3),
                    "draw_offset_y_mm": round(_as_float(mask.get("draw_offset_y_mm", 0.0), 0.0), 3),
                    "occupied_area_mm2": round(_as_float(mask.get("occupied_area_mm2", 0.0), 0.0), 2),
                    "shape_outer_polygons": tuple(mask.get("shape_outer_polygons", ()) or ()),
                    "shape_hole_polygons": tuple(mask.get("shape_hole_polygons", ()) or ()),
                }
                if best is None or _shape_candidate_score(candidate) < _shape_candidate_score(best):
                    best = candidate
                placed_for_orientation = True
                break
    return best


def _placement_score(candidate: dict[str, Any]) -> tuple[float, ...]:
    return (
        0.0 if not bool(candidate.get("new_shelf")) else 1.0,
        _as_float(candidate.get("y", 0.0), 0.0) + _as_float(candidate.get("place_h", 0.0), 0.0),
        _as_float(candidate.get("waste", 0.0), 0.0),
        _as_float(candidate.get("height_gap", 0.0), 0.0),
        _as_float(candidate.get("place_h", 0.0), 0.0),
    )


def _profile_usable_dimensions(profile: dict[str, Any], edge_margin_mm: float) -> tuple[float, float]:
    width_mm = max(0.0, _as_float(profile.get("width_mm", 0.0), 0.0))
    height_mm = max(0.0, _as_float(profile.get("height_mm", 0.0), 0.0))
    usable_width = width_mm - (2.0 * max(0.0, edge_margin_mm))
    usable_height = height_mm - (2.0 * max(0.0, edge_margin_mm))
    if usable_width <= 0.0 or usable_height <= 0.0:
        raise ValueError("A margem a borda e maior do que a chapa util disponivel.")
    return usable_width, usable_height


def _try_place_on_sheet(
    sheet: dict[str, Any],
    item: NestItem,
    *,
    usable_width: float,
    usable_height: float,
    part_spacing_mm: float,
    allow_rotate: bool,
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
            if best is None or _placement_score(candidate) < _placement_score(best):
                best = candidate
        cursor_y = _as_float(sheet.get("cursor_y", 0.0), 0.0)
        if cursor_y + place_h <= usable_height + 1e-6:
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
            if best is None or _placement_score(candidate) < _placement_score(best):
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
            "x_mm": round(draw_x, 3),
            "y_mm": round(draw_y, 3),
            "width_mm": round(_as_float(placement.get("width", 0.0), 0.0), 3),
            "height_mm": round(_as_float(placement.get("height", 0.0), 0.0), 3),
            "net_area_mm2": round(item.net_area_mm2, 2),
            "bbox_area_mm2": round(item.bbox_width_mm * item.bbox_height_mm, 2),
            "layout_area_mm2": layout_area_mm2,
            "copy_index": int(row.get("copy_index", 0) or 0),
            "shape_mode": "grid" if "mask_cells" in placement else "bbox",
            "shape_outer_polygons": _translate_polygons(tuple(placement.get("shape_outer_polygons", ()) or ()), draw_x, draw_y),
            "shape_hole_polygons": _translate_polygons(tuple(placement.get("shape_hole_polygons", ()) or ()), draw_x, draw_y),
            "preview_paths": _translate_paths(_item_preview_paths(item, bool(placement.get("rotated"))), draw_x, draw_y),
        }
    )
    sheet["placements"] = placements
    sheet["used_net_area_mm2"] = _as_float(sheet.get("used_net_area_mm2", 0.0), 0.0) + item.net_area_mm2
    sheet["used_bbox_area_mm2"] = _as_float(sheet.get("used_bbox_area_mm2", 0.0), 0.0) + layout_area_mm2


def _strategy_sort_key(name: str):
    normalized = str(name or "").strip().lower()
    if normalized.startswith("shape-"):
        normalized = normalized[6:]
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


def _result_score(result: dict[str, Any]) -> tuple[float, ...]:
    summary = dict(result.get("summary", {}) or {})
    return (
        _as_int(summary.get("part_count_unplaced", 0), 0),
        _as_float(summary.get("purchase_sheet_area_mm2", 0.0), 0.0),
        _as_float(summary.get("total_sheet_area_mm2", 0.0), 0.0),
        _as_int(summary.get("sheet_count", 0), 0),
        _as_float(summary.get("layout_span_area_mm2", 0.0), 0.0),
        -_as_float(summary.get("layout_compactness_pct", 0.0), 0.0),
        -_as_float(summary.get("utilization_net_pct", 0.0), 0.0),
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
    layout_span_area = layout_span_width * layout_span_height
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
        "remaining_bbox_area_mm2": round(max(0.0, area_mm2 - used_bbox), 2),
        "geometry_validation": _sheet_overlap_diagnostics({"placements": placements}),
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
        if solid_overlap_pair_count > 0:
            pair_labels = ", ".join(f"{left}/{right}" for left, right in list(geometry_validation.get("solid_overlap_pairs", []) or [])[:6])
            warnings.append(
                f"Chapa {index + 1}: foram detetadas {solid_overlap_pair_count} colisoes geometricas reais entre pecas ({pair_labels})."
            )
        elif part_in_part_pair_count > 0:
            warnings.append(
                f"Chapa {index + 1}: foram detetados {part_in_part_pair_count} encaixes internos por contorno (part-in-part), sem sobreposicao real de geometria."
            )
        sheet_rows.append(row)
        summary["sheet_count"] += 1
        summary["part_count_placed"] += int(row.get("part_count", 0) or 0)
        summary["used_net_area_mm2"] += _as_float(row.get("used_net_area_mm2", 0.0), 0.0)
        summary["used_bbox_area_mm2"] += _as_float(row.get("used_bbox_area_mm2", 0.0), 0.0)
        summary["layout_span_area_mm2"] += _as_float(row.get("layout_span_area_mm2", 0.0), 0.0)
        summary["geometry_solid_overlap_pair_count"] = int(summary.get("geometry_solid_overlap_pair_count", 0) or 0) + solid_overlap_pair_count
        summary["geometry_part_in_part_pair_count"] = int(summary.get("geometry_part_in_part_pair_count", 0) or 0) + part_in_part_pair_count
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

    for strategy_name in ("longest-side", "area", "height-first", "width-first"):
        ordered_rows = sorted(list(expanded or []), key=_strategy_sort_key(strategy_name), reverse=True)
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
                candidate_score = (0.0, *_placement_score(candidate), float(sheet_index))
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
                best_candidate = {"placement": candidate, "score": (1.0, *_placement_score(candidate), float(len(sheets) - 1))}
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
    max_cells = 0
    for profile in candidates:
        try:
            usable_width, usable_height = _profile_usable_dimensions(profile, edge_margin_mm)
        except Exception:
            continue
        cells = int(math.ceil(usable_width / safe_grid) * math.ceil(usable_height / safe_grid))
        max_cells = max(max_cells, cells)
    if max_cells > 140_000:
        return False, f"Grelha de {safe_grid:g} mm demasiado fina para a chapa configurada."
    if sum(max(1, int(item.qty or 0)) for item in list(items or [])) > 400:
        return False, "Quantidade de pecas demasiado elevada para o modo por contorno nesta fase."
    return True, ""


def _pack_profile_shape(
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
    grid_mm: float,
) -> dict[str, Any]:
    normalized_profile = _normalize_sheet_profile(profile, 0)
    if normalized_profile is None:
        raise ValueError("Seleciona um formato de chapa valido.")
    best_result: dict[str, Any] | None = None
    shape_cache: dict[tuple[str, bool, float, float], dict[str, Any]] = {}

    for strategy_name in ("shape-longest-side", "shape-area", "shape-height-first", "shape-width-first"):
        ordered_rows = sorted(list(expanded or []), key=_strategy_sort_key(strategy_name), reverse=True)
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
                    grid_mm=grid_mm,
                    part_spacing_mm=part_spacing_mm,
                    edge_margin_mm=edge_margin_mm,
                    cache=shape_cache,
                )
                if candidate is None:
                    continue
                candidate_score = (0.0, *_shape_candidate_score(candidate), float(sheet_index))
                if best_candidate is None or candidate_score < best_candidate["score"]:
                    best_candidate = {"placement": candidate, "score": candidate_score}
                    target_sheet = sheet
            if best_candidate is None:
                candidate_sheet = _new_shape_sheet(normalized_profile, edge_margin_mm=edge_margin_mm, grid_mm=grid_mm)
                candidate = _try_place_on_shape_sheet(
                    candidate_sheet,
                    item,
                    allow_rotate=allow_rotate,
                    grid_mm=grid_mm,
                    part_spacing_mm=part_spacing_mm,
                    edge_margin_mm=edge_margin_mm,
                    cache=shape_cache,
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
                best_candidate = {"placement": candidate, "score": (1.0, *_shape_candidate_score(candidate), float(len(sheets) - 1))}
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

    for strategy_name in ("longest-side", "area", "height-first", "width-first"):
        ordered_rows = sorted(list(expanded or []), key=_strategy_sort_key(strategy_name), reverse=True)
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
                )
                if candidate is None:
                    continue
                candidate_score = (0.0, *_placement_score(candidate), float(sheet_index))
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
                    )
                    if candidate is None:
                        continue
                    source_priority = 0.0 if str(stock_profile.get("source_kind", "") or "").strip().lower() == "retalho" else 1.0
                    candidate_score = (1.0, source_priority, _as_float(stock_profile.get("area_mm2", 0.0), 0.0), *_placement_score(candidate), float(stock_index))
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
                )
                if candidate is not None:
                    best_candidate = {"placement": candidate, "score": (2.0, _as_float(normalized_purchase.get("area_mm2", 0.0), 0.0), *_placement_score(candidate))}
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


def _pack_with_stock_shape(
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
    grid_mm: float,
) -> dict[str, Any]:
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
    shape_cache: dict[tuple[str, bool, float, float], dict[str, Any]] = {}

    for strategy_name in ("shape-longest-side", "shape-area", "shape-height-first", "shape-width-first"):
        ordered_rows = sorted(list(expanded or []), key=_strategy_sort_key(strategy_name), reverse=True)
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
                    grid_mm=grid_mm,
                    part_spacing_mm=part_spacing_mm,
                    edge_margin_mm=edge_margin_mm,
                    cache=shape_cache,
                )
                if candidate is None:
                    continue
                candidate_score = (0.0, *_shape_candidate_score(candidate), float(sheet_index))
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
                        grid_mm=grid_mm,
                        part_spacing_mm=part_spacing_mm,
                        edge_margin_mm=edge_margin_mm,
                        cache=shape_cache,
                    )
                    if candidate is None:
                        continue
                    source_priority = 0.0 if str(stock_profile.get("source_kind", "") or "").strip().lower() == "retalho" else 1.0
                    candidate_score = (1.0, source_priority, _as_float(stock_profile.get("area_mm2", 0.0), 0.0), *_shape_candidate_score(candidate), float(stock_index))
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
                        grid_mm=grid_mm,
                        part_spacing_mm=part_spacing_mm,
                        edge_margin_mm=edge_margin_mm,
                        cache=shape_cache,
                    )
                    if candidate is not None:
                        best_candidate = {"placement": candidate, "score": (2.0, _as_float(normalized_purchase.get("area_mm2", 0.0), 0.0), *_shape_candidate_score(candidate))}
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
    shape_grid_mm: float | None = None,
) -> dict[str, Any]:
    settings = merge_laser_quote_settings(laser_settings)
    nesting_options = default_nesting_options(settings)
    items, warnings = build_nesting_items(rows, settings)
    expanded = _expand_items(items)
    normalized_stock = [profile for index, row in enumerate(list(stock_sheet_candidates or [])) if (profile := _normalize_stock_sheet_candidate(row, index)) is not None]
    use_shape_engine = bool(nesting_options.get("shape_aware", True) if shape_aware is None else shape_aware)
    grid_mm = max(2.0, _as_float(nesting_options.get("shape_grid_mm", 10.0) if shape_grid_mm is None else shape_grid_mm, 10.0), 2.0)
    requested_engine = "shape" if use_shape_engine else "bbox"

    def _choose_engine_variant(*, bbox_result: dict[str, Any] | None, shape_result: dict[str, Any] | None) -> dict[str, Any]:
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
        if use_shape_engine and not shape_active and shape_reason:
            warnings.append(f"Nesting por contorno desativado automaticamente: {shape_reason}")
        best_result: dict[str, Any] | None = None
        candidate_rows: list[dict[str, Any]] = []
        candidate_errors: list[str] = []

        if use_stock_first and normalized_stock:
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
            stock_only_shape = (
                _pack_with_stock_shape(
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
                    grid_mm=grid_mm,
                )
                if shape_active
                else None
            )
            stock_only_result = _choose_engine_variant(bbox_result=stock_only_bbox, shape_result=stock_only_shape)
            candidate_rows.append(_candidate_row_from_result("Apenas stock", stock_only_result))
            best_result = stock_only_result

        if not profiles and not (use_stock_first and normalized_stock and not allow_purchase_fallback):
            raise ValueError("Define pelo menos um formato de chapa valido para a escolha automatica.")

        for profile in profiles:
            try:
                if use_stock_first and normalized_stock:
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
                    shape_result = (
                        _pack_with_stock_shape(
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
                            grid_mm=grid_mm,
                        )
                        if shape_active
                        else None
                    )
                    result = _choose_engine_variant(bbox_result=bbox_result, shape_result=shape_result)
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
                    shape_result = (
                        _pack_profile_shape(
                            items,
                            expanded,
                            profile=profile,
                            part_spacing_mm=part_spacing_mm,
                            edge_margin_mm=edge_margin_mm,
                            allow_rotate=allow_rotate,
                            settings=settings,
                            base_warnings=warnings,
                            selection_mode="auto",
                            grid_mm=grid_mm,
                        )
                        if shape_active
                        else None
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
    if use_shape_engine and not shape_active and shape_reason:
        warnings.append(f"Nesting por contorno desativado automaticamente: {shape_reason}")
    if use_stock_first and normalized_stock:
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
        shape_result = (
            _pack_with_stock_shape(
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
                grid_mm=grid_mm,
            )
            if shape_active
            else None
        )
        return _choose_engine_variant(bbox_result=bbox_result, shape_result=shape_result)
    if profile is None:
        raise ValueError("Seleciona um formato de chapa valido.")
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
    shape_result = (
        _pack_profile_shape(
            items,
            expanded,
            profile=profile,
            part_spacing_mm=part_spacing_mm,
            edge_margin_mm=edge_margin_mm,
            allow_rotate=allow_rotate,
            settings=settings,
            base_warnings=warnings,
            selection_mode="manual",
            grid_mm=grid_mm,
        )
        if shape_active
        else None
    )
    return _choose_engine_variant(bbox_result=bbox_result, shape_result=shape_result)
