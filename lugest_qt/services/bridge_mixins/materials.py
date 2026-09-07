from __future__ import annotations

import csv
import json
import math
import re
from datetime import datetime
from lugest_core.materials import profile_entry as _profile_entry, profile_sizes as _profile_sizes
from lugest_core.search import search_matches as _smart_search_matches
from lugest_qt.services.bridge_helpers import (
    _ANGLE_SECTION_OPTIONS,
    _BAR_SECTION_OPTIONS,
    _PROFILE_SECTION_OPTIONS,
    _PROFILE_STANDARD_KG_M,
    _STEEL_DENSITY_G_CM3,
    _TUBE_SECTION_OPTIONS,
    _checker_plate_weight_kg,
    _commercial_thickness_base_mm,
    _commercial_thickness_label,
    _detect_profile_catalog_from_text,
    _profile_catalog_lookup_key,
    _profile_size_lookup_key,
)
from pathlib import Path
from typing import Any


class MaterialsBackendMixin:
    """Legacy adapter for materials; see BACKEND_GUIDE.md."""

    def _next_material_id(self) -> str:
        highest = 0
        for row in self.ensure_data().get("materiais", []):
            try:
                highest = max(highest, int(str(row.get("id", "")).replace("MAT", "")))
            except Exception:
                continue
        return f"MAT{highest + 1:05d}"

    def _next_material_internal_lot(self) -> str:
        year = datetime.now().year
        prefix = f"LOTE{year}"
        data = self.ensure_data()
        highest = 0
        for row in list(data.get("materiais", []) or []):
            value = str(row.get("lote_interno", "") or "").strip().upper()
            if value.startswith(prefix) and value[len(prefix):].isdigit():
                try:
                    highest = max(highest, int(value[len(prefix):]))
                except Exception:
                    continue
        seq_key = f"material_lote:{year}"
        local_next = max(highest + 1, int(data.setdefault("seq", {}).get(seq_key, 1) or 1))
        counter_fn = getattr(self.desktop_main, "_mysql_next_counter", None)
        n = None
        if callable(counter_fn):
            try:
                n = counter_fn(seq_key, local_next)
            except Exception:
                n = None
        n = int(n or local_next)
        data.setdefault("seq", {})[seq_key] = n + 1
        return f"{prefix}{n:05d}"

    def material_family_options(self) -> list[dict[str, Any]]:
        helper = getattr(self.materia_actions, "_material_family_options", None)
        if callable(helper):
            try:
                return [dict(row or {}) for row in list(helper(include_auto=True) or [])]
            except Exception:
                pass
        return [
            {"key": "", "label": "Auto", "density": 0.0},
            {"key": "steel", "label": "Aço / Ferro", "density": 7.85},
            {"key": "stainless", "label": "Inox", "density": 7.93},
            {"key": "aluminum", "label": "Alumínio", "density": 2.70},
            {"key": "brass", "label": "Latão", "density": 8.50},
            {"key": "copper", "label": "Cobre", "density": 8.96},
        ]

    def material_family_profile(self, material: Any = "", family: Any = "") -> dict[str, Any]:
        helper = getattr(self.materia_actions, "_resolve_material_family", None)
        if callable(helper):
            try:
                profile = dict(helper(material, family) or {})
                return {
                    "key": str(profile.get("key", "") or "").strip(),
                    "label": str(profile.get("label", "") or "").strip(),
                    "density": round(self._parse_float(profile.get("density", 7.85), 7.85), 3),
                    "explicit": bool(profile.get("explicit")),
                }
            except Exception:
                pass
        family_key = str(family or "").strip()
        options = {str(row.get("key", "") or "").strip(): dict(row or {}) for row in self.material_family_options()}
        if family_key and family_key in options:
            row = options[family_key]
            return {
                "key": family_key,
                "label": str(row.get("label", "") or "").strip(),
                "density": round(self._parse_float(row.get("density", 7.85), 7.85), 3),
                "explicit": True,
            }
        return {
            "key": "steel",
            "label": "Aço / Ferro",
            "density": 7.85,
            "explicit": False,
        }

    def material_presets(self) -> dict[str, list[str]]:
        data = self.ensure_data()
        materiais = list(
            dict.fromkeys(
                list(self.desktop_main.MATERIAIS_PRESET)
                + list(data.get("materiais_hist", []))
                + [str(m.get("material", "")).strip() for m in data.get("materiais", []) if str(m.get("material", "")).strip()]
            )
        )
        espessuras = list(
            dict.fromkeys(
                [self._fmt(v) for v in self.desktop_main.ESPESSURAS_PRESET]
                + [str(v).strip() for v in data.get("espessuras_hist", []) if str(v).strip()]
                + [str(m.get("espessura", "")).strip() for m in data.get("materiais", []) if str(m.get("espessura", "")).strip()]
            )
        )
        locais = list(
            dict.fromkeys(
                list(self.desktop_main.LOCALIZACOES_PRESET)
                + [self._localizacao(m) for m in data.get("materiais", []) if self._localizacao(m)]
                + ["RETALHO"]
            )
        )
        return {
            "formatos": list(self.desktop_main.MATERIA_FORMATOS),
            "materiais": materiais,
            "espessuras": espessuras,
            "locais": locais,
        }

    def material_section_options(self, formato: Any = "") -> list[dict[str, Any]]:
        formato_txt = str(formato or "").strip().title()
        if formato_txt == "Tubo":
            return [dict(row or {}) for row in _TUBE_SECTION_OPTIONS]
        if formato_txt == "Perfil":
            return [dict(row or {}) for row in _PROFILE_SECTION_OPTIONS]
        if formato_txt == "Cantoneira":
            return [dict(row or {}) for row in _ANGLE_SECTION_OPTIONS]
        if formato_txt == "Barra":
            return [dict(row or {}) for row in _BAR_SECTION_OPTIONS]
        return []

    def material_profile_size_options(self, secao_tipo: Any = "") -> list[str]:
        lookup_key = _profile_catalog_lookup_key(secao_tipo)
        if not lookup_key:
            return []
        return _profile_sizes(lookup_key)

    def material_profile_entry(
        self,
        secao_tipo: Any = "",
        tamanho: Any = "",
        metros: Any = 0,
    ) -> dict[str, Any]:
        key = self._material_section_type("Perfil", {"secao_tipo": secao_tipo})
        row = dict(_profile_entry(key, tamanho) or {})
        if not row:
            return {}
        length_m = self._parse_float(metros, 0)
        row["metros"] = length_m
        row["peso_unid"] = round(float(row.get("kg_m", 0) or 0) * length_m, 4)
        return row

    def _material_section_type(self, formato: str, row: dict[str, Any] | None = None) -> str:
        payload = dict(row or {})
        raw_value = str(payload.get("secao_tipo", payload.get("tipo_secao", "")) or "").strip()
        material_txt = str(payload.get("material", "") or "").strip()
        formato_txt = str(formato or "").strip().title() or "Chapa"
        if formato_txt == "Tubo":
            for option in _TUBE_SECTION_OPTIONS:
                if str(option.get("label", "") or "").strip().lower() == raw_value.lower():
                    return str(option.get("key", "") or "").strip()
            token = str(raw_value or "").strip().lower()
            if token in {"redondo", "round", "tubo redondo"}:
                return "redondo"
            if token in {"quadrado", "square", "tubo quadrado"}:
                return "quadrado"
            if token in {"retangular", "rectangular", "tubo retangular"}:
                return "retangular"
            mat_token = self._norm_material_token(material_txt)
            if any(key in mat_token for key in ("redond", "diam", "ø")):
                return "redondo"
            if "retang" in mat_token:
                return "retangular"
            if "quadrad" in mat_token:
                return "quadrado"
            diametro = self._parse_dimension_mm(payload.get("diametro", 0), 0)
            comp = self._parse_dimension_mm(payload.get("comprimento", 0), 0)
            larg = self._parse_dimension_mm(payload.get("largura", 0), 0)
            if diametro > 0:
                return "redondo"
            if comp > 0 and larg > 0:
                if abs(comp - larg) <= 1e-6:
                    return "quadrado"
                return "retangular"
            return "quadrado"
        if formato_txt == "Perfil":
            for option in _PROFILE_SECTION_OPTIONS:
                if str(option.get("label", "") or "").strip().lower() == raw_value.lower():
                    return str(option.get("key", "") or "").strip()
            lookup_key = _profile_catalog_lookup_key(raw_value)
            if lookup_key:
                return lookup_key
            detected_series, _detected_size = _detect_profile_catalog_from_text(raw_value or material_txt)
            if detected_series:
                return detected_series
            normalized = str(raw_value or "").strip().upper()
            if normalized in {"L", "T", "U", "I", "H", "OUTRO"}:
                return normalized
            mat_token = self._norm_material_token(material_txt)
            if re.search(r"\bperfil\s+l\b", mat_token):
                return "L"
            if re.search(r"\bperfil\s+t\b", mat_token):
                return "T"
            if re.search(r"\bperfil\s+u\b", mat_token):
                return "U"
            if re.search(r"\bperfil\s+i\b", mat_token):
                return "I"
            if re.search(r"\bperfil\s+h\b", mat_token):
                return "H"
            return "OUTRO"
        if formato_txt == "Cantoneira":
            for option in _ANGLE_SECTION_OPTIONS:
                if str(option.get("label", "") or "").strip().lower() == raw_value.lower():
                    return str(option.get("key", "") or "").strip()
            token = str(raw_value or "").strip().lower()
            if "desigu" in token:
                return "abas_desiguais"
            if "igual" in token or token in {"l", "cantoneira"}:
                return "abas_iguais"
            mat_token = self._norm_material_token(material_txt)
            if "desigu" in mat_token:
                return "abas_desiguais"
            return "abas_iguais"
        if formato_txt == "Barra":
            for option in _BAR_SECTION_OPTIONS:
                if str(option.get("label", "") or "").strip().lower() == raw_value.lower():
                    return str(option.get("key", "") or "").strip()
            token = str(raw_value or "").strip().lower()
            if "quadrad" in token:
                return "quadrada"
            if "retang" in token:
                return "retangular"
            if "chata" in token or "barra" in token:
                return "chata"
            mat_token = self._norm_material_token(material_txt)
            if "quadrad" in mat_token:
                return "quadrada"
            if "retang" in mat_token:
                return "retangular"
            return "chata"
        if "nervurado" in self.desktop_main.norm_text(formato_txt):
            return "nervurado"
        return ""

    def _material_section_label(self, formato: str, secao_tipo: Any = "") -> str:
        key = str(secao_tipo or "").strip()
        if not key:
            return "-"
        formato_txt = str(formato or "").strip().title()
        if formato_txt == "Tubo":
            labels = {str(row.get("key", "") or "").strip(): str(row.get("label", "") or "").strip() for row in _TUBE_SECTION_OPTIONS}
            return labels.get(key, key.title())
        if formato_txt == "Perfil":
            labels = {str(row.get("key", "") or "").strip(): str(row.get("label", "") or "").strip() for row in _PROFILE_SECTION_OPTIONS}
            return labels.get(key, key)
        if formato_txt == "Cantoneira":
            labels = {str(row.get("key", "") or "").strip(): str(row.get("label", "") or "").strip() for row in _ANGLE_SECTION_OPTIONS}
            return labels.get(key, key.replace("_", " ").title())
        if formato_txt == "Barra":
            labels = {str(row.get("key", "") or "").strip(): str(row.get("label", "") or "").strip() for row in _BAR_SECTION_OPTIONS}
            return labels.get(key, key.replace("_", " ").title())
        if "nervurado" in self.desktop_main.norm_text(formato_txt):
            return "Nervurado"
        return key

    def material_geometry_preview(self, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        row = dict(payload or {})
        formato_raw = str(row.get("formato") or self.desktop_main.detect_materia_formato(row) or "Chapa").strip() or "Chapa"
        formato_norm = self.desktop_main.norm_text(formato_raw)
        formato = "Varão nervurado" if "nervurado" in formato_norm else formato_raw.title()
        material = str(row.get("material", "") or "").strip()
        material_familia = str(row.get("material_familia", row.get("familia", "")) or "").strip()
        family_profile = self.material_family_profile(material, material_familia)
        density = round(self._parse_float(family_profile.get("density", _STEEL_DENSITY_G_CM3), _STEEL_DENSITY_G_CM3), 3)
        explicit_density = self._parse_float(row.get("densidade", row.get("density", 0)), 0)
        if explicit_density > 0:
            density = round(explicit_density / 1000.0 if explicit_density > 100 else explicit_density, 3)
        espessura = str(row.get("espessura", "") or "").strip()
        commercial_espessura_label = _commercial_thickness_label(espessura)
        espessura_mm = _commercial_thickness_base_mm(espessura) or self._parse_float(espessura, 0)
        comprimento = round(self._parse_dimension_mm(row.get("comprimento", 0), 0), 3)
        largura = round(self._parse_dimension_mm(row.get("largura", 0), 0), 3)
        altura = round(self._parse_dimension_mm(row.get("altura", 0), 0), 3)
        diametro = round(self._parse_dimension_mm(row.get("diametro", 0), 0), 3)
        metros = round(self._parse_float(row.get("metros", 0), 0), 4)
        kg_m_manual = round(self._parse_float(row.get("kg_m", row.get("peso_metro", 0)), 0), 4)
        peso_existente = round(self._parse_float(row.get("peso_unid", 0), 0), 4)
        secao_tipo = self._material_section_type(formato, row)
        secao_label = self._material_section_label(formato, secao_tipo)
        kg_m = 0.0
        peso_unid = 0.0
        area_mm2 = 0.0
        altura_nominal = altura
        lookup_size_key = ""
        base_lookup_kg_m = 0.0
        uses_catalog = False
        auto_weight = True
        dimension_label = "-"
        dim_a_text = self._fmt(comprimento) if comprimento > 0 else "-"
        dim_b_text = self._fmt(largura) if largura > 0 else "-"
        calc_hint = ""

        if formato == "Chapa":
            if comprimento > 0 and largura > 0 and espessura_mm > 0:
                peso_unid = round((comprimento * largura * espessura_mm * density) / 1000000.0, 4)
            elif peso_existente > 0:
                peso_unid = peso_existente
                auto_weight = False
            dimension_label = f"{self._fmt(comprimento)} x {self._fmt(largura)} mm" if comprimento > 0 and largura > 0 else "-"
            calc_hint = "Chapa: comprimento x largura x espessura x densidade."
            if commercial_espessura_label or any(token in self.desktop_main.norm_text(" ".join(str(row.get(key, "") or "") for key in ("material", "descricao", "obs", "formato"))) for token in ("gota", "xadrez", "antiderrap")):
                table_weight = _checker_plate_weight_kg(comprimento, largura, commercial_espessura_label)
                if table_weight > 0:
                    peso_unid = round(table_weight, 4)
                calc_hint = (
                    "Chapa gota/xadrez: peso comercial baseado na tabela de referência por dimensão/área e espessura "
                    f"{commercial_espessura_label or self._fmt(espessura_mm)}."
                    if table_weight > 0
                    else (
                        "Chapa gota/xadrez: peso calculado pela espessura base "
                        f"{self._fmt(espessura_mm)} mm e espessura comercial {commercial_espessura_label or self._fmt(espessura_mm)}."
                    )
                )
        elif formato == "Tubo":
            if metros <= 0 and comprimento > 0 and largura <= 0 and diametro <= 0:
                metros = round(comprimento / 1000.0, 4)
            if secao_tipo == "redondo":
                if diametro <= 0 and comprimento > 0 and largura <= 0:
                    diametro = comprimento
                inner_d = max(0.0, diametro - (2.0 * espessura_mm))
                if diametro > 0 and espessura_mm > 0:
                    area_mm2 = max(0.0, math.pi * ((diametro ** 2) - (inner_d ** 2)) / 4.0)
                dim_a_text = f"Ø{self._fmt(diametro)}" if diametro > 0 else "-"
                dim_b_text = "-"
                dimension_label = f"Ø{self._fmt(diametro)} x {self._fmt(espessura_mm)} mm" if diametro > 0 and espessura_mm > 0 else "-"
            else:
                if largura <= 0 and altura > 0:
                    largura = altura
                if altura <= 0 and largura > 0:
                    altura = largura
                inner_w = max(0.0, comprimento - (2.0 * espessura_mm))
                inner_h = max(0.0, largura - (2.0 * espessura_mm))
                if comprimento > 0 and largura > 0 and espessura_mm > 0:
                    area_mm2 = max(0.0, (comprimento * largura) - (inner_w * inner_h))
                dim_a_text = self._fmt(comprimento) if comprimento > 0 else "-"
                dim_b_text = self._fmt(largura) if largura > 0 else "-"
                if comprimento > 0 and largura > 0 and espessura_mm > 0:
                    dimension_label = f"{self._fmt(comprimento)} x {self._fmt(largura)} x {self._fmt(espessura_mm)} mm"
            if area_mm2 > 0:
                kg_m = round((area_mm2 * density) / 1000.0, 4)
                peso_unid = round(kg_m * metros, 4)
            elif peso_existente > 0:
                peso_unid = peso_existente
                auto_weight = False
            calc_hint = "Tubo: secção metálica x densidade x comprimento da barra."
        elif formato == "Perfil":
            detected_series, detected_size = _detect_profile_catalog_from_text(material)
            if _profile_catalog_lookup_key(secao_tipo):
                uses_catalog = True
                altura_nominal = altura_nominal or self._parse_dimension_mm(row.get("perfil_tamanho", row.get("size", 0)), 0)
                if altura_nominal <= 0 and detected_series == secao_tipo:
                    altura_nominal = self._parse_dimension_mm(detected_size, 0)
                lookup_size_key = _profile_size_lookup_key(altura_nominal)
                if lookup_size_key:
                    base_lookup_kg_m = float(_PROFILE_STANDARD_KG_M.get(secao_tipo, {}).get(lookup_size_key, 0.0) or 0.0)
                if base_lookup_kg_m > 0:
                    kg_m = round(base_lookup_kg_m * (density / _STEEL_DENSITY_G_CM3), 4)
                elif kg_m_manual > 0:
                    kg_m = kg_m_manual
                    uses_catalog = False
            else:
                if detected_series and not row.get("secao_tipo"):
                    secao_tipo = detected_series
                    secao_label = self._material_section_label(formato, secao_tipo)
                if altura_nominal <= 0:
                    altura_nominal = self._parse_dimension_mm(detected_size, 0)
                kg_m = kg_m_manual
                if kg_m <= 0 and peso_existente > 0 and metros > 0:
                    kg_m = round(peso_existente / metros, 4)
            if kg_m > 0 and metros > 0:
                peso_unid = round(kg_m * metros, 4)
            elif peso_existente > 0:
                peso_unid = peso_existente
                auto_weight = False
            dim_a_text = self._fmt(altura_nominal) if altura_nominal > 0 else "-"
            dim_b_text = secao_tipo or "-"
            if secao_tipo and altura_nominal > 0:
                dimension_label = f"{secao_tipo} {self._fmt(altura_nominal)}"
            elif secao_tipo:
                dimension_label = secao_tipo
            elif altura_nominal > 0:
                dimension_label = f"{self._fmt(altura_nominal)} mm"
            if kg_m > 0:
                calc_hint = "Perfil: kg/m da secção x comprimento da barra."
                if uses_catalog:
                    calc_hint = "Perfil: catálogo técnico normalizado ajustado pela densidade da família selecionada x comprimento da barra."
        elif formato == "Cantoneira":
            if comprimento <= 0 and largura > 0:
                comprimento = largura
            if largura <= 0 and comprimento > 0:
                largura = comprimento
            if secao_tipo == "abas_iguais" and comprimento > 0:
                largura = comprimento
            if comprimento > 0 and largura > 0 and espessura_mm > 0:
                area_mm2 = max(0.0, espessura_mm * ((comprimento + largura) - espessura_mm))
                kg_m = round((area_mm2 * density) / 1000.0, 4)
                peso_unid = round(kg_m * metros, 4) if metros > 0 else 0.0
            elif peso_existente > 0:
                peso_unid = peso_existente
                auto_weight = False
            dim_a_text = self._fmt(comprimento) if comprimento > 0 else "-"
            dim_b_text = self._fmt(largura) if largura > 0 else "-"
            if comprimento > 0 and largura > 0 and espessura_mm > 0:
                dimension_label = f"{self._fmt(comprimento)} x {self._fmt(largura)} x {self._fmt(espessura_mm)} mm"
            calc_hint = "Cantoneira: área aproximada t x (a + b - t) x densidade x comprimento."
        elif formato == "Barra":
            if secao_tipo == "quadrada":
                if comprimento <= 0 and largura > 0:
                    comprimento = largura
                if largura <= 0 and comprimento > 0:
                    largura = comprimento
            elif largura <= 0 and espessura_mm > 0:
                largura = espessura_mm
            if espessura_mm <= 0 and largura > 0:
                espessura_mm = largura
            if comprimento > 0 and largura > 0:
                area_mm2 = max(0.0, comprimento * largura)
                kg_m = round((area_mm2 * density) / 1000.0, 4)
                peso_unid = round(kg_m * metros, 4) if metros > 0 else 0.0
            elif peso_existente > 0:
                peso_unid = peso_existente
                auto_weight = False
            dim_a_text = self._fmt(comprimento) if comprimento > 0 else "-"
            dim_b_text = self._fmt(largura) if largura > 0 else "-"
            if comprimento > 0 and largura > 0:
                dimension_label = f"{self._fmt(comprimento)} x {self._fmt(largura)} mm"
            calc_hint = "Barra maciça: lado A x lado B x densidade x comprimento da barra."
        elif formato == "Varão nervurado":
            if diametro <= 0 and espessura_mm > 0:
                diametro = espessura_mm
            if espessura_mm <= 0 and diametro > 0:
                espessura_mm = diametro
            if diametro > 0:
                area_mm2 = max(0.0, math.pi * (diametro ** 2) / 4.0)
                kg_m = round((area_mm2 * density) / 1000.0, 4)
                peso_unid = round(kg_m * metros, 4) if metros > 0 else 0.0
            elif kg_m_manual > 0:
                kg_m = kg_m_manual
                peso_unid = round(kg_m * metros, 4) if metros > 0 else 0.0
            elif peso_existente > 0:
                peso_unid = peso_existente
                auto_weight = False
            dim_a_text = f"Ø{self._fmt(diametro)}" if diametro > 0 else "-"
            dim_b_text = "-"
            if diametro > 0:
                dimension_label = f"Ø{self._fmt(diametro)} mm"
            calc_hint = "Varão nervurado: secção circular maciça x densidade x comprimento da barra."
        else:
            peso_unid = peso_existente
            auto_weight = False

        resolved_espessura = commercial_espessura_label or (self._fmt(espessura_mm) if espessura_mm > 0 else espessura)

        return {
            "formato": formato,
            "secao_tipo": secao_tipo,
            "secao_label": secao_label,
            "comprimento": round(comprimento, 3),
            "largura": round(largura, 3),
            "altura": round(altura_nominal if formato == "Perfil" else altura, 3),
            "diametro": round(diametro, 3),
            "espessura": resolved_espessura,
            "espessura_mm": round(espessura_mm, 3),
            "metros": round(metros, 4),
            "kg_m": round(kg_m, 4),
            "peso_unid": round(peso_unid, 4),
            "area_mm2": round(area_mm2, 3),
            "material_familia": material_familia,
            "material_familia_resolved": str(family_profile.get("key", "") or "").strip(),
            "material_familia_label": str(family_profile.get("label", "") or "").strip(),
            "densidade": density,
            "usa_catalogo": uses_catalog,
            "altura_lookup_key": lookup_size_key,
            "kg_m_catalogo": round(base_lookup_kg_m, 4),
            "dimension_label": dimension_label,
            "dim_a_text": dim_a_text,
            "dim_b_text": dim_b_text,
            "peso_auto": auto_weight,
            "calc_hint": calc_hint,
        }

    def material_price_preview(self, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        row = dict(payload or {})
        geometry = self.material_geometry_preview(row)
        formato = str(geometry.get("formato", "Chapa") or "Chapa").strip() or "Chapa"
        p_compra = self._parse_float(row.get("p_compra", 0), 0)
        preco_unid = float(
            self.materia_actions._materia_preco_unid_record(
                {
                    "formato": formato,
                    "metros": geometry.get("metros", 0),
                    "peso_unid": geometry.get("peso_unid", 0),
                    "p_compra": p_compra,
                }
            )
            or 0.0
        )
        return {
            "base_label": "EUR/m" if formato == "Tubo" else "EUR/kg",
            "p_compra": round(p_compra, 4),
            "preco_unid": round(preco_unid, 4),
            "espessura_required": formato in {"Chapa", "Tubo", "Cantoneira", "Varão nervurado"},
            **geometry,
        }

    def material_rows(self, filter_text: str = "", in_stock_only: bool = False) -> list[dict[str, Any]]:
        data = self.ensure_data()
        rows: list[dict[str, Any]] = []
        assigned_internal_lots = False
        for index, material in enumerate(data.get("materiais", [])):
            if not str(material.get("lote_interno", "") or "").strip():
                material["lote_interno"] = self._next_material_internal_lot()
                assigned_internal_lots = True
            if bool(material.get("is_sobra")) or str(material.get("Localizacao", material.get("Localização", "")) or "").strip().upper() == "RETALHO":
                self.materia_actions._hydrate_retalho_record(data, material)
            preview = self.material_price_preview(material)
            material["preco_unid"] = float(preview.get("preco_unid", 0.0) or 0.0)
            disponivel = self._parse_float(material.get("quantidade", 0), 0) - self._parse_float(material.get("reservado", 0), 0)
            if in_stock_only and disponivel <= 0:
                continue
            formato = str(material.get("formato") or self.desktop_main.detect_materia_formato(material) or "Chapa").strip()
            quality_blocked = self._material_quality_is_blocked(material)
            quality_status = str(material.get("quality_status", "") or material.get("inspection_status", "") or "").strip()
            try:
                has_contorno = bool(self._parse_material_contour_points(material.get("contorno_points", material.get("shape_points", []))))
            except Exception:
                has_contorno = bool(material.get("contorno_points") or material.get("shape_points"))
            tipo = "Retalho" if material.get("is_sobra") else "Normal"
            if has_contorno:
                tipo = f"{tipo} contorno"
            secao_txt = str(preview.get("secao_label", "") or "").strip()
            if secao_txt and secao_txt != "-":
                tipo = f"{tipo} / {secao_txt}"
            lote_interno_txt = str(material.get("lote_interno", "") or "").strip()
            lote_fornecedor_txt = str(material.get("lote_fornecedor", "") or "").strip()
            lote_txt = lote_interno_txt or lote_fornecedor_txt
            origem_lotes = list(material.get("origem_lotes_baixa", []) or [])
            if bool(material.get("is_sobra")) and origem_lotes:
                lote_txt = " + ".join(str(item or "").strip() for item in origem_lotes if str(item or "").strip()) or lote_txt
            elif bool(material.get("is_sobra")) and not lote_txt:
                lote_txt = str(material.get("origem_lote", "") or "").strip()
            espessura_raw = str(material.get("espessura", "") or "").strip()
            values = {
                "lote": lote_txt,
                "lote_fornecedor": lote_fornecedor_txt,
                "material": str(material.get("material", "")).strip(),
                "comprimento": str(preview.get("dim_a_text", self._fmt(material.get("comprimento", 0))) or "-"),
                "largura": str(preview.get("dim_b_text", self._fmt(material.get("largura", 0))) or "-"),
                "espessura": self._fmt(espessura_raw) if espessura_raw else "-",
                "quantidade": self._fmt(material.get("quantidade", 0)),
                "reservado": self._fmt(material.get("reservado", 0)),
                "formato": formato,
                "metros": self._fmt(material.get("metros", 0)),
                "peso_unid": self._fmt(preview.get("peso_unid", material.get("peso_unid", 0))),
                "p_compra": self._fmt(material.get("p_compra", 0)),
                "preco_unid": self._fmt(preview.get("preco_unid", material.get("preco_unid", 0))),
                "disponivel": self._fmt(disponivel),
                "tipo": f"{formato} / {tipo}",
                "local": self._localizacao(material),
                "id": str(material.get("id", "")).strip(),
                "scan_code": str(material.get("scan_code", "") or self.inventory_scan_code("MAT", material.get("id"))).strip(),
            }
            # Keep the free search focused on identifying the stock item. Stock
            # quantities and prices are intentionally excluded: "chapa s235jr
            # 6mm" must target thickness 6, not an 8 mm sheet with quantity 6.
            search_values = [
                values.get("lote", ""),
                values.get("lote_fornecedor", ""),
                values.get("material", ""),
                values.get("comprimento", ""),
                values.get("largura", ""),
                values.get("espessura", ""),
                values.get("formato", ""),
                values.get("tipo", ""),
                values.get("local", ""),
                values.get("id", ""),
                values.get("scan_code", ""),
                f"espessura {values.get('espessura', '')} mm",
                f"dimensoes {values.get('comprimento', '')} x {values.get('largura', '')} mm",
                preview.get("dimension_label", ""),
                preview.get("secao_label", ""),
                preview.get("secao_tipo", ""),
                preview.get("material_familia_label", ""),
                material.get("Localização", ""),
                material.get("Localizacao", ""),
                material.get("obs", ""),
                material.get("observacoes", ""),
                material.get("referencia", ""),
                material.get("descricao", ""),
            ]
            if not _smart_search_matches(search_values, filter_text):
                continue
            severity = "ok"
            if quality_blocked:
                severity = "critical"
            elif self._parse_float(material.get("quantidade", 0), 0) == 1:
                severity = "one"
            elif disponivel <= float(self.desktop_main.STOCK_VERMELHO):
                severity = "critical"
            elif disponivel <= float(self.desktop_main.STOCK_AMARELO):
                severity = "warning"
            rows.append(
                {
                    "row": values,
                    "severity": severity,
                    "band": "even" if index % 2 == 0 else "odd",
                    "record": material,
                }
            )
        if assigned_internal_lots:
            try:
                self._save(force=True)
            except Exception:
                pass
        return rows

    def material_by_id(self, material_id: str) -> dict[str, Any] | None:
        material_id = str(material_id or "").strip()
        return next((m for m in self.ensure_data().get("materiais", []) if str(m.get("id", "")).strip() == material_id), None)

    def _material_contour_bbox(self, points: list[list[float]] | list[tuple[float, float]]) -> dict[str, float]:
        raw = [(float(point[0]), float(point[1])) for point in list(points or []) if isinstance(point, (list, tuple)) and len(point) >= 2]
        if not raw:
            return {"min_x": 0.0, "min_y": 0.0, "max_x": 0.0, "max_y": 0.0, "width": 0.0, "height": 0.0}
        xs = [point[0] for point in raw]
        ys = [point[1] for point in raw]
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

    def _parse_material_contour_points(self, value: Any) -> list[list[float]]:
        if isinstance(value, dict):
            value = value.get("points", value.get("outer", value.get("outer_polygon", [])))
        raw_points: Any = value
        if isinstance(value, str):
            text = value.strip()
            if not text:
                return []
            try:
                parsed = json.loads(text)
            except Exception:
                parsed = None
            if isinstance(parsed, dict):
                raw_points = parsed.get("points", parsed.get("outer", parsed.get("outer_polygon", [])))
            elif parsed is not None:
                raw_points = parsed
            else:
                raw_points = []
                for chunk in text.replace("|", ";").split(";"):
                    piece = str(chunk or "").strip()
                    if not piece:
                        continue
                    xy = [part.strip() for part in piece.split(",")]
                    if len(xy) != 2:
                        raise ValueError("Contorno invalido. Usa o formato x,y; x,y; x,y.")
                    try:
                        raw_points.append([float(xy[0].replace(",", ".")), float(xy[1].replace(",", "."))])
                    except Exception as exc:
                        raise ValueError("Contorno invalido. Usa apenas coordenadas numericas.") from exc
        if not isinstance(raw_points, (list, tuple)):
            raise ValueError("Contorno invalido. Usa uma lista de pontos ou texto x,y; x,y.")
        points: list[list[float]] = []
        for point in list(raw_points or []):
            if not isinstance(point, (list, tuple)) or len(point) < 2:
                raise ValueError("Contorno invalido. Cada ponto deve ter X e Y.")
            x = round(float(point[0]), 3)
            y = round(float(point[1]), 3)
            candidate = [x, y]
            if points and points[-1] == candidate:
                continue
            points.append(candidate)
        if len(points) >= 2 and points[0] == points[-1]:
            points = points[:-1]
        if not points:
            return []
        if len(points) < 3:
            raise ValueError("Contorno invalido. Define pelo menos 3 pontos.")
        bbox = self._material_contour_bbox(points)
        return [
            [round(point[0] - bbox["min_x"], 3), round(point[1] - bbox["min_y"], 3)]
            for point in points
        ]

    def format_material_contour_points(self, value: Any) -> str:
        try:
            points = self._parse_material_contour_points(value)
        except Exception:
            return str(value or "").strip()
        return "; ".join(f"{self._fmt(point[0])},{self._fmt(point[1])}" for point in points)

    def _normalise_material_payload(self, payload: dict[str, Any]) -> dict[str, Any]:
        formato_raw = str(payload.get("formato", "Chapa") or "Chapa").strip() or "Chapa"
        formato_norm = self.desktop_main.norm_text(formato_raw)
        formato = "Varão nervurado" if "nervurado" in formato_norm else formato_raw.title()
        material = str(payload.get("material", "")).strip()
        material_familia = str(payload.get("material_familia", payload.get("familia", "")) or "").strip()
        espessura = str(payload.get("espessura", "")).strip()
        comprimento = self._parse_dimension_mm(payload.get("comprimento", 0), 0)
        largura = self._parse_dimension_mm(payload.get("largura", 0), 0)
        altura = self._parse_dimension_mm(payload.get("altura", 0), 0)
        diametro = self._parse_dimension_mm(payload.get("diametro", 0), 0)
        metros = self._parse_float(payload.get("metros", 0), 0)
        quantidade = self._parse_float(payload.get("quantidade", 0), 0)
        reservado = self._parse_float(payload.get("reservado", 0), 0)
        peso_unid = self._parse_float(payload.get("peso_unid", 0), 0)
        kg_m = self._parse_float(payload.get("kg_m", payload.get("peso_metro", 0)), 0)
        p_compra = self._parse_float(payload.get("p_compra", 0), 0)
        local = str(payload.get("local", "")).strip()
        lote_interno = str(payload.get("lote_interno", "") or "").strip()
        lote = str(payload.get("lote_fornecedor", "")).strip()
        secao_tipo = str(payload.get("secao_tipo", payload.get("tipo_secao", "")) or "").strip()
        contorno_points = self._parse_material_contour_points(payload.get("contorno_points", payload.get("shape_points", [])))
        if contorno_points:
            contour_bbox = self._material_contour_bbox(contorno_points)
            comprimento = max(comprimento, contour_bbox["height"])
            largura = max(largura, contour_bbox["width"])
        if material_familia:
            material_familia = str(self.material_family_profile(material, material_familia).get("key", "") or "").strip()
        else:
            material_familia = ""
        geometry = self.material_geometry_preview(
            {
                "formato": formato,
                "material": material,
                "material_familia": material_familia,
                "espessura": espessura,
                "comprimento": comprimento,
                "largura": largura,
                "altura": altura,
                "diametro": diametro,
                "metros": metros,
                "peso_unid": peso_unid,
                "kg_m": kg_m,
                "secao_tipo": secao_tipo,
            }
        )
        comprimento = float(geometry.get("comprimento", comprimento) or 0)
        largura = float(geometry.get("largura", largura) or 0)
        altura = float(geometry.get("altura", altura) or 0)
        diametro = float(geometry.get("diametro", diametro) or 0)
        espessura = str(geometry.get("espessura", espessura) or "").strip()
        metros = float(geometry.get("metros", metros) or 0)
        peso_unid = float(geometry.get("peso_unid", peso_unid) or 0)
        kg_m = float(geometry.get("kg_m", kg_m) or 0)
        secao_tipo = str(geometry.get("secao_tipo", secao_tipo) or "").strip()
        if not material or quantidade <= 0:
            raise ValueError("Material e quantidade sao obrigatorios.")
        if formato in {"Chapa", "Tubo", "Cantoneira", "Varão nervurado"} and not espessura:
            raise ValueError("Para chapa, tubo, cantoneira e varão nervurado, espessura/diâmetro e obrigatoria.")
        if reservado < 0 or reservado > quantidade:
            raise ValueError("Reserva invalida.")
        if formato == "Chapa" and (comprimento <= 0 or largura <= 0):
            raise ValueError("Para chapa, comprimento e largura sao obrigatorios.")
        if formato == "Tubo":
            if metros <= 0:
                raise ValueError("Para tubo, o comprimento da barra e obrigatorio.")
            if secao_tipo == "redondo" and diametro <= 0:
                raise ValueError("Para tubo redondo, o diametro exterior e obrigatorio.")
            if secao_tipo != "redondo" and (comprimento <= 0 or largura <= 0):
                raise ValueError("Para tubo quadrado/retangular, indica lado A e lado B.")
            if peso_unid <= 0:
                raise ValueError("Nao foi possivel calcular o peso do tubo com os dados indicados.")
        if formato == "Perfil":
            if metros <= 0:
                raise ValueError("Para perfil, o comprimento da barra e obrigatorio.")
            if not secao_tipo:
                raise ValueError("Para perfil, indica o tipo ou serie.")
            if _profile_catalog_lookup_key(secao_tipo):
                if altura <= 0:
                    raise ValueError("Para perfis de tabela, indica a altura/tamanho nominal.")
                if kg_m <= 0:
                    raise ValueError("Nao existe kg/m tabelado para a serie e altura indicadas.")
            elif kg_m <= 0:
                raise ValueError("Para perfis manuais, indica o peso por metro (kg/m).")
            if peso_unid <= 0:
                raise ValueError("Nao foi possivel calcular o peso do perfil com os dados indicados.")
        if formato == "Cantoneira":
            if metros <= 0:
                raise ValueError("Para cantoneira, o comprimento da barra e obrigatorio.")
            if comprimento <= 0 or largura <= 0:
                raise ValueError("Para cantoneira, indica aba A e aba B.")
            if peso_unid <= 0:
                raise ValueError("Nao foi possivel calcular o peso da cantoneira com os dados indicados.")
        if formato == "Barra":
            if metros <= 0:
                raise ValueError("Para barra, o comprimento da barra e obrigatorio.")
            if comprimento <= 0 or largura <= 0:
                raise ValueError("Para barra, indica lado A e lado B.")
            if peso_unid <= 0:
                raise ValueError("Nao foi possivel calcular o peso da barra com os dados indicados.")
        if formato == "Varão nervurado":
            if metros <= 0:
                raise ValueError("Para varão nervurado, o comprimento da barra e obrigatorio.")
            if diametro <= 0:
                raise ValueError("Para varão nervurado, indica o diâmetro.")
            if peso_unid <= 0:
                raise ValueError("Nao foi possivel calcular o peso do varão nervurado com os dados indicados.")
        return {
            "formato": formato,
            "material": material,
            "espessura": espessura,
            "comprimento": comprimento,
            "largura": largura,
            "altura": altura,
            "diametro": diametro,
            "metros": metros,
            "kg_m": kg_m,
            "quantidade": quantidade,
            "reservado": reservado,
            "peso_unid": peso_unid,
            "p_compra": p_compra,
            "local": local,
            "lote_interno": lote_interno,
            "lote_fornecedor": lote,
            "secao_tipo": secao_tipo,
            "material_familia": material_familia,
            "contorno_points": contorno_points,
        }

    def add_material(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        values = self._normalise_material_payload(payload)
        record = {
            "id": self._next_material_id(),
            "lote_interno": values["lote_interno"] or self._next_material_internal_lot(),
            "formato": values["formato"],
            "material": values["material"],
            "espessura": values["espessura"],
            "comprimento": values["comprimento"],
            "largura": values["largura"],
            "altura": values["altura"],
            "diametro": values["diametro"],
            "metros": values["metros"],
            "kg_m": values["kg_m"],
            "quantidade": values["quantidade"],
            "reservado": 0.0,
            "Localização": values["local"],
            "Localizacao": values["local"],
            "lote_fornecedor": values["lote_fornecedor"],
            "secao_tipo": values["secao_tipo"],
            "material_familia": values["material_familia"],
            "peso_unid": values["peso_unid"],
            "p_compra": values["p_compra"],
            "contorno_points": [list(point) for point in list(values.get("contorno_points", []) or [])],
            "preco_unid": float(
                self.materia_actions._materia_preco_unid_record(
                    {
                        "formato": values["formato"],
                        "metros": values["metros"],
                        "peso_unid": values["peso_unid"],
                        "p_compra": values["p_compra"],
                    }
                )
            ),
            "is_sobra": False,
            "atualizado_em": self.desktop_main.now_iso(),
        }
        record = self.materia_actions._hydrate_retalho_record(data, record)
        data.setdefault("materiais", []).append(record)
        self.desktop_main.push_unique(data.setdefault("materiais_hist", []), values["material"])
        if values["espessura"]:
            self.desktop_main.push_unique(data.setdefault("espessuras_hist", []), values["espessura"])
        self.desktop_main.log_stock(
            data,
            "ADICIONAR",
            f"{values['material']} {values['espessura']} qtd={values['quantidade']}",
            operador=self._current_user_label(),
        )
        self._sync_ne_from_materia()
        self._save(force=True)
        try:
            self.conjunto_refresh_prices()
        except Exception:
            pass
        return record

    def update_material(self, material_id: str, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        record = self.material_by_id(material_id)
        if record is None:
            raise ValueError("Material nÃ£o encontrado.")
        values = self._normalise_material_payload(payload)
        record.update(
            {
                "formato": values["formato"],
                "material": values["material"],
                "espessura": values["espessura"],
                "comprimento": values["comprimento"],
                "largura": values["largura"],
                "altura": values["altura"],
                "diametro": values["diametro"],
                "metros": values["metros"],
                "kg_m": values["kg_m"],
                "quantidade": values["quantidade"],
                "reservado": values["reservado"],
                "LocalizaÃ§Ã£o": values["local"],
                "Localizacao": values["local"],
                "lote_interno": str(record.get("lote_interno", "") or values["lote_interno"] or self._next_material_internal_lot()).strip(),
                "lote_fornecedor": values["lote_fornecedor"],
                "secao_tipo": values["secao_tipo"],
                "material_familia": values["material_familia"],
                "peso_unid": values["peso_unid"],
                "p_compra": values["p_compra"],
                "contorno_points": [list(point) for point in list(values.get("contorno_points", []) or [])],
                "atualizado_em": self.desktop_main.now_iso(),
            }
        )
        self.materia_actions._hydrate_retalho_record(data, record)
        record["preco_unid"] = float(self.materia_actions._materia_preco_unid_record(record))
        self.desktop_main.log_stock(
            data,
            "EDITAR",
            f"{record.get('id')} qtd={record.get('quantidade', 0)} reservado={record.get('reservado', 0)}",
            operador=self._current_user_label(),
        )
        self._sync_ne_from_materia()
        self._save(force=True)
        try:
            self.conjunto_refresh_prices()
        except Exception:
            pass
        return record

    def remove_material(self, material_id: str) -> None:
        self.remove_materials([material_id])

    def remove_materials(self, material_ids: list[str]) -> int:
        data = self.ensure_data()
        requested = {str(value or "").strip() for value in material_ids if str(value or "").strip()}
        if not requested:
            raise ValueError("Seleciona pelo menos um material.")
        rows = list(data.get("materiais", []) or [])
        removed = [row for row in rows if str(row.get("id", "") or "").strip() in requested]
        if not removed:
            raise ValueError("Os materiais selecionados já não existem.")
        data["materiais"] = [row for row in rows if str(row.get("id", "") or "").strip() not in requested]
        operator = self._current_user_label()
        for record in removed:
            self.desktop_main.log_stock(
                data,
                "REMOVER",
                f"{record.get('id')} qtd={record.get('quantidade', 0)} reservado={record.get('reservado', 0)}",
                operador=operator,
            )
        self._save(force=True)
        try:
            self.ensure_inventory_scan_codes(persist=True)
        except Exception:
            pass
        try:
            self.conjunto_refresh_prices()
        except Exception:
            pass
        return len(removed)

    def correct_material_stock(self, material_id: str, quantidade: Any, reservado: Any, metros: Any) -> dict[str, Any]:
        data = self.ensure_data()
        record = self.material_by_id(material_id)
        if record is None:
            raise ValueError("Material não encontrado.")
        qtd = self._parse_float(quantidade, -1)
        res = self._parse_float(reservado, -1)
        met = self._parse_float(metros, 0)
        if qtd < 0 or res < 0 or res > qtd:
            raise ValueError("Valores inválidos.")
        record["quantidade"] = qtd
        record["reservado"] = res
        record["metros"] = met
        record["preco_unid"] = float(self.materia_actions._materia_preco_unid_record(record))
        record["atualizado_em"] = self.desktop_main.now_iso()
        self.desktop_main.log_stock(
            data,
            "CORRIGIR",
            f"{record.get('id')} qtd={qtd} reservado={res}",
            operador=self._current_user_label(),
        )
        self._sync_ne_from_materia()
        self._save(force=True)
        return record

    def consume_material(self, material_id: str, quantidade: Any, retalho: dict[str, Any] | None = None) -> dict[str, Any]:
        data = self.ensure_data()
        record = self.material_by_id(material_id)
        if record is None:
            raise ValueError("Material não encontrado.")
        if self._material_quality_is_blocked(record):
            raise ValueError(
                f"Material {material_id} bloqueado pela qualidade: "
                f"{str(record.get('quality_status', record.get('inspection_status', '')) or 'em inspeção')}."
            )
        retalho = dict(retalho or {})
        has_retalho = any(str(retalho.get(key, "")).strip() for key in ("comprimento", "largura", "contorno_points", "shape_points", "quantidade", "metros"))
        qtd = self._parse_float(quantidade, 0)
        stock_qtd = self._parse_float(record.get("quantidade", 0), 0)
        if qtd < 0 or qtd > stock_qtd or (qtd <= 0 and not has_retalho):
            raise ValueError("Quantidade invalida.")
        retalho_row = None
        if has_retalho:
            comp = self._parse_float(retalho.get("comprimento", 0), 0)
            larg = self._parse_float(retalho.get("largura", 0), 0)
            q_retalho = self._parse_float(retalho.get("quantidade", 0), 0)
            metros = self._parse_float(retalho.get("metros", 0), 0)
            contorno_points = self._parse_material_contour_points(retalho.get("contorno_points", retalho.get("shape_points", [])))
            if contorno_points:
                contour_bbox = self._material_contour_bbox(contorno_points)
                comp = max(comp, contour_bbox["height"])
                larg = max(larg, contour_bbox["width"])
            if q_retalho <= 0:
                raise ValueError("Quantidade do retalho invalida.")
            retalho_row = {
                "id": self._next_material_id(),
                "lote_interno": self._next_material_internal_lot(),
                "formato": record.get("formato", self.desktop_main.detect_materia_formato(record)),
                "material": record.get("material", ""),
                "espessura": record.get("espessura", ""),
                "comprimento": comp,
                "largura": larg,
                "metros": metros,
                "quantidade": q_retalho,
                "reservado": 0.0,
                "Localizacao": self._localizacao(record),
                "lote_fornecedor": record.get("lote_fornecedor", ""),
                "origem_lote_interno": str(record.get("lote_interno", "") or "").strip(),
                "peso_unid": 0.0,
                "p_compra": record.get("p_compra", 0),
                "preco_unid": 0.0,
                "is_sobra": True,
                "contorno_points": [list(point) for point in list(contorno_points or [])],
                "origem_material_id": str(record.get("id", "") or "").strip(),
                "origem_lote": str(record.get("lote_fornecedor", "") or "").strip(),
                "atualizado_em": self.desktop_main.now_iso(),
            }
            self.materia_actions._hydrate_retalho_record(data, retalho_row, template=record)
        record["quantidade"] = self._parse_float(record.get("quantidade", 0), 0) - qtd
        record["atualizado_em"] = self.desktop_main.now_iso()
        if qtd > 0:
            self.desktop_main.log_stock(data, "BAIXA", f"{record.get('id')} qtd={qtd}", operador=self._current_user_label())
        if retalho_row is not None:
            data.setdefault("materiais", []).append(retalho_row)
            self.desktop_main.log_stock(
                data,
                "RETALHO",
                f"{record.get('id')} qtd={retalho_row.get('quantidade', 0)}",
                operador=self._current_user_label(),
            )
        self._sync_ne_from_materia()
        self._save(force=True)
        return record

    def material_candidates(self, material: str, espessura: str, *, include_reserved: bool = False) -> list[dict[str, Any]]:
        material_norm = self.encomendas_actions._norm_material(material)
        esp_norm = self.encomendas_actions._norm_espessura(espessura)
        rows: list[dict[str, Any]] = []
        for stock in list(self.ensure_data().get("materiais", []) or []):
            if self._material_quality_is_blocked(stock):
                continue
            if self.encomendas_actions._norm_material(stock.get("material")) != material_norm:
                continue
            if self.encomendas_actions._norm_espessura(stock.get("espessura")) != esp_norm:
                continue
            try:
                contorno_points = [list(point) for point in list(self._parse_material_contour_points(stock.get("contorno_points", stock.get("shape_points", []))) or [])]
            except Exception:
                contorno_points = []
            total_qty = self._parse_float(stock.get("quantidade", 0), 0)
            reserved = self._parse_float(stock.get("reservado", 0), 0)
            disponivel = total_qty if include_reserved else max(0.0, total_qty - reserved)
            if disponivel <= 0:
                continue
            comprimento = round(self._parse_float(stock.get("comprimento", 0), 0), 2)
            largura = round(self._parse_float(stock.get("largura", 0), 0), 2)
            dimensao = "x".join(
                part
                for part in (self._fmt(comprimento), self._fmt(largura))
                if str(part).strip() and str(part).strip() != "0"
            ) or "-"
            rows.append(
                {
                    "material_id": str(stock.get("id", "") or "").strip(),
                    "material": str(stock.get("material", "") or "").strip(),
                    "espessura": str(stock.get("espessura", "") or "").strip(),
                    "dimensao": dimensao,
                    "comprimento": comprimento,
                    "largura": largura,
                    "disponivel": round(disponivel, 2),
                    "quantidade_total": round(total_qty, 2),
                    "reservado": round(reserved, 2),
                    "local": self._localizacao(stock),
                    "lote": str(stock.get("lote_interno", "") or stock.get("lote_fornecedor", "") or "").strip(),
                    "lote_fornecedor": str(stock.get("lote_fornecedor", "") or "").strip(),
                    "origem_lote": str(stock.get("origem_lote", "") or "").strip(),
                    "origem_encomenda": str(stock.get("origem_encomenda", "") or "").strip(),
                    "peso_unid": round(self._parse_float(stock.get("peso_unid", 0), 0), 3),
                    "p_compra": round(self._parse_float(stock.get("p_compra", 0), 0), 6),
                    "is_retalho": bool(stock.get("is_sobra")),
                    "contorno_points": contorno_points,
                }
            )
        rows.sort(
            key=lambda row: (
                str(row.get("lote", "") or ""),
                float(row.get("disponivel", 0) or 0),
                str(row.get("material_id", "") or ""),
            ),
            reverse=True,
        )
        return rows

    def material_price_rows(self, formato_filter: str = "") -> list[dict[str, Any]]:
        filtro = str(formato_filter or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for material in list(self.ensure_data().get("materiais", []) or []):
            if not isinstance(material, dict):
                continue
            preview = dict(self.material_price_preview(material) or {})
            formato = str(preview.get("formato", material.get("formato", "")) or "").strip()
            if filtro and formato.lower() != filtro:
                continue
            price_kg = 0.0
            kg_m = float(preview.get("kg_m", material.get("kg_m", 0)) or 0.0)
            base_price = float(material.get("p_compra", 0) or 0.0)
            if formato == "Tubo":
                price_kg = round(base_price / kg_m, 4) if kg_m > 0 else 0.0
            else:
                price_kg = round(base_price, 4)
            rows.append(
                {
                    "id": str(material.get("id", "") or "").strip(),
                    "formato": formato,
                    "material": str(material.get("material", "") or "").strip(),
                    "secao_tipo": str(preview.get("secao_tipo", material.get("secao_tipo", "")) or "").strip(),
                    "dimension_label": str(preview.get("dimension_label", "") or "").strip(),
                    "espessura": str(preview.get("espessura", material.get("espessura", "")) or "").strip(),
                    "kg_m": round(kg_m, 4),
                    "peso_unid": round(float(preview.get("peso_unid", material.get("peso_unid", 0)) or 0.0), 4),
                    "p_compra": round(base_price, 4),
                    "price_kg": price_kg,
                    "preco_unid": round(float(preview.get("preco_unid", material.get("preco_unid", 0)) or 0.0), 4),
                    "base_label": str(preview.get("base_label", "EUR/kg") or "EUR/kg").strip(),
                    "quantidade": round(float(material.get("quantidade", 0) or 0.0), 2),
                }
            )
        rows.sort(key=lambda item: (item.get("formato", ""), item.get("material", ""), item.get("dimension_label", ""), item.get("id", "")))
        return rows

    def material_default_price_kg(self, formato: str = "", material_name: str = "") -> float:
        formato_txt = str(formato or "").strip().lower()
        material_txt = str(material_name or "").strip().lower()
        rows = list(self.material_price_rows(formato_filter=formato) or [])
        exact = [row for row in rows if material_txt and material_txt in str(row.get("material", "") or "").strip().lower()]
        source = exact or rows
        values = [float(row.get("price_kg", 0) or 0.0) for row in source if float(row.get("price_kg", 0) or 0.0) > 0]
        if not values:
            return 0.0
        return round(sum(values) / len(values), 4)

    def material_update_price_kg(self, material_id: str, price_kg: Any) -> dict[str, Any]:
        record = self.material_by_id(material_id)
        if record is None:
            raise ValueError("Material não encontrado.")
        price_kg_value = round(self._parse_float(price_kg, 0), 4)
        if price_kg_value <= 0:
            raise ValueError("Preço/kg inválido.")
        preview = dict(self.material_price_preview(record) or {})
        formato = str(preview.get("formato", record.get("formato", "")) or "").strip() or self.desktop_main.detect_materia_formato(record)
        kg_m = float(preview.get("kg_m", record.get("kg_m", 0)) or 0.0)
        if str(formato).strip().lower() == "tubo":
            base_value = round(price_kg_value * kg_m, 4) if kg_m > 0 else 0.0
        else:
            base_value = price_kg_value
        if base_value <= 0:
            raise ValueError("Não foi possível calcular o preço base do material.")
        data = self.ensure_data()
        record["p_compra"] = base_value
        record["preco_unid"] = float(self.materia_actions._materia_preco_unid_record(record))
        record["atualizado_em"] = self.desktop_main.now_iso()
        self.desktop_main.log_stock(
            data,
            "PRECO",
            f"{record.get('id')} base={base_value} price_kg={price_kg_value}",
            operador=self._current_user_label(),
        )
        self._sync_ne_from_materia()
        self._save(force=True)
        updated_preview = dict(self.material_price_preview(record) or {})
        return {
            "id": str(record.get("id", "") or "").strip(),
            "p_compra": round(float(record.get("p_compra", 0) or 0.0), 4),
            "price_kg": round(price_kg_value, 4),
            "preco_unid": round(float(updated_preview.get("preco_unid", record.get("preco_unid", 0)) or 0.0), 4),
            "base_label": str(updated_preview.get("base_label", "EUR/kg") or "EUR/kg").strip(),
        }

    def laser_sheet_stock_candidates(self, material: str, espessura: str, *, include_reserved: bool = False) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        for index, row in enumerate(self.material_candidates(material, espessura, include_reserved=include_reserved)):
            width_mm = round(self._parse_float(row.get("largura", 0), 0), 3)
            height_mm = round(self._parse_float(row.get("comprimento", 0), 0), 3)
            quantity_available = max(0, int(self._parse_float(row.get("disponivel", 0), 0)))
            if width_mm <= 0.0 or height_mm <= 0.0 or quantity_available <= 0:
                continue
            is_retalho = bool(row.get("is_retalho"))
            material_id = str(row.get("material_id", "") or "").strip()
            source_kind = "retalho" if is_retalho else "stock"
            source_label = str(row.get("dimensao", "") or "").strip() or f"{height_mm:g} x {width_mm:g}"
            lot_label = str(row.get("lote", "") or "").strip() or str(row.get("origem_lote", "") or "").strip()
            if not lot_label:
                lot_label = material_id or f"stock-{index + 1}"
            rows.append(
                {
                    "name": f"{'Retalho' if is_retalho else 'Stock'} {lot_label} | {source_label}",
                    "source_kind": source_kind,
                    "source_label": f"{'Retalho' if is_retalho else 'Stock'} {lot_label}",
                    "material_id": material_id,
                    "lote": lot_label,
                    "local": str(row.get("local", "") or "").strip(),
                    "material": str(row.get("material", "") or "").strip(),
                    "espessura": str(row.get("espessura", "") or "").strip(),
                    "width_mm": width_mm,
                    "height_mm": height_mm,
                    "area_mm2": round(width_mm * height_mm, 2),
                    "quantity_available": quantity_available,
                    "p_compra": round(self._parse_float(row.get("p_compra", 0), 0), 6),
                    "peso_unid": round(self._parse_float(row.get("peso_unid", 0), 0), 3),
                    "is_retalho": is_retalho,
                    "outer_polygons": [[list(point) for point in list(row.get("contorno_points", []) or [])]] if list(row.get("contorno_points", []) or []) else [],
                }
            )
        rows.sort(
            key=lambda row: (
                0 if str(row.get("source_kind", "") or "") == "retalho" else 1,
                float(row.get("area_mm2", 0) or 0),
                str(row.get("lote", "") or ""),
                str(row.get("material_id", "") or ""),
            )
        )
        return rows

    def _inherit_retalho_source_geometry(self, retalho: dict[str, Any], source_stock: dict[str, Any]) -> dict[str, Any]:
        record = dict(retalho or {})
        source = dict(source_stock or {})
        formato = str(record.get("formato", "") or source.get("formato", "") or self.desktop_main.detect_materia_formato(source)).strip()
        formato_norm = self.desktop_main.norm_text(formato)
        record["formato"] = formato or "Chapa"
        numeric_keys = {"altura", "diametro", "kg_m"}
        for key in ("material_familia", "secao_tipo", "tipo", "altura", "diametro", "kg_m"):
            missing_value = self._parse_float(record.get(key, 0), 0) <= 0 if key in numeric_keys else not str(record.get(key, "") or "").strip()
            if missing_value:
                record[key] = source.get(key, record.get(key, "" if key not in {"altura", "diametro", "kg_m"} else 0.0))
        if not str(record.get("lote_fornecedor", "") or "").strip():
            record["lote_fornecedor"] = str(source.get("lote_fornecedor", "") or "").strip()
        source_lote_interno = str(source.get("lote_interno", "") or "").strip()
        source_lote_fornecedor = str(source.get("lote_fornecedor", "") or "").strip()
        source_lote = source_lote_fornecedor or source_lote_interno
        record["origem_material_id"] = str(source.get("id", "") or "").strip()
        record["origem_lote_interno"] = source_lote_interno
        record["origem_lote"] = source_lote
        origem_lotes = [str(item or "").strip() for item in list(record.get("origem_lotes_baixa", []) or []) if str(item or "").strip()]
        record["origem_lotes_baixa"] = list(dict.fromkeys(origem_lotes + [source_lote]))
        if not str(record.get("Localizacao", "") or record.get("Localização", "") or "").strip():
            record["Localizacao"] = "RETALHO"
        if formato_norm != "chapa":
            if self._parse_float(record.get("comprimento", 0), 0) <= 0:
                record["comprimento"] = self._parse_float(source.get("comprimento", 0), 0)
            if self._parse_float(record.get("largura", 0), 0) <= 0:
                record["largura"] = self._parse_float(source.get("largura", 0), 0)
            if self._parse_float(record.get("metros", 0), 0) <= 0:
                record["metros"] = self._parse_float(source.get("metros", 0), 0)
            kg_m = self._parse_float(record.get("kg_m", source.get("kg_m", 0)), 0)
            metros = self._parse_float(record.get("metros", 0), 0)
            if kg_m > 0 and metros > 0:
                record["peso_unid"] = round(kg_m * metros, 4)
        if self._parse_float(record.get("p_compra", 0), 0) <= 0:
            record["p_compra"] = source.get("p_compra", 0)
        return record

    def consume_material_allocations(
        self,
        allocations: list[dict[str, Any]],
        *,
        retalho: dict[str, Any] | None = None,
        source_material_id: str = "",
        reason: str = "",
    ) -> dict[str, Any]:
        data = self.ensure_data()
        cleaned: list[tuple[dict[str, Any], float]] = []
        for row in list(allocations or []):
            material_id = str((row or {}).get("material_id", "") or "").strip()
            qty = self._parse_float((row or {}).get("quantidade", 0), 0)
            if not material_id or qty <= 0:
                continue
            stock = self.material_by_id(material_id)
            if stock is None:
                raise ValueError(f"Material não encontrado: {material_id}")
            if self._material_quality_is_blocked(stock):
                raise ValueError(f"Material {material_id} bloqueado pela qualidade.")
            if qty > self._parse_float(stock.get("quantidade", 0), 0):
                raise ValueError(f"Quantidade superior ao stock em {material_id}.")
            cleaned.append((stock, qty))
        if not cleaned:
            raise ValueError("Nenhuma quantidade definida para baixa.")

        retalho_payload = dict(retalho or {})
        has_retalho = any(str(retalho_payload.get(key, "")).strip() for key in ("comprimento", "largura", "quantidade", "metros"))
        chosen_source_id = str(source_material_id or "").strip()
        if has_retalho and not chosen_source_id:
            unique_ids = {str(stock.get("id", "") or "").strip() for stock, _qty in cleaned}
            if len(unique_ids) == 1:
                chosen_source_id = next(iter(unique_ids))
            else:
                raise ValueError("Seleciona o lote de origem do retalho.")
        source_stock = self.material_by_id(chosen_source_id) if chosen_source_id else None
        if chosen_source_id and source_stock is None:
            raise ValueError("Lote de origem do retalho não encontrado.")
        if source_stock is not None and not any(str(stock.get("id", "") or "").strip() == chosen_source_id for stock, _qty in cleaned):
            raise ValueError("O lote escolhido para o retalho tem de fazer parte da baixa.")

        consumed_total = 0.0
        used_lots: list[str] = []
        for stock, qty in cleaned:
            stock["quantidade"] = max(0.0, self._parse_float(stock.get("quantidade", 0), 0) - qty)
            stock["atualizado_em"] = self.desktop_main.now_iso()
            consumed_total += qty
            lote = str(stock.get("lote_fornecedor", "") or stock.get("lote_interno", "") or "").strip()
            if lote:
                used_lots.append(lote)
            obs = f"{stock.get('id', '')} qtd={qty}"
            if reason:
                obs = f"{obs} motivo={reason}"
            self.desktop_main.log_stock(data, "BAIXA", obs, operador=self._current_user_label())

        created_retalho = None
        if has_retalho and source_stock is not None:
            comp = self._parse_float(retalho_payload.get("comprimento", 0), 0)
            larg = self._parse_float(retalho_payload.get("largura", 0), 0)
            q_retalho = self._parse_float(retalho_payload.get("quantidade", 0), 0)
            metros = self._parse_float(retalho_payload.get("metros", 0), 0)
            contorno_points = self._parse_material_contour_points(retalho_payload.get("contorno_points", retalho_payload.get("shape_points", [])))
            if contorno_points:
                contour_bbox = self._material_contour_bbox(contorno_points)
                comp = max(comp, contour_bbox["height"])
                larg = max(larg, contour_bbox["width"])
            if q_retalho <= 0:
                raise ValueError("Quantidade do retalho invalida.")
            created_retalho = {
                "id": self._next_material_id(),
                "lote_interno": self._next_material_internal_lot(),
                "formato": source_stock.get("formato", self.desktop_main.detect_materia_formato(source_stock)),
                "material": source_stock.get("material", ""),
                "espessura": source_stock.get("espessura", ""),
                "comprimento": comp,
                "largura": larg,
                "metros": metros,
                "quantidade": q_retalho,
                "reservado": 0.0,
                "Localização": "RETALHO",
                "Localizacao": "RETALHO",
                "lote_fornecedor": source_stock.get("lote_fornecedor", ""),
                "peso_unid": 0.0,
                "p_compra": source_stock.get("p_compra", 0),
                "preco_unid": 0.0,
                "is_sobra": True,
                "contorno_points": [list(point) for point in list(contorno_points or [])],
                "origem_material_id": str(source_stock.get("id", "") or "").strip(),
                "origem_lote": str(source_stock.get("lote_fornecedor", "") or "").strip(),
                "origem_lotes_baixa": list(dict.fromkeys(used_lots)),
                "atualizado_em": self.desktop_main.now_iso(),
            }
            created_retalho = self._inherit_retalho_source_geometry(created_retalho, source_stock)
            self.materia_actions._hydrate_retalho_record(data, created_retalho, template=source_stock)
            created_retalho["preco_unid"] = float(self.materia_actions._materia_preco_unid_record(created_retalho))
            data.setdefault("materiais", []).append(created_retalho)
            log_msg = f"{source_stock.get('id', '')} qtd={created_retalho.get('quantidade', 0)}"
            if reason:
                log_msg = f"{log_msg} motivo={reason}"
            self.desktop_main.log_stock(data, "RETALHO", log_msg, operador=self._current_user_label())
        self._sync_ne_from_materia()
        self._save(force=True)
        return {
            "consumed_total": round(consumed_total, 2),
            "retalho_id": str((created_retalho or {}).get("id", "") or "").strip(),
            "used_lots": list(dict.fromkeys(used_lots)),
        }

    def export_materials_csv(self, path: str | Path, filter_text: str = "") -> Path:
        target = Path(path)
        rows = self.material_rows(filter_text)
        with target.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.writer(handle, delimiter=";")
            writer.writerow(
                [
                    "Lote",
                    "Material",
                    "Comprimento",
                    "Largura",
                    "Espessura",
                    "Quantidade",
                    "Reserva",
                    "Formato",
                    "Metros (m)",
                    "Peso/Un. (kg)",
                    "Compra (EUR/kg|EUR/m)",
                    "Preco/Unid (EUR)",
                    "Disponivel",
                    "Tipo",
                    "Localizacao",
                    "ID",
                ]
            )
            for row in rows:
                values = row["row"]
                writer.writerow(
                    [
                        values["lote"],
                        values["material"],
                        values["comprimento"],
                        values["largura"],
                        values["espessura"],
                        values["quantidade"],
                        values["reservado"],
                        values["formato"],
                        values["metros"],
                        values["peso_unid"],
                        values["p_compra"],
                        values["preco_unid"],
                        values["disponivel"],
                        values["tipo"],
                        values["local"],
                        values["id"],
                    ]
                )
        return target

    def stock_log_rows(self, limit: int = 18) -> list[dict[str, str]]:
        rows = []
        for entry in list(reversed(self.ensure_data().get("stock_log", [])[-limit:])):
            rows.append(
                {
                    "data": str(entry.get("data", "")),
                    "acao": str(entry.get("acao", "")),
                    "operador": str(entry.get("operador", "") or "").strip(),
                    "detalhes": str(entry.get("detalhes", "")),
                }
            )
        return rows
