from __future__ import annotations

from lugest_core.operation_costing import OperationCostingEngine

from lugest_core.laser.quote_engine import (
    estimate_laser_quote,
    estimate_profile_laser_quote,
    merge_laser_quote_settings,
)
from typing import Any


class OperationCostingBackendMixin:
    """Legacy adapter for operation costing; see BACKEND_GUIDE.md."""

    def laser_quote_settings(self) -> dict[str, Any]:
        cfg = self._load_qt_config()
        return merge_laser_quote_settings(dict(cfg.get("laser_quote", {}) or {}))

    def laser_quote_save_settings(self, payload: dict[str, Any]) -> dict[str, Any]:
        cfg = self._load_qt_config()
        merged = merge_laser_quote_settings(dict(payload or {}))
        cfg["laser_quote"] = merged
        self._save_qt_config(cfg)
        return merge_laser_quote_settings(dict(cfg.get("laser_quote", {}) or {}))

    def laser_quote_analyze(self, payload: dict[str, Any]) -> dict[str, Any]:
        return estimate_laser_quote(dict(payload or {}), self.laser_quote_settings())

    def laser_quote_build_line(self, payload: dict[str, Any]) -> dict[str, Any]:
        analysis = self.laser_quote_analyze(payload)
        return {
            "analysis": analysis,
            "line": dict(analysis.get("line_suggestion", {}) or {}),
        }

    def profile_laser_quote_analyze(self, payload: dict[str, Any]) -> dict[str, Any]:
        return estimate_profile_laser_quote(dict(payload or {}), self.laser_quote_settings())

    def profile_laser_quote_build_line(self, payload: dict[str, Any]) -> dict[str, Any]:
        analysis = self.profile_laser_quote_analyze(payload)
        return {
            "analysis": analysis,
            "line": dict(analysis.get("line_suggestion", {}) or {}),
        }

    def _default_operation_cost_settings(self) -> dict[str, Any]:
        return OperationCostingEngine.default_settings()

    def _merge_operation_cost_settings(self, stored: dict[str, Any] | None = None) -> dict[str, Any]:
        return self._operation_costing_engine().merge_settings(stored)

    def operation_cost_settings(self) -> dict[str, Any]:
        cfg = self._load_qt_config()
        return self._merge_operation_cost_settings(dict(cfg.get("operation_costing", {}) or {}))

    def operation_cost_save_settings(self, payload: dict[str, Any]) -> dict[str, Any]:
        cfg = self._load_qt_config()
        merged = self._merge_operation_cost_settings(dict(payload or {}))
        cfg["operation_costing"] = merged
        self._save_qt_config(cfg)
        return self._merge_operation_cost_settings(dict(cfg.get("operation_costing", {}) or {}))

    def operation_cost_estimate(self, payload: dict[str, Any]) -> dict[str, Any]:
        return self._operation_costing_engine().estimate(payload, self.operation_cost_settings())

    def _operation_costing_engine(self) -> OperationCostingEngine:
        return OperationCostingEngine(
            available_operations=self.desktop_main.OFF_OPERACOES_DISPONIVEIS,
            normalize_operation=self.desktop_main.normalize_operacao_nome,
            parse_number=self._parse_float,
            parse_operations=self.quote_parse_operacoes_lista,
        )
