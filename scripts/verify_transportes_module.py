from __future__ import annotations

import sys
import tempfile
from datetime import date, timedelta
from pathlib import Path
from uuid import uuid4

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


def _ensure_client(backend: LegacyBackend) -> str:
    rows = list(backend.order_clients() or [])
    if rows:
        return str(rows[0].get("codigo", "") or "").strip()
    data = backend.ensure_data()
    client = {
        "codigo": "CL-TRANSPORT-TEST",
        "nome": "Cliente Teste Transportes",
        "nif": "999999991",
        "morada": "Rua da Validacao 1",
        "contacto": "910000000",
        "email": "teste.transportes@lugest.local",
    }
    data.setdefault("clientes", []).append(client)
    return str(client["codigo"])


def _ensure_supplier(backend: LegacyBackend) -> tuple[str, str]:
    data = backend.ensure_data()
    supplier_id = f"FOR-TR-{uuid4().hex[:6].upper()}"
    supplier_name = f"Transportes Validacao {supplier_id[-4:]}"
    data.setdefault("fornecedores", []).append(
        {
            "id": supplier_id,
            "nome": supplier_name,
            "nif": "999999992",
            "morada": "Parque logistico de validacao",
            "contacto": "910000002",
            "email": f"{supplier_id.lower()}@lugest.local",
            "cond_pagamento": "30 dias",
            "prazo_entrega_dias": 1,
            "website": "",
            "obs": "Fornecedor temporario de validacao do modulo transportes.",
        }
    )
    return supplier_id, supplier_name


def main() -> None:
    backend = LegacyBackend()
    backend.reload(force=True)
    backend._save = lambda force=False: None  # type: ignore[assignment]
    trip_number = f"TR-VERIFY-{uuid4().hex[:8].upper()}"
    pdf_path = Path(tempfile.gettempdir()) / "lugest_transport_verify.pdf"
    created_order = ""
    try:
        client_code = _ensure_client(backend)
        supplier_id, supplier_name = _ensure_supplier(backend)
        tomorrow = (date.today() + timedelta(days=1)).isoformat()
        zone_name = "Grande Porto"
        tariff = backend.transport_tariff_save(
            {
                "transportadora_id": supplier_id,
                "transportadora_nome": supplier_name,
                "zona": zone_name,
                "valor_base": 15.0,
                "valor_por_palete": 7.5,
                "valor_por_kg": 0.01,
                "valor_por_m3": 12.0,
                "custo_minimo": 30.0,
                "ativo": True,
                "observacoes": "Tarifario temporario de validacao.",
            }
        )
        assert tariff.get("id"), "Tarifario nao criado."
        expected_suggested_cost = 101.5
        order = backend.order_create_or_update(
            {
                "cliente": client_code,
                "posto_trabalho": "Maquina 3030",
                "nota_cliente": "Validacao modulo transportes",
                "nota_transporte": "Subcontratado",
                "local_descarga": "Cliente Teste Transportes - cais principal",
                "preco_transporte": 125.5,
                "custo_transporte": 0,
                "paletes": 3,
                "peso_bruto_kg": 1450,
                "volume_m3": 4.125,
                "transportadora_id": supplier_id,
                "transportadora_nome": supplier_name,
                "referencia_transporte": "EXT-TR-001",
                "zona_transporte": zone_name,
                "data_entrega": tomorrow,
                "tempo_estimado": 1.0,
                "observacoes": "Criado pelo verify_transportes_module.py",
                "cativar": False,
            }
        )
        created_order = str(order.get("numero", "") or "").strip()
        assert created_order, "Encomenda temporaria nao criada."

        trip = backend.transport_create_or_update(
            {
                "numero": trip_number,
                "tipo_responsavel": "Subcontratado",
                "estado": "Planeado",
                "data_planeada": tomorrow,
                "hora_saida": "08:00",
                "viatura": "Carga externa",
                "matricula": "00-TS-00",
                "motorista": "Motorista teste",
                "telefone_motorista": "910000001",
                "origem": "Instalacoes LuGEST",
                "transportadora_id": supplier_id,
                "transportadora_nome": supplier_name,
                "referencia_transporte": "EXT-TR-001",
                "custo_previsto": 0,
                "observacoes": "Viagem criada pelo verify_transportes_module.py",
            }
        )
        assert str(trip.get("numero", "") or "").strip() == trip_number, "Viagem nao criada."

        backend.transport_assign_orders(trip_number, [created_order])
        detail = backend.transport_detail(trip_number)
        stops = list(detail.get("paragens", []) or [])
        stop = next(
            (
                row
                for row in stops
                if str(row.get("encomenda_numero", "") or "").strip() == created_order
            ),
            {},
        )
        assert stop, "Encomenda nao ficou afeta a viagem."
        assert str(detail.get("tipo_responsavel", "") or "").strip() == "Subcontratado", "Tipo da viagem incorreto."
        assert str(detail.get("transportadora_nome", "") or "").strip() == supplier_name, "Transportadora da viagem nao foi guardada."
        assert abs(float(detail.get("paletes", 0) or 0) - 3.0) < 0.001, "Total de paletes incorreto."
        assert abs(float(detail.get("peso_bruto_kg", 0) or 0) - 1450.0) < 0.001, "Peso total incorreto."
        assert abs(float(detail.get("volume_m3", 0) or 0) - 4.125) < 0.001, "Volume total incorreto."
        assert abs(float(detail.get("preco_total", 0) or 0) - 125.5) < 0.001, "Preco total incorreto."
        assert abs(float(detail.get("custo_sugerido_total", 0) or 0) - expected_suggested_cost) < 0.001, "Custo sugerido total incorreto."
        assert str(stop.get("transportadora_nome", "") or "").strip() == supplier_name, "Transportadora da paragem nao foi herdada."
        assert str(stop.get("zona_transporte", "") or "").strip() == zone_name, "Zona da paragem nao foi herdada."
        assert abs(float(stop.get("paletes", 0) or 0) - 3.0) < 0.001, "Paletes da paragem incorretas."
        assert abs(float(stop.get("preco_transporte", 0) or 0) - 125.5) < 0.001, "Preco da paragem incorreto."
        assert abs(float(stop.get("custo_transporte", 0) or 0) - expected_suggested_cost) < 0.001, "Custo da paragem incorreto."
        assert abs(float(stop.get("custo_sugerido", 0) or 0) - expected_suggested_cost) < 0.001, "Sugestao de custo incorreta."
        assert str(stop.get("tarifario_id", "") or "").strip(), "Tarifario da paragem nao ficou associado."

        guide_number = f"GT-VERIFY-{uuid4().hex[:6].upper()}"
        backend.ensure_data().setdefault("expedicoes", []).append(
            {
                "numero": guide_number,
                "encomenda": created_order,
                "cliente": client_code,
                "cliente_nome": "Cliente Teste Transportes",
                "tipo": "Expedicao",
                "data_emissao": tomorrow,
                "data_transporte": tomorrow,
                "local_descarga": "Cliente Teste Transportes - cais principal",
                "estado": "Emitida",
                "anulada": False,
                "linhas": [],
            }
        )
        guide_options = backend.transport_guide_options(created_order)
        assert any(str(row.get("numero", "") or "").strip() == guide_number for row in guide_options), "Guia da encomenda nao apareceu nas opcoes."
        detail = backend.transport_update_stop(
            trip_number,
            created_order,
            {
                "expedicao_numero": guide_number,
                "local_descarga": "Cliente Teste Transportes - cais lateral",
                "contacto": "Carlos Silva",
                "telefone": "919999999",
                "data_planeada": f"{tomorrow}T09:30:00",
                "check_carga_ok": True,
                "check_docs_ok": True,
                "check_paletes_ok": True,
                "pod_estado": "Recebido",
                "pod_recebido_nome": "Carlos Silva",
                "pod_recebido_at": f"{tomorrow}T10:05:00",
                "pod_obs": "Entrega sem reservas.",
                "observacoes": "Ajustado pelo verify_transportes_module.py",
            },
        )
        updated_stop = next(
            (
                row
                for row in list(detail.get("paragens", []) or [])
                if str(row.get("encomenda_numero", "") or "").strip() == created_order
            ),
            {},
        )
        assert str(updated_stop.get("guia_numero", "") or "").strip() == guide_number, "Guia nao ficou associada a paragem."
        assert str(updated_stop.get("local_descarga", "") or "").strip() == "Cliente Teste Transportes - cais lateral", "Descarga da paragem nao atualizou."
        assert bool(updated_stop.get("check_carga_ok")), "Checklist de carga nao ficou guardado."
        assert str(updated_stop.get("pod_estado", "") or "").strip() == "Recebido", "POD nao ficou guardado."

        detail = backend.transport_request_service(
            trip_number,
            {
                "pedido_transporte_estado": "Confirmado",
                "transportadora_id": supplier_id,
                "transportadora_nome": supplier_name,
                "pedido_transporte_ref": "PED-EXT-001",
                "paletes_total_manual": 4,
                "peso_total_manual_kg": 1520,
                "volume_total_manual_m3": 4.380,
                "pedido_transporte_obs": "Pedido emitido ao parceiro externo.",
                "pedido_resposta_obs": "Carga aceite para recolha.",
            },
        )
        assert str(detail.get("pedido_transporte_estado", "") or "").strip() == "Confirmado", "Estado do pedido nao foi guardado."
        assert str(detail.get("transportadora_nome", "") or "").strip() == supplier_name, "Transportadora do pedido nao ficou associada."
        assert str(detail.get("pedido_transporte_ref", "") or "").strip() == "PED-EXT-001", "Referencia do pedido nao foi guardada."
        assert str(detail.get("pedido_resposta_obs", "") or "").strip() == "Carga aceite para recolha.", "Resposta do parceiro nao foi guardada."
        assert str(detail.get("pedido_confirmado_at", "") or "").strip(), "Timestamp de confirmacao nao foi preenchido."
        assert abs(float(detail.get("paletes", 0) or 0) - 4.0) < 0.001, "Paletes manuais nao foram aplicadas."
        assert abs(float(detail.get("peso_bruto_kg", 0) or 0) - 1520.0) < 0.001, "Peso manual nao foi aplicado."
        assert bool(detail.get("carga_manual")), "Carga manual devia ficar assinalada."
        detail = backend.transport_apply_suggested_cost(trip_number)
        assert abs(float(detail.get("custo_previsto", 0) or 0) - expected_suggested_cost) < 0.001, "Aplicacao do custo sugerido falhou."
        assert abs(float(detail.get("custo_total", 0) or 0) - expected_suggested_cost) < 0.001, "Custo total apos aplicar sugestao incorreto."

        backend.transport_set_stop_status(trip_number, created_order, "Carregada")
        backend.transport_set_status(trip_number, "Em carga")
        detail_after = backend.transport_detail(trip_number)
        assert str(detail_after.get("estado", "") or "").strip() == "Em carga", "Estado da viagem nao atualizou."

        backend.transport_route_sheet_render(trip_number, pdf_path)
        assert pdf_path.exists(), "Folha de rota nao foi gerada."

        print(f"transportes-ok order={created_order} trip={trip_number} stops={len(stops)} pdf={pdf_path}")
    finally:
        if pdf_path.exists():
            try:
                pdf_path.unlink()
            except Exception:
                pass


if __name__ == "__main__":
    main()
