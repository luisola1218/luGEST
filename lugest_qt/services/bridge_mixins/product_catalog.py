from __future__ import annotations

import os
import re
import unicodedata
from datetime import datetime
from lugest_core.intelligence import (
    RemoteProductAIClient as _RemoteProductAIClient,
    extract_product_attributes as _extract_product_attributes,
    normalize_product_description as _normalize_product_description,
    product_similarity as _product_similarity,
)
from typing import Any


class ProductCatalogBackendMixin:
    """Legacy adapter for product catalog; see BACKEND_GUIDE.md."""

    def product_presets(self) -> dict[str, list[str]]:
        return self.product_catalog_options()

    def _product_catalog_slug(self, value: Any, fallback: str = "item") -> str:
        text = str(value or "").strip().lower()
        clean: list[str] = []
        last_dash = False
        for char in text:
            if char.isalnum():
                clean.append(char)
                last_dash = False
            elif not last_dash:
                clean.append("-")
                last_dash = True
        slug = "".join(clean).strip("-")
        return slug or fallback

    def product_taxonomy(self) -> dict[str, Any]:
        taxonomy = {
            "categories": [
                {
                    "id": "inox",
                    "label": "Inox",
                    "icon": "INOX",
                    "badge": "Metal nobre",
                    "tone": "steel",
                    "subcategories": [
                        {"id": "inox-tubo", "label": "Tubo", "types": ["Redondo", "Quadrado", "Retangular", "Sanitario"]},
                        {"id": "inox-chapa", "label": "Chapa", "types": ["Escovada", "Polida", "Decapada", "Perfurada"]},
                        {"id": "inox-perfil", "label": "Perfil", "types": ["L", "U", "T", "Cantoneira", "Rectangular"]},
                        {"id": "inox-varao", "label": "Varao", "types": ["Redondo", "Sextavado", "Roscado"]},
                    ],
                },
                {
                    "id": "aluminio",
                    "label": "Aluminio",
                    "icon": "AL",
                    "badge": "Leve",
                    "tone": "sky",
                    "subcategories": [
                        {"id": "aluminio-chapa", "label": "Chapa", "types": ["Lisa", "Xadrez", "Lacada", "Perfurada"]},
                        {"id": "aluminio-tubo", "label": "Tubo", "types": ["Redondo", "Quadrado", "Retangular"]},
                        {"id": "aluminio-perfil", "label": "Perfil", "types": ["U", "L", "T", "Omega", "Calha"]},
                        {"id": "aluminio-varao", "label": "Varao", "types": ["Redondo", "Sextavado"]},
                    ],
                },
                {
                    "id": "ferro",
                    "label": "Ferro",
                    "icon": "FE",
                    "badge": "Estrutural",
                    "tone": "graphite",
                    "subcategories": [
                        {"id": "ferro-chapa", "label": "Chapa", "types": ["S235JR", "S275JR", "S355JR", "Galvanizada", "Decapada"]},
                        {"id": "ferro-tubo", "label": "Tubo", "types": ["Redondo", "Quadrado", "Retangular", "Mecanico"]},
                        {"id": "ferro-perfil", "label": "Perfil", "types": ["UPN", "IPE", "HEA", "HEB", "Cantoneira"]},
                        {"id": "ferro-varao", "label": "Varao", "types": ["Liso", "Nervurado", "Sextavado"]},
                    ],
                },
                {
                    "id": "epis",
                    "label": "EPIs",
                    "icon": "EPI",
                    "badge": "Seguranca",
                    "tone": "amber",
                    "subcategories": [
                        {"id": "epis-luvas", "label": "Luvas", "types": ["Corte", "Soldadura", "Nitrilo", "Termicas"]},
                        {"id": "epis-capacetes", "label": "Capacetes", "types": ["Industrial", "Eletrico", "Com viseira"]},
                        {"id": "epis-mascaras", "label": "Mascaras", "types": ["FFP2", "FFP3", "Soldadura", "Respiratoria"]},
                        {"id": "epis-botas", "label": "Botas", "types": ["S1P", "S3", "Borracha", "Soldador"]},
                        {"id": "epis-oculos", "label": "Oculos", "types": ["Transparentes", "Escuros", "Panoramicos"]},
                    ],
                },
                {
                    "id": "eletronica",
                    "label": "Eletronica",
                    "icon": "PCB",
                    "badge": "Controlo",
                    "tone": "indigo",
                    "subcategories": [
                        {"id": "eletronica-sensores", "label": "Sensores", "types": ["Indutivo", "Capacitivo", "Optico", "Temperatura"]},
                        {"id": "eletronica-automacao", "label": "Automacao", "types": ["PLC", "HMI", "Reles", "Fontes"]},
                        {
                            "id": "eletronica-cablagem",
                            "label": "Cablagem",
                            "types": ["Cabo", "Ficha", "Terminal", "Calha", "Bucim / Prensa-cabos"],
                        },
                    ],
                },
                {
                    "id": "pneumatica",
                    "label": "Pneumatica",
                    "icon": "AIR",
                    "badge": "Ar comprimido",
                    "tone": "cyan",
                    "subcategories": [
                        {"id": "pneumatica-valvulas", "label": "Valvulas", "types": ["2 vias", "3 vias", "5 vias", "Proporcional"]},
                        {"id": "pneumatica-cilindros", "label": "Cilindros", "types": ["Compacto", "ISO", "Guiado", "Sem haste"]},
                        {"id": "pneumatica-ligacoes", "label": "Ligacoes", "types": ["Reto", "Cotovelo", "T", "Regulador"]},
                    ],
                },
                {
                    "id": "hidraulica",
                    "label": "Hidraulica",
                    "icon": "HYD",
                    "badge": "Potencia",
                    "tone": "blue",
                    "subcategories": [
                        {"id": "hidraulica-valvulas", "label": "Valvulas", "types": ["Esfera", "Retencao", "Alivio", "Direcional"]},
                        {"id": "hidraulica-mangueiras", "label": "Mangueiras", "types": ["Alta pressao", "Retorno", "Aspiracao"]},
                        {"id": "hidraulica-acessorios", "label": "Acessorios", "types": ["Conexao", "Adaptador", "Vedante", "Filtro"]},
                    ],
                },
                {
                    "id": "fixacao",
                    "label": "Fixacao",
                    "icon": "FIX",
                    "badge": "Montagem",
                    "tone": "slate",
                    "subcategories": [
                        {
                            "id": "fixacao-parafusos",
                            "label": "Parafusos",
                            "types": [
                                "Fenda",
                                "Phillips / Cruz (PH)",
                                "Pozidriv (PZ)",
                                "Torx (TX)",
                                "Allen / Umbrako",
                                "Sextavado exterior",
                                "Quadrado",
                                "Seguranca",
                                "Sem acionamento",
                            ],
                        },
                        {
                            "id": "fixacao-porcas",
                            "label": "Porcas",
                            "types": ["Sextavada", "Travante / Nyloc", "Flangeada", "Castelo", "Gaiola", "Cega"],
                        },
                        {
                            "id": "fixacao-anilhas",
                            "label": "Anilhas",
                            "types": ["Lisa", "Pressao", "Dentada", "Belleville", "Vedacao"],
                        },
                        {
                            "id": "fixacao-rebites",
                            "label": "Rebites",
                            "types": ["Cego / POP", "Roscado", "Estrutural", "Maciço"],
                        },
                        {
                            "id": "fixacao-ancoragem",
                            "label": "Buchas e ancoragens",
                            "types": ["Bucha plastica", "Bucha metalica", "Quimica", "Mecanica", "Chumbadouro"],
                        },
                        {
                            "id": "fixacao-pinos",
                            "label": "Pinos e cavilhas",
                            "types": ["Cilindrico", "Conico", "Elastico", "Cavilha", "Contrapino"],
                        },
                        {
                            "id": "fixacao-inserts",
                            "label": "Inserts roscados",
                            "types": ["Helicoil", "Rebite roscado", "Madeira", "Plastico"],
                        },
                        {
                            "id": "fixacao-abracadeiras",
                            "label": "Abracadeiras",
                            "types": ["Metalica", "Nylon", "Mangueira", "Tubo"],
                        },
                    ],
                },
                {
                    "id": "consumiveis",
                    "label": "Consumiveis",
                    "icon": "CON",
                    "badge": "Uso diario",
                    "tone": "orange",
                    "subcategories": [
                        {"id": "consumiveis-abrasivos", "label": "Abrasivos", "types": ["Disco corte", "Disco flap", "Lixa", "Escova"]},
                        {"id": "consumiveis-quimicos", "label": "Quimicos", "types": ["Desengordurante", "Spray zincado", "Lubrificante", "Cola"]},
                        {"id": "consumiveis-embalagem", "label": "Embalagem", "types": ["Filme", "Fita", "Cantoneira", "Caixa"]},
                    ],
                },
                {
                    "id": "tintas-quimicos",
                    "label": "Tintas / Quimicos",
                    "icon": "CHM",
                    "badge": "Acabamento",
                    "tone": "orange",
                    "subcategories": [
                        {
                            "id": "tintas-revestimentos",
                            "label": "Tintas e revestimentos",
                            "types": ["Esmalte", "Primario", "Tinta tecnica", "Verniz", "Spray"],
                        },
                        {
                            "id": "tintas-solventes",
                            "label": "Solventes e diluentes",
                            "types": ["Diluente", "Desengordurante", "Acetona", "Limpeza"],
                        },
                        {
                            "id": "tintas-adesivos",
                            "label": "Adesivos e selantes",
                            "types": ["Cola", "Silicone", "Vedante", "Trava roscas"],
                        },
                        {
                            "id": "tintas-tratamento",
                            "label": "Tratamento de superficie",
                            "types": ["Zincado", "Decapante", "Passivante", "Anticorrosivo"],
                        },
                    ],
                },
                {
                    "id": "corte-laser",
                    "label": "Corte Laser",
                    "icon": "LAS",
                    "badge": "Processo",
                    "tone": "violet",
                    "subcategories": [
                        {"id": "corte-laser-consumiveis", "label": "Consumiveis", "types": ["Bico", "Lente", "Ceramica", "Filtro"]},
                        {"id": "corte-laser-gases", "label": "Gases", "types": ["Oxigenio", "Azoto", "Ar comprimido"]},
                    ],
                },
                {
                    "id": "quinagem",
                    "label": "Quinagem",
                    "icon": "QNG",
                    "badge": "Processo",
                    "tone": "purple",
                    "subcategories": [
                        {"id": "quinagem-ferramentas", "label": "Ferramentas", "types": ["Puncao", "Matriz V", "Adaptador"]},
                        {"id": "quinagem-servicos", "label": "Servicos", "types": ["Dobragens simples", "Dobragens serie", "Ajuste"]},
                    ],
                },
                {
                    "id": "soldadura",
                    "label": "Soldadura",
                    "icon": "WLD",
                    "badge": "Processo",
                    "tone": "red",
                    "subcategories": [
                        {"id": "soldadura-consumiveis", "label": "Consumiveis", "types": ["Arame MIG", "Eletrodo", "Vareta TIG", "Anti salpicos"]},
                        {"id": "soldadura-gases", "label": "Gases", "types": ["Argon", "Mistura", "CO2"]},
                        {"id": "soldadura-acessorios", "label": "Acessorios", "types": ["Tocha", "Bocal", "Difusor", "Pinca"]},
                    ],
                },
                {
                    "id": "maquinacao",
                    "label": "Maquinacao",
                    "icon": "CNC",
                    "badge": "Precisao",
                    "tone": "emerald",
                    "subcategories": [
                        {"id": "maquinacao-fresas", "label": "Fresas", "types": ["Topo", "Esferica", "Chanfrar", "Disco"]},
                        {"id": "maquinacao-brocas", "label": "Brocas", "types": ["HSS", "Carbureto", "Escalonada", "Centrar"]},
                        {"id": "maquinacao-fixacao", "label": "Fixacao maquina", "types": ["Mordaca", "Porta ferramenta", "Pinca ER"]},
                    ],
                },
                {
                    "id": "ferramentas",
                    "label": "Ferramentas",
                    "icon": "TOOL",
                    "badge": "Oficina",
                    "tone": "stone",
                    "subcategories": [
                        {"id": "ferramentas-corte", "label": "Corte", "types": ["Serra", "X-ato", "Tesoura chapa", "Corta tubos"]},
                        {"id": "ferramentas-medicao", "label": "Medicao", "types": ["Paquimetro", "Micrometro", "Esquadro", "Fita metrica"]},
                        {"id": "ferramentas-manuais", "label": "Manuais", "types": ["Chave", "Alicate", "Martelo", "Torquimetro"]},
                    ],
                },
                {
                    "id": "rolamentos-transmissao",
                    "label": "Rolamentos & Transmissao",
                    "icon": "BRG",
                    "badge": "Movimento",
                    "tone": "teal",
                    "subcategories": [
                        {"id": "rolamentos", "label": "Rolamentos", "types": ["Esferas", "Rolos", "Agulhas", "Flange"]},
                        {"id": "correntes", "label": "Correntes", "types": ["Simples", "Dupla", "Inox"]},
                        {"id": "pinhoes-polia", "label": "Pinhoes e polias", "types": ["Pinhao", "Polia", "Correia", "Casquilho taper"]},
                    ],
                },
                {
                    "id": "movimentacao",
                    "label": "Movimentacao",
                    "icon": "MOV",
                    "badge": "Mobilidade",
                    "tone": "blue",
                    "subcategories": [
                        {
                            "id": "movimentacao-rodizios",
                            "label": "Rodizios industriais",
                            "types": [
                                "Giratorio",
                                "Giratorio com travao total",
                                "Fixo",
                                "Roda avulsa",
                            ],
                        },
                        {
                            "id": "movimentacao-componentes",
                            "label": "Componentes de movimentacao",
                            "types": ["Suporte", "Placa", "Eixo", "Travão"],
                        },
                    ],
                },
                {
                    "id": "motores-redutores",
                    "label": "Motores & Redutores",
                    "icon": "MOT",
                    "badge": "Acionamento",
                    "tone": "navy",
                    "subcategories": [
                        {"id": "motores", "label": "Motores", "types": ["Monofasico", "Trifasico", "Servo", "Passo a passo"]},
                        {"id": "redutores", "label": "Redutores", "types": ["Coaxial", "Sem fim", "Planetario", "Eixo paralelo"]},
                        {"id": "variadores", "label": "Variadores", "types": ["VFD", "Soft starter", "Controlador servo"]},
                    ],
                },
                {
                    "id": "plasticos-tecnicos",
                    "label": "Plasticos Tecnicos",
                    "icon": "PLA",
                    "badge": "Tecnico",
                    "tone": "sage",
                    "subcategories": [
                        {"id": "plasticos-chapa", "label": "Chapa", "types": ["PEAD", "PVC", "PTFE", "Policarbonato"]},
                        {"id": "plasticos-varao", "label": "Varao", "types": ["Nylon", "POM", "PEEK", "PTFE"]},
                        {"id": "plasticos-tubo", "label": "Tubo", "types": ["PVC", "PU", "PTFE"]},
                    ],
                },
                {
                    "id": "vedacao-borracha",
                    "label": "Vedacao & Borracha",
                    "icon": "SEAL",
                    "badge": "Estanquidade",
                    "tone": "brown",
                    "subcategories": [
                        {"id": "vedacao-juntas", "label": "Juntas", "types": ["O-ring", "Plana", "Espiral", "Cortica"]},
                        {"id": "vedacao-retentores", "label": "Retentores", "types": ["Radial", "Cassete", "V-ring"]},
                        {"id": "borracha-tecnica", "label": "Borracha tecnica", "types": ["EPDM", "NBR", "Silicone", "Neoprene"]},
                    ],
                },
                {
                    "id": "mro-manutencao",
                    "label": "MRO & Manutencao",
                    "icon": "MRO",
                    "badge": "Suporte",
                    "tone": "grey",
                    "subcategories": [
                        {"id": "mro-lubrificacao", "label": "Lubrificacao", "types": ["Massa", "Oleo", "Spray tecnico", "Doseador"]},
                        {"id": "mro-limpeza", "label": "Limpeza", "types": ["Panos", "Desengordurante", "Absorvente", "Escova"]},
                        {"id": "mro-eletrico", "label": "Material eletrico", "types": ["Disjuntor", "Contator", "Borne", "Canaleta"]},
                    ],
                },
                {
                    "id": "escritorio-papelaria",
                    "label": "Escritorio & Papelaria",
                    "icon": "OFF",
                    "badge": "Administrativo",
                    "tone": "blue",
                    "subcategories": [
                        {
                            "id": "escritorio-cadernos",
                            "label": "Cadernos e blocos",
                            "types": ["Caderno", "Bloco de notas", "Agenda", "Livro de registo"],
                        },
                        {
                            "id": "escritorio-papel",
                            "label": "Papel e etiquetas",
                            "types": ["Papel A4", "Papel A3", "Etiquetas", "Papel tecnico"],
                        },
                        {
                            "id": "escritorio-escrita",
                            "label": "Escrita e marcacao",
                            "types": ["Caneta", "Marcador", "Lapis", "Lapiseira", "Corretor"],
                        },
                        {
                            "id": "escritorio-arquivo",
                            "label": "Arquivo e organizacao",
                            "types": ["Dossier", "Pasta", "Separador", "Caixa de arquivo"],
                        },
                        {
                            "id": "escritorio-impressao",
                            "label": "Consumiveis de impressao",
                            "types": ["Toner", "Tinteiro", "Tambor", "Ribbon"],
                        },
                    ],
                },
                {
                    "id": "informatica",
                    "label": "Informatica",
                    "icon": "IT",
                    "badge": "Tecnologia",
                    "tone": "indigo",
                    "subcategories": [
                        {
                            "id": "informatica-perifericos",
                            "label": "Perifericos de computador",
                            "types": ["Rato", "Teclado", "Monitor", "Webcam", "Headset", "Colunas"],
                        },
                        {
                            "id": "informatica-computadores",
                            "label": "Computadores",
                            "types": ["Portatil", "Desktop", "Workstation", "Mini PC"],
                        },
                        {
                            "id": "informatica-tablets",
                            "label": "Tablets e dispositivos moveis",
                            "types": ["iPad", "Tablet Android", "Tablet Windows", "E-reader"],
                        },
                        {
                            "id": "informatica-redes",
                            "label": "Redes e conectividade",
                            "types": ["Switch", "Router", "Access point", "Adaptador", "Cabo de rede"],
                        },
                        {
                            "id": "informatica-armazenamento",
                            "label": "Armazenamento",
                            "types": ["SSD", "Disco rigido", "Pen USB", "Cartao de memoria"],
                        },
                    ],
                },
                {
                    "id": "outros",
                    "label": "Outros",
                    "icon": "ETC",
                    "badge": "Flexivel",
                    "tone": "neutral",
                    "subcategories": [
                        {"id": "outros-outros", "label": "Outros", "types": ["Outros"]},
                    ],
                },
            ]
        }
        try:
            learned_rules = list(dict(self._load_qt_config() or {}).get("product_catalog_learning", []) or [])
        except Exception:
            learned_rules = []
        categories = list(taxonomy.get("categories", []) or [])
        for learned in learned_rules:
            if not isinstance(learned, dict):
                continue
            category_label = str(learned.get("categoria", "") or "").strip()
            subcategory_label = str(learned.get("subcat", "") or "").strip()
            type_label = str(learned.get("tipo", "") or "").strip()
            if not category_label:
                continue
            category = next(
                (row for row in categories if str(row.get("label", "") or "").casefold() == category_label.casefold()),
                None,
            )
            if category is None:
                category = {
                    "id": self._product_catalog_slug(category_label, "categoria"),
                    "label": category_label,
                    "icon": "AI",
                    "badge": "Aprendido",
                    "tone": "teal",
                    "subcategories": [],
                }
                categories.append(category)
            if not subcategory_label:
                continue
            subcategories = category.setdefault("subcategories", [])
            subcategory = next(
                (
                    row
                    for row in subcategories
                    if str(row.get("label", "") or "").casefold() == subcategory_label.casefold()
                ),
                None,
            )
            if subcategory is None:
                subcategory = {
                    "id": self._product_catalog_slug(
                        f"{category.get('id', '')}-{subcategory_label}",
                        "subcategoria",
                    ),
                    "label": subcategory_label,
                    "types": [],
                }
                subcategories.append(subcategory)
            types = subcategory.setdefault("types", [])
            if type_label and not any(str(value or "").casefold() == type_label.casefold() for value in types):
                types.append(type_label)
        taxonomy["categories"] = categories
        return taxonomy

    def _product_taxonomy_nodes(self) -> tuple[dict[str, Any], dict[tuple[str, str], dict[str, Any]], dict[tuple[str, str, str], dict[str, Any]]]:
        cached = getattr(self, "_product_taxonomy_nodes_cache", None)
        if cached is not None:
            return cached
        category_map: dict[str, Any] = {}
        subcategory_map: dict[tuple[str, str], dict[str, Any]] = {}
        type_map: dict[tuple[str, str, str], dict[str, Any]] = {}
        for category in list(self.product_taxonomy().get("categories", []) or []):
            category_label = str(category.get("label", "") or "").strip()
            if not category_label:
                continue
            category_map[category_label.casefold()] = dict(category)
            for subcategory in list(category.get("subcategories", []) or []):
                sub_label = str(subcategory.get("label", "") or "").strip()
                if not sub_label:
                    continue
                subcategory_map[(category_label.casefold(), sub_label.casefold())] = dict(subcategory)
                for type_label in list(subcategory.get("types", []) or []):
                    clean_type = str(type_label or "").strip()
                    if clean_type:
                        type_map[(category_label.casefold(), sub_label.casefold(), clean_type.casefold())] = {"label": clean_type}
        self._product_taxonomy_nodes_cache = (category_map, subcategory_map, type_map)
        return self._product_taxonomy_nodes_cache

    def product_catalog_options(self, category: str = "", subcategory: str = "") -> dict[str, Any]:
        taxonomy = self.product_taxonomy()
        categories = [dict(row or {}) for row in list(taxonomy.get("categories", []) or []) if isinstance(row, dict)]
        category_labels = [str(row.get("label", "") or "").strip() for row in categories if str(row.get("label", "") or "").strip()]
        selected_category = str(category or "").strip()
        selected_subcategory = str(subcategory or "").strip()
        category_row = next(
            (row for row in categories if str(row.get("label", "") or "").strip().casefold() == selected_category.casefold()),
            None,
        )
        subcategories = [dict(row or {}) for row in list((category_row or {}).get("subcategories", []) or []) if isinstance(row, dict)]
        subcategory_labels = [str(row.get("label", "") or "").strip() for row in subcategories if str(row.get("label", "") or "").strip()]
        subcategory_row = next(
            (row for row in subcategories if str(row.get("label", "") or "").strip().casefold() == selected_subcategory.casefold()),
            None,
        )
        type_labels = [str(value or "").strip() for value in list((subcategory_row or {}).get("types", []) or []) if str(value or "").strip()]
        return {
            "taxonomy": taxonomy,
            "categorias": category_labels,
            "subcats": subcategory_labels,
            "tipos": type_labels,
            "unidades": [str(v) for v in list(getattr(self.desktop_main, "PROD_UNIDS", []) or [])],
            "category_meta": {
                str(row.get("label", "") or "").strip(): {
                    "id": str(row.get("id", "") or "").strip(),
                    "icon": str(row.get("icon", "") or "").strip(),
                    "badge": str(row.get("badge", "") or "").strip(),
                    "tone": str(row.get("tone", "") or "").strip(),
                }
                for row in categories
                if str(row.get("label", "") or "").strip()
            },
        }

    def _product_catalog_text(self, value: Any) -> str:
        text = unicodedata.normalize("NFKD", str(value or ""))
        text = "".join(char for char in text if not unicodedata.combining(char))
        text = text.casefold()
        text = re.sub(r"(?<=\d)\s*[x×]\s*(?=\d)", "x", text)
        text = re.sub(r"[^a-z0-9+./-]+", " ", text)
        return re.sub(r"\s+", " ", text).strip()

    def _product_learning_keyword(self, description: str) -> str:
        text = self._product_catalog_text(description)
        ignored = {
            "de", "do", "da", "dos", "das", "com", "sem", "para", "e", "em",
            "inox", "aco", "aluminio", "ferro", "zincado", "galvanizado", "preto",
            "branco", "grande", "pequeno", "industrial", "tecnico", "tecnica",
        }
        for token in re.findall(r"[a-z][a-z0-9+./-]{2,}", text):
            if token in ignored or re.fullmatch(r"m?\d+(?:[x./-]\d+)*", token):
                continue
            return token[:-1] if token.endswith("s") and len(token) > 4 else token
        return ""

    def _product_learned_catalog_suggestion(self, description: str) -> dict[str, Any]:
        text = self._product_catalog_text(description)
        try:
            rules = list(dict(self._load_qt_config() or {}).get("product_catalog_learning", []) or [])
        except Exception:
            rules = []
        matches: list[dict[str, Any]] = []
        for row in rules:
            if not isinstance(row, dict):
                continue
            keyword = self._product_catalog_text(row.get("keyword", ""))
            if keyword and re.search(rf"\b{re.escape(keyword)}\b", text):
                matches.append(dict(row))
        if not matches:
            return {}
        matches.sort(key=lambda row: len(str(row.get("keyword", ""))), reverse=True)
        learned = matches[0]
        return {
            "categoria": str(learned.get("categoria", "") or "").strip(),
            "subcat": str(learned.get("subcat", "") or "").strip(),
            "tipo": str(learned.get("tipo", "") or "").strip(),
            "dimensoes": "",
            "confidence": 0.96,
            "reason": f"Classificação aprendida para “{learned.get('keyword', '')}”",
            "learned": True,
            "needs_learning": False,
        }

    def product_catalog_teach(
        self,
        description: str,
        category: str,
        subcategory: str,
        product_type: str = "",
    ) -> dict[str, Any]:
        keyword = self._product_learning_keyword(description)
        category = str(category or "").strip()
        subcategory = str(subcategory or "").strip()
        product_type = str(product_type or "").strip()
        if not keyword:
            raise ValueError("Não foi possível identificar o termo principal da descrição.")
        if not category or not subcategory:
            raise ValueError("Seleciona pelo menos a categoria e a subcategoria antes de ensinar.")
        taxonomy_before = self.product_taxonomy()
        existing_category = next(
            (
                row
                for row in list(taxonomy_before.get("categories", []) or [])
                if str(row.get("label", "") or "").casefold() == category.casefold()
            ),
            None,
        )
        existing_subcategory = next(
            (
                row
                for row in list((existing_category or {}).get("subcategories", []) or [])
                if str(row.get("label", "") or "").casefold() == subcategory.casefold()
            ),
            None,
        )
        existing_types = {
            str(item.get("label", "") if isinstance(item, dict) else item or "").strip().casefold()
            for item in list((existing_subcategory or {}).get("types", []) or [])
        }
        cfg = self._load_qt_config()
        rules = [
            dict(row)
            for row in list(cfg.get("product_catalog_learning", []) or [])
            if isinstance(row, dict)
        ]
        learned = {
            "keyword": keyword,
            "categoria": category,
            "subcat": subcategory,
            "tipo": product_type,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }
        replaced = False
        for index, row in enumerate(rules):
            if str(row.get("keyword", "") or "").casefold() == keyword.casefold():
                rules[index] = learned
                replaced = True
                break
        if not replaced:
            rules.append(learned)
        cfg["product_catalog_learning"] = rules
        self._save_qt_config(cfg)
        return {
            **learned,
            "created_category": existing_category is None,
            "created_subcategory": existing_subcategory is None,
            "created_type": bool(product_type) and product_type.casefold() not in existing_types,
        }

    def product_catalog_suggestion(self, description: str) -> dict[str, Any]:
        """Infer catalog fields from a product description without external AI.

        Rules are deliberately deterministic and conservative so the same
        description always produces the same catalog assignment, including
        when the workstation is offline.
        """

        text = self._product_catalog_text(description)
        empty = {
            "categoria": "",
            "subcat": "",
            "tipo": "",
            "dimensoes": "",
            "confidence": 0.0,
            "reason": "",
        }
        if not text:
            return empty

        learned_suggestion = self._product_learned_catalog_suggestion(description)

        def contains(*patterns: str) -> bool:
            return any(re.search(pattern, text) for pattern in patterns)

        metric_match = re.search(
            r"\bm\s*(\d+(?:[.,]\d+)?)(?:\s*[x×]\s*(\d+(?:[.,]\d+)?))?\b",
            text,
        )
        commercial_size_match = re.search(
            r"\b(\d+(?:[.,]\d+)?)\s*[x×]\s*(\d+(?:[.,]\d+)?)\s*mm\b",
            text,
        )
        dimensions = ""
        if metric_match:
            diameter = metric_match.group(1).replace(",", ".")
            length = str(metric_match.group(2) or "").replace(",", ".")
            dimensions = f"M{diameter}" + (f"x{length}" if length else "")
        elif commercial_size_match:
            first = commercial_size_match.group(1).replace(",", ".")
            second = commercial_size_match.group(2).replace(",", ".")
            dimensions = f"{first}x{second} mm"

        if learned_suggestion:
            learned_suggestion["dimensoes"] = dimensions
            return learned_suggestion

        def result(category: str, subcategory: str, product_type: str = "", reason: str = "") -> dict[str, Any]:
            return {
                "categoria": category,
                "subcat": subcategory,
                "tipo": product_type,
                "dimensoes": dimensions,
                "confidence": 0.98 if product_type else 0.90,
                "reason": reason or f"{category} / {subcategory}",
            }

        # Fixing elements come first: material words such as "inox" describe
        # the finish/material and must not turn a screw into an Inox stock item.
        if contains(r"\bparafus", r"\bperno\b", r"\bbolt\b", r"\bscrew\b"):
            drive = ""
            if contains(r"\btorx\b", r"\btx\s*\d"):
                drive = "Torx (TX)"
            elif contains(r"\bpozidriv\b", r"\bpz\s*\d"):
                drive = "Pozidriv (PZ)"
            elif contains(r"\bphillips\b", r"\bphilips\b", r"\bph\s*\d", r"\bcruz\b"):
                drive = "Phillips / Cruz (PH)"
            elif contains(r"\ballen\b", r"\bumbrak", r"\bunbrak", r"\bsextavad[oa]\s+interior\b", r"\bhex\s+socket\b"):
                drive = "Allen / Umbrako"
            elif contains(r"\bfenda\b", r"\bslotted\b"):
                drive = "Fenda"
            elif contains(r"\bsextavad[oa]\b", r"\bhexagonal\b"):
                drive = "Sextavado exterior"
            elif contains(r"\bquadrad[oa]\b", r"\brobertson\b"):
                drive = "Quadrado"
            return result("Fixacao", "Parafusos", drive, "Descrição identificada como parafuso")

        if contains(r"\bporcas?\b", r"\bnuts?\b"):
            nut_type = ""
            if contains(
                r"\bnyloc\b",
                r"\btravante\b",
                r"\bautobloc",
                r"\bauto\s*bloc",
                r"\bauto[-\s]*bloqueio\b",
                r"\bautotravante\b",
                r"\bfreio\b",
            ):
                nut_type = "Travante / Nyloc"
            elif contains(r"\bflange"):
                nut_type = "Flangeada"
            elif contains(r"\bcastelo\b"):
                nut_type = "Castelo"
            elif contains(r"\bgaiola\b"):
                nut_type = "Gaiola"
            elif contains(r"\bcega\b"):
                nut_type = "Cega"
            elif contains(r"\bsextavad"):
                nut_type = "Sextavada"
            return result("Fixacao", "Porcas", nut_type, "Descrição identificada como porca")

        if contains(r"\banilh", r"\bwashers?\b"):
            washer_type = ""
            if contains(r"\bpressao\b", r"\bgrower\b"):
                washer_type = "Pressao"
            elif contains(r"\bdentad"):
                washer_type = "Dentada"
            elif contains(r"\bbelleville\b", r"\bprato\b"):
                washer_type = "Belleville"
            elif contains(r"\bvedacao\b", r"\bvedante\b"):
                washer_type = "Vedacao"
            elif contains(r"\blisa\b", r"\bplana\b"):
                washer_type = "Lisa"
            return result("Fixacao", "Anilhas", washer_type, "Descrição identificada como anilha")

        if contains(r"\bhelicoil\b", r"\binsert\s+roscad", r"\brebite\s+roscad"):
            insert_type = "Helicoil" if "helicoil" in text else "Rebite roscado"
            return result("Fixacao", "Inserts roscados", insert_type, "Descrição identificada como insert roscado")

        if contains(r"\brebite\b", r"\brivet\b"):
            rivet_type = ""
            if contains(r"\bpop\b", r"\bcego\b"):
                rivet_type = "Cego / POP"
            elif contains(r"\bestrutural\b"):
                rivet_type = "Estrutural"
            elif contains(r"\bmacic"):
                rivet_type = "Maciço"
            return result("Fixacao", "Rebites", rivet_type, "Descrição identificada como rebite")

        if contains(r"\bbucha\b", r"\bancor", r"\bchumbadour"):
            anchor_type = ""
            if contains(r"\bquimic"):
                anchor_type = "Quimica"
            elif contains(r"\bchumbadour"):
                anchor_type = "Chumbadouro"
            elif contains(r"\bplastic"):
                anchor_type = "Bucha plastica"
            elif contains(r"\bmetal"):
                anchor_type = "Bucha metalica"
            return result("Fixacao", "Buchas e ancoragens", anchor_type, "Descrição identificada como ancoragem")

        if contains(r"\bcavilh", r"\bcontrapino\b", r"\bpino\b"):
            pin_type = ""
            if contains(r"\bcontrapino\b"):
                pin_type = "Contrapino"
            elif contains(r"\bcavilh"):
                pin_type = "Cavilha"
            elif contains(r"\belastic"):
                pin_type = "Elastico"
            elif contains(r"\bconic"):
                pin_type = "Conico"
            elif contains(r"\bcilindr"):
                pin_type = "Cilindrico"
            return result("Fixacao", "Pinos e cavilhas", pin_type, "Descrição identificada como pino ou cavilha")

        if contains(r"\babracadeir", r"\babraçadeir"):
            clamp_type = ""
            if contains(r"\bnylon\b", r"\bplast"):
                clamp_type = "Nylon"
            elif contains(r"\bmangueira\b"):
                clamp_type = "Mangueira"
            elif contains(r"\btubo\b"):
                clamp_type = "Tubo"
            elif contains(r"\bmetal"):
                clamp_type = "Metalica"
            return result("Fixacao", "Abracadeiras", clamp_type, "Descrição identificada como abraçadeira")

        tente_model = re.search(r"\b(347[078])\s*ufr\s*(\d{3})\s*p(\d{2})\b", text)
        if contains(r"\brodizio", r"\broda\b", r"\broulett", r"\bcaster\b") or tente_model:
            caster_type = ""
            if tente_model:
                caster_type = {
                    "3470": "Giratorio",
                    "3477": "Giratorio com travao total",
                    "3478": "Fixo",
                }.get(tente_model.group(1), "")
            if not caster_type and contains(r"\btravao\b", r"\btravão\b", r"\bfreio\b", r"\btotal\s+lock\b"):
                caster_type = "Giratorio com travao total"
            elif not caster_type and contains(r"\bfix[oa]\b", r"\bfixed\b"):
                caster_type = "Fixo"
            elif not caster_type and contains(r"\bgiratori", r"\bpivotante\b", r"\bswivel\b"):
                caster_type = "Giratorio"
            elif not caster_type:
                caster_type = "Roda avulsa"
            return result(
                "Movimentacao",
                "Rodizios industriais",
                caster_type,
                "Descrição identificada como roda ou rodízio industrial",
            )

        # The remaining catalog is described by ordered family rules. Each
        # family first identifies the object and then refines its type. This
        # keeps the classifier extensible and avoids a fragile list of exact
        # product descriptions.
        catalog_rules: list[tuple[str, str, tuple[str, ...], tuple[tuple[str, tuple[str, ...]], ...]]] = [
            ("EPIs", "Luvas", (r"\bluva", r"\bglove"), (
                ("Soldadura", (r"\bsoldadur", r"\bsoldador")),
                ("Nitrilo", (r"\bnitril",)),
                ("Termicas", (r"\btermic",)),
                ("Corte", (r"\banticorte", r"\bcorte\b")),
            )),
            ("EPIs", "Capacetes", (r"\bcapacete", r"\bhelmet"), (
                ("Com viseira", (r"\bviseira",)),
                ("Eletrico", (r"\beletric",)),
                ("Industrial", (r"\bindustrial",)),
            )),
            ("EPIs", "Mascaras", (r"\bmascara", r"\brespirador"), (
                ("FFP3", (r"\bffp\s*3\b",)),
                ("FFP2", (r"\bffp\s*2\b",)),
                ("Soldadura", (r"\bsoldadur",)),
                ("Respiratoria", (r"\brespirat",)),
            )),
            ("EPIs", "Botas", (r"\bbota", r"\bsapato\s+seguranca"), (
                ("S1P", (r"\bs1p\b",)),
                ("S3", (r"\bs3\b",)),
                ("Soldador", (r"\bsoldador",)),
                ("Borracha", (r"\bborracha",)),
            )),
            ("EPIs", "Oculos", (r"\boculos", r"\bgoggle"), (
                ("Panoramicos", (r"\bpanoram",)),
                ("Escuros", (r"\bescuro", r"\bfumado")),
                ("Transparentes", (r"\btransparent",)),
            )),
            ("Eletronica", "Sensores", (r"\bsensor",), (
                ("Indutivo", (r"\bindutiv",)),
                ("Capacitivo", (r"\bcapacit",)),
                ("Optico", (r"\boptic", r"\bfotoele")),
                ("Temperatura", (r"\btemperatur", r"\btermopar")),
            )),
            ("Eletronica", "Automacao", (r"\bplc\b", r"\bhmi\b", r"\brele\b", r"\bfonte\s+(?:de\s+)?aliment"), (
                ("PLC", (r"\bplc\b",)),
                ("HMI", (r"\bhmi\b",)),
                ("Reles", (r"\brele\b",)),
                ("Fontes", (r"\bfonte",)),
            )),
            (
                "Eletronica",
                "Cablagem",
                (
                    r"\bbucim",
                    r"\bprensa[-\s]*cabos?\b",
                    r"\bcable\s+gland\b",
                    r"\bcabo\b",
                    r"\bficha\b",
                    r"\bterminal\b",
                    r"\bcalha\s+tecnica",
                ),
                (
                ("Bucim / Prensa-cabos", (r"\bbucim", r"\bprensa[-\s]*cabos?\b", r"\bcable\s+gland\b")),
                ("Cabo", (r"\bcabo\b",)),
                ("Ficha", (r"\bficha\b",)),
                ("Terminal", (r"\bterminal\b",)),
                ("Calha", (r"\bcalha",)),
                ),
            ),
            ("Pneumatica", "Valvulas", (r"\bvalvula\s+pneumat", r"\bpneumat.+\bvalvula"), (
                ("5 vias", (r"\b5\s*(?:vias|/2)\b",)),
                ("3 vias", (r"\b3\s*(?:vias|/2)\b",)),
                ("2 vias", (r"\b2\s*(?:vias|/2)\b",)),
                ("Proporcional", (r"\bproporcional",)),
            )),
            ("Pneumatica", "Cilindros", (r"\bcilindro\s+pneumat", r"\bpneumat.+\bcilindro"), (
                ("Sem haste", (r"\bsem\s+haste",)),
                ("Guiado", (r"\bguiad",)),
                ("Compacto", (r"\bcompact",)),
                ("ISO", (r"\biso\b",)),
            )),
            ("Pneumatica", "Ligacoes", (r"\bracor\s+pneumat", r"\bligacao\s+pneumat", r"\bcotovelo\s+pneumat"), (
                ("Cotovelo", (r"\bcotovelo",)),
                ("Regulador", (r"\bregulador",)),
                ("T", (r"\b(?:racor|ligacao)\s+t\b",)),
                ("Reto", (r"\breto\b",)),
            )),
            ("Hidraulica", "Valvulas", (r"\bvalvula\s+hidraulic", r"\bhidraulic.+\bvalvula"), (
                ("Retencao", (r"\bretenc",)),
                ("Alivio", (r"\balivio",)),
                ("Direcional", (r"\bdirecional",)),
                ("Esfera", (r"\besfera",)),
            )),
            ("Hidraulica", "Mangueiras", (r"\bmangueira\s+hidraulic", r"\bhidraulic.+\bmangueira"), (
                ("Alta pressao", (r"\balta\s+pressao",)),
                ("Retorno", (r"\bretorno",)),
                ("Aspiracao", (r"\baspir",)),
            )),
            ("Hidraulica", "Acessorios", (r"\b(?:conexao|adaptador|filtro)\s+hidraulic", r"\bhidraulic.+\b(?:conexao|adaptador|filtro)"), (
                ("Conexao", (r"\bconexao",)),
                ("Adaptador", (r"\badaptador",)),
                ("Vedante", (r"\bvedante",)),
                ("Filtro", (r"\bfiltro",)),
            )),
            ("Consumiveis", "Abrasivos", (r"\bdisco\s+(?:de\s+)?corte\b", r"\bdisco\s+flap\b", r"\blixa\b", r"\bescova\s+abras"), (
                ("Disco corte", (r"\bdisco\s+(?:de\s+)?corte\b",)),
                ("Disco flap", (r"\bflap\b",)),
                ("Lixa", (r"\blixa\b",)),
                ("Escova", (r"\bescova",)),
            )),
            ("Consumiveis", "Embalagem", (r"\bfilme\s+estiravel\b", r"\bfita\s+embal", r"\bcantoneira\s+cartao\b", r"\bcaixa\s+cartao\b"), (
                ("Filme", (r"\bfilme",)),
                ("Fita", (r"\bfita",)),
                ("Cantoneira", (r"\bcantoneira",)),
                ("Caixa", (r"\bcaixa",)),
            )),
            ("Corte Laser", "Consumiveis", (r"\bbico\s+laser\b", r"\blente\s+laser\b", r"\bceramica\s+laser\b", r"\bfiltro\s+laser\b"), (
                ("Bico", (r"\bbico",)),
                ("Lente", (r"\blente",)),
                ("Ceramica", (r"\bceramica",)),
                ("Filtro", (r"\bfiltro",)),
            )),
            ("Soldadura", "Consumiveis", (r"\barame\s+mig\b", r"\beletrodo\b", r"\bvareta\s+tig\b", r"\banti\s*salpico"), (
                ("Arame MIG", (r"\bmig\b",)),
                ("Vareta TIG", (r"\btig\b",)),
                ("Anti salpicos", (r"\bsalpico",)),
                ("Eletrodo", (r"\beletrodo",)),
            )),
            ("Soldadura", "Acessorios", (r"\btocha\s+(?:mig|tig|sold)", r"\bbocal\s+(?:mig|tig|sold)", r"\bdifusor\s+(?:mig|tig|sold)"), (
                ("Tocha", (r"\btocha",)),
                ("Bocal", (r"\bbocal",)),
                ("Difusor", (r"\bdifusor",)),
                ("Pinca", (r"\bpinca",)),
            )),
            ("Maquinacao", "Fresas", (r"\bfresa\b",), (
                ("Esferica", (r"\besferic",)),
                ("Chanfrar", (r"\bchanfr",)),
                ("Disco", (r"\bdisco",)),
                ("Topo", (r"\btopo",)),
            )),
            ("Maquinacao", "Brocas", (r"\bbroca\b",), (
                ("Carbureto", (r"\bcarburet",)),
                ("Escalonada", (r"\bescalon",)),
                ("Centrar", (r"\bcentrar",)),
                ("HSS", (r"\bhss\b",)),
            )),
            ("Ferramentas", "Medicao", (r"\bpaquimetro\b", r"\bmicrometro\b", r"\besquadro\b", r"\bfita\s+metrica\b"), (
                ("Paquimetro", (r"\bpaquimetro",)),
                ("Micrometro", (r"\bmicrometro",)),
                ("Esquadro", (r"\besquadro",)),
                ("Fita metrica", (r"\bfita\s+metrica",)),
            )),
            ("Ferramentas", "Manuais", (r"\bchave\b", r"\balicate\b", r"\bmartelo\b", r"\btorquimetro\b"), (
                ("Alicate", (r"\balicate",)),
                ("Martelo", (r"\bmartelo",)),
                ("Torquimetro", (r"\btorquimetro",)),
                ("Chave", (r"\bchave",)),
            )),
            ("Motores & Redutores", "Motores", (r"\bmotor\b",), (
                ("Passo a passo", (r"\bpasso\s+a\s+passo\b", r"\bstepper\b")),
                ("Servo", (r"\bservo",)),
                ("Trifasico", (r"\btrifas",)),
                ("Monofasico", (r"\bmonofas",)),
            )),
            ("Motores & Redutores", "Redutores", (r"\bredutor\b",), (
                ("Sem fim", (r"\bsem\s+fim\b", r"\bcoroa",)),
                ("Planetario", (r"\bplanet",)),
                ("Eixo paralelo", (r"\beixo\s+paralelo",)),
                ("Coaxial", (r"\bcoaxial",)),
            )),
            ("Motores & Redutores", "Variadores", (r"\bvariador\b", r"\binversor\s+frequencia\b", r"\bsoft\s*starter\b"), (
                ("Soft starter", (r"\bsoft\s*starter",)),
                ("Controlador servo", (r"\bservo",)),
                ("VFD", (r"\bvfd\b", r"\bvariador\b", r"\binversor",)),
            )),
            ("Vedacao & Borracha", "Juntas", (r"\bo[-\s]*ring\b", r"\bjunta\b"), (
                ("O-ring", (r"\bo[-\s]*ring",)),
                ("Espiral", (r"\bespiral",)),
                ("Cortica", (r"\bcortica",)),
                ("Plana", (r"\bplana",)),
            )),
            ("Vedacao & Borracha", "Retentores", (r"\bretentor\b", r"\bv[-\s]*ring\b"), (
                ("V-ring", (r"\bv[-\s]*ring",)),
                ("Cassete", (r"\bcassete",)),
                ("Radial", (r"\bradial",)),
            )),
            ("Vedacao & Borracha", "Borracha tecnica", (r"\bepdm\b", r"\bnbr\b", r"\bneoprene\b"), (
                ("EPDM", (r"\bepdm\b",)),
                ("NBR", (r"\bnbr\b",)),
                ("Neoprene", (r"\bneoprene\b",)),
            )),
            ("MRO & Manutencao", "Material eletrico", (r"\bdisjuntor\b", r"\bcontator\b", r"\bborne\b", r"\bcanaleta\b"), (
                ("Disjuntor", (r"\bdisjuntor",)),
                ("Contator", (r"\bcontator",)),
                ("Borne", (r"\bborne",)),
                ("Canaleta", (r"\bcanaleta",)),
            )),
            ("MRO & Manutencao", "Lubrificacao", (r"\bmassa\s+lubrificante\b", r"\boleo\s+lubrificante\b", r"\bspray\s+tecnico\b", r"\bdoseador\b"), (
                ("Massa", (r"\bmassa",)),
                ("Oleo", (r"\boleo",)),
                ("Spray tecnico", (r"\bspray",)),
                ("Doseador", (r"\bdoseador",)),
            )),
            ("Escritorio & Papelaria", "Cadernos e blocos", (r"\bcaderno\b", r"\bbloco\s+de\s+notas\b", r"\bagenda\b", r"\blivro\s+de\s+registo\b"), (
                ("Bloco de notas", (r"\bbloco\s+de\s+notas\b",)),
                ("Agenda", (r"\bagenda\b",)),
                ("Livro de registo", (r"\blivro\s+de\s+registo\b",)),
                ("Caderno", (r"\bcaderno\b",)),
            )),
            ("Escritorio & Papelaria", "Papel e etiquetas", (r"\bpapel\s+a[34]\b", r"\bresma\b", r"\betiqueta\b"), (
                ("Papel A3", (r"\ba3\b",)),
                ("Papel A4", (r"\ba4\b", r"\bresma\b")),
                ("Etiquetas", (r"\betiqueta\b",)),
                ("Papel tecnico", (r"\bpapel\b",)),
            )),
            ("Escritorio & Papelaria", "Escrita e marcacao", (r"\bcaneta\b", r"\bmarcador\b", r"\blapis\b", r"\blapiseira\b", r"\bcorretor\b"), (
                ("Marcador", (r"\bmarcador\b",)),
                ("Lapiseira", (r"\blapiseira\b",)),
                ("Lapis", (r"\blapis\b",)),
                ("Corretor", (r"\bcorretor\b",)),
                ("Caneta", (r"\bcaneta\b",)),
            )),
            ("Escritorio & Papelaria", "Arquivo e organizacao", (r"\bdossier\b", r"\bpasta\s+de\s+arquivo\b", r"\bseparador\b", r"\bcaixa\s+de\s+arquivo\b"), (
                ("Dossier", (r"\bdossier\b",)),
                ("Separador", (r"\bseparador\b",)),
                ("Caixa de arquivo", (r"\bcaixa\s+de\s+arquivo\b",)),
                ("Pasta", (r"\bpasta\b",)),
            )),
            ("Escritorio & Papelaria", "Consumiveis de impressao", (r"\btoner\b", r"\btinteiro\b", r"\btambor\s+de\s+impress", r"\bribbon\b"), (
                ("Toner", (r"\btoner\b",)),
                ("Tinteiro", (r"\btinteiro\b",)),
                ("Tambor", (r"\btambor\b",)),
                ("Ribbon", (r"\bribbon\b",)),
            )),
            ("Informatica", "Perifericos de computador", (r"\brato\s+(?:usb|sem\s+fios|wireless)\b", r"\bmouse\b", r"\bteclado\b", r"\bmonitor\b", r"\bwebcam\b", r"\bheadset\b"), (
                ("Rato", (r"\brato\b", r"\bmouse\b")),
                ("Teclado", (r"\bteclado\b",)),
                ("Monitor", (r"\bmonitor\b",)),
                ("Webcam", (r"\bwebcam\b",)),
                ("Headset", (r"\bheadset\b",)),
                ("Colunas", (r"\bcolunas?\b",)),
            )),
            ("Informatica", "Computadores", (r"\bcomputador\b", r"\bportatil\b", r"\blaptop\b", r"\bworkstation\b", r"\bmini\s*pc\b"), (
                ("Portatil", (r"\bportatil\b", r"\blaptop\b")),
                ("Workstation", (r"\bworkstation\b",)),
                ("Mini PC", (r"\bmini\s*pc\b",)),
                ("Desktop", (r"\bdesktop\b", r"\bcomputador\b")),
            )),
            ("Informatica", "Tablets e dispositivos moveis", (r"\bipad\b", r"\btablet\b", r"\be[\s-]?reader\b", r"\bkindle\b"), (
                ("iPad", (r"\bipad\b",)),
                ("Tablet Android", (r"\btablet\b.*\bandroid\b", r"\bandroid\b.*\btablet\b")),
                ("Tablet Windows", (r"\btablet\b.*\bwindows\b", r"\bwindows\b.*\btablet\b")),
                ("E-reader", (r"\be[\s-]?reader\b", r"\bkindle\b")),
            )),
            ("Informatica", "Redes e conectividade", (r"\bswitch\s+(?:de\s+)?rede\b", r"\brouter\b", r"\baccess\s*point\b", r"\bcabo\s+(?:de\s+)?rede\b"), (
                ("Switch", (r"\bswitch\b",)),
                ("Router", (r"\brouter\b",)),
                ("Access point", (r"\baccess\s*point\b",)),
                ("Cabo de rede", (r"\bcabo\b",)),
                ("Adaptador", (r"\badaptador\b",)),
            )),
            ("Informatica", "Armazenamento", (r"\bssd\b", r"\bdisco\s+rigido\b", r"\bpen\s*usb\b", r"\bcartao\s+de\s+memoria\b"), (
                ("SSD", (r"\bssd\b",)),
                ("Disco rigido", (r"\bdisco\s+rigido\b",)),
                ("Pen USB", (r"\bpen\s*usb\b",)),
                ("Cartao de memoria", (r"\bcartao\b",)),
            )),
        ]
        for category, subcategory, family_patterns, type_rules in catalog_rules:
            if not contains(*family_patterns):
                continue
            inferred_type = ""
            for type_label, type_patterns in type_rules:
                if contains(*type_patterns):
                    inferred_type = type_label
                    break
            return result(category, subcategory, inferred_type)

        if contains(r"\btinta\b", r"\besmalte\b", r"\bprimario\b", r"\bverniz\b", r"\bspray\b"):
            coating_type = (
                "Esmalte" if "esmalte" in text else
                "Primario" if "primario" in text else
                "Verniz" if "verniz" in text else
                "Spray" if "spray" in text else
                "Tinta tecnica"
            )
            return result("Tintas / Quimicos", "Tintas e revestimentos", coating_type)
        if contains(r"\bdiluente\b", r"\bsolvente\b", r"\bacetona\b", r"\bdesengordurante\b"):
            solvent_type = (
                "Diluente" if "diluente" in text else
                "Acetona" if "acetona" in text else
                "Desengordurante" if "desengordurante" in text else
                "Limpeza"
            )
            return result("Tintas / Quimicos", "Solventes e diluentes", solvent_type)
        if contains(r"\bsilicone\b", r"\bcola\b", r"\bvedante\b", r"\btrava\s*roscas\b"):
            adhesive_type = (
                "Silicone" if "silicone" in text else
                "Trava roscas" if contains(r"\btrava\s*roscas\b") else
                "Vedante" if "vedante" in text else
                "Cola"
            )
            return result("Tintas / Quimicos", "Adesivos e selantes", adhesive_type)
        if contains(r"\bzincado\b", r"\bdecapante\b", r"\bpassivante\b", r"\banticorrosiv"):
            treatment_type = (
                "Zincado" if contains(r"\bzincado\b") else
                "Decapante" if contains(r"\bdecapante\b") else
                "Passivante" if contains(r"\bpassivante\b") else
                "Anticorrosivo"
            )
            return result("Tintas / Quimicos", "Tratamento de superficie", treatment_type)

        if contains(r"\brolamento\b"):
            bearing_type = (
                "Agulhas" if "agulh" in text else
                "Rolos" if "rolo" in text else
                "Flange" if "flange" in text else
                "Esferas"
            )
            return result("Rolamentos & Transmissao", "Rolamentos", bearing_type)
        if contains(r"\bcorrente\b"):
            chain_type = "Dupla" if "dupla" in text else "Inox" if "inox" in text else "Simples"
            return result("Rolamentos & Transmissao", "Correntes", chain_type)
        if contains(r"\bpinhao\b", r"\bpolia\b", r"\bcorreia\b"):
            transmission_type = "Pinhao" if "pinhao" in text else "Polia" if "polia" in text else "Correia"
            return result("Rolamentos & Transmissao", "Pinhoes e polias", transmission_type)

        # Raw material shapes only apply when no finished-product family above
        # matched the description.
        material_category = ""
        if contains(r"\binox\b", r"\ba(?:isi)?\s*30[34]\b"):
            material_category = "Inox"
        elif contains(r"\baluminio\b"):
            material_category = "Aluminio"
        elif contains(r"\bferro\b", r"\bs235", r"\bs275", r"\bs355", r"\baco\b"):
            material_category = "Ferro"
        if material_category:
            shape = ""
            if contains(r"\bchapa\b"):
                shape = "Chapa"
            elif contains(r"\btubo\b"):
                shape = "Tubo"
            elif contains(r"\bperfil\b", r"\bcantoneira\b", r"\bupn\b", r"\bipe\b", r"\bhea\b", r"\bheb\b"):
                shape = "Perfil"
            elif contains(r"\bvarao\b"):
                shape = "Varao"
            if shape:
                return result(material_category, shape, "", "Material e formato identificados na descrição")

        # Technical plastics are evaluated last so words such as "nylon" in a
        # lock nut do not override the finished-product family.
        plastic_material = ""
        if contains(r"\bpead\b"):
            plastic_material = "PEAD"
        elif contains(r"\bpvc\b"):
            plastic_material = "PVC"
        elif contains(r"\bptfe\b", r"\bteflon\b"):
            plastic_material = "PTFE"
        elif contains(r"\bpolicarbonato\b"):
            plastic_material = "Policarbonato"
        elif contains(r"\bpom\b"):
            plastic_material = "POM"
        elif contains(r"\bpeek\b"):
            plastic_material = "PEEK"
        elif contains(r"\bnylon\b"):
            plastic_material = "Nylon"
        if plastic_material:
            plastic_shape = (
                "Chapa" if contains(r"\bchapa\b", r"\bplaca\b") else
                "Tubo" if contains(r"\btubo\b") else
                "Varao" if contains(r"\bvarao\b", r"\bbarra\b") else ""
            )
            if plastic_shape:
                return result("Plasticos Tecnicos", plastic_shape, plastic_material)

        # Unknown objects are not silently discarded. The Copilot creates a
        # conservative provisional entry that the user can correct and teach.
        keyword = self._product_learning_keyword(description)
        if keyword:
            provisional_type = keyword.replace("-", " ").title()
            return {
                "categoria": "Outros",
                "subcat": "Outros",
                "tipo": provisional_type,
                "dimensoes": dimensions,
                "confidence": 0.42,
                "reason": f"Termo novo detetado: “{keyword}”",
                "learned": False,
                "needs_learning": True,
                "learning_keyword": keyword,
            }
        return empty

    def product_copilot_analysis(self, description: str, current_code: str = "") -> dict[str, Any]:
        """Combine classification, attribute extraction and duplicate checks."""

        clean_description = str(description or "").strip()
        suggestion = dict(self.product_catalog_suggestion(clean_description) or {})
        attributes = _extract_product_attributes(clean_description)
        normalized_description = _normalize_product_description(clean_description)
        if str(suggestion.get("categoria", "") or "") == "Movimentacao":
            caster_type = str(suggestion.get("tipo", "") or "")
            caster_label = {
                "Giratorio": "Rodízio giratório",
                "Giratorio com travao total": "Rodízio giratório com travão total",
                "Fixo": "Rodízio fixo",
                "Roda avulsa": "Roda industrial",
            }.get(caster_type, "Rodízio industrial")
            parts = [caster_label]
            if attributes.get("fabricante"):
                parts.append(str(attributes["fabricante"]))
            if attributes.get("modelo"):
                parts.append(str(attributes["modelo"]))
            if attributes.get("medida"):
                parts.append(str(attributes["medida"]))
            if attributes.get("cor"):
                parts.append(str(attributes["cor"]))
            normalized_description = " – ".join(parts)
        current_key = str(current_code or "").strip().casefold()
        similar: list[dict[str, Any]] = []
        for product in list(self.ensure_data().get("produtos", []) or []):
            code = str(product.get("codigo", "") or "").strip()
            candidate = str(product.get("descricao", "") or "").strip()
            if not candidate or (current_key and code.casefold() == current_key):
                continue
            score = _product_similarity(clean_description, candidate)
            if score < 0.68:
                continue
            similar.append(
                {
                    "codigo": code,
                    "descricao": candidate,
                    "score": score,
                    "percent": int(round(score * 100)),
                }
            )
        similar.sort(key=lambda row: (-float(row.get("score", 0)), str(row.get("codigo", ""))))
        confidence = float(suggestion.get("confidence", 0) or 0)
        return {
            **suggestion,
            "descricao_normalizada": normalized_description,
            "atributos": attributes,
            "semelhantes": similar[:3],
            "duplicate_warning": bool(similar and float(similar[0].get("score", 0)) >= 0.84),
            "confidence_percent": int(round(confidence * 100)),
            "fabricante": str(attributes.get("fabricante", "") or ""),
            "modelo": str(attributes.get("modelo", "") or ""),
            "engine": "local-rules-v2",
        }

    def product_ai_lookup(
        self,
        description: str,
        *,
        user_instruction: str = "",
        previous_candidate: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """Ask the optional LuGEST gateway for a source-backed identification."""

        cfg = self._load_qt_config()
        endpoint = str(
            os.getenv("LUGEST_AI_ENDPOINT")
            or cfg.get("product_ai_endpoint")
            or ""
        ).strip()
        access_token = str(
            os.getenv("LUGEST_AI_ACCESS_TOKEN")
            or cfg.get("product_ai_access_token")
            or ""
        ).strip()
        gemini_api_key = str(
            os.getenv("LUGEST_GEMINI_API_KEY")
            or os.getenv("GEMINI_API_KEY")
            or ""
        ).strip()
        gemini_model = str(
            os.getenv("LUGEST_GEMINI_MODEL")
            or cfg.get("product_ai_gemini_model")
            or "gemini-3.6-flash"
        ).strip()
        ollama_url = str(
            os.getenv("LUGEST_OLLAMA_URL")
            or cfg.get("product_ai_ollama_url")
            or "http://127.0.0.1:11434"
        ).strip()
        ollama_model = str(
            os.getenv("LUGEST_OLLAMA_MODEL")
            or cfg.get("product_ai_ollama_model")
            or "qwen3:4b"
        ).strip()
        try:
            ai_timeout = float(
                os.getenv("LUGEST_AI_TIMEOUT_SECONDS")
                or cfg.get("product_ai_timeout_seconds")
                or 45
            )
        except (TypeError, ValueError):
            ai_timeout = 45.0
        local_candidate = self.product_copilot_analysis(description)
        taxonomy = self.product_taxonomy()
        compact_taxonomy = {
            "categories": [
                {
                    "label": row.get("label", ""),
                    "subcategories": [
                        {
                            "label": sub.get("label", ""),
                            "types": [
                                (
                                    item.get("label", "")
                                    if isinstance(item, dict)
                                    else str(item or "")
                                )
                                for item in list(sub.get("types", []) or [])
                            ],
                        }
                        for sub in list(row.get("subcategories", []) or [])
                    ],
                }
                for row in list(taxonomy.get("categories", []) or [])
            ]
        }
        client = _RemoteProductAIClient(
            endpoint,
            access_token,
            gemini_api_key=gemini_api_key,
            gemini_model=gemini_model,
            ollama_url=ollama_url,
            ollama_model=ollama_model,
            timeout_seconds=ai_timeout,
        )
        result = client.lookup(
            description,
            local_candidate=local_candidate,
            taxonomy=compact_taxonomy,
            locale="pt-PT",
            user_instruction=user_instruction,
            previous_candidate=previous_candidate,
        )
        result["local_candidate"] = local_candidate
        return result

    def _product_resolve_catalog_fields(self, payload: dict[str, Any]) -> dict[str, str]:
        category_map, subcategory_map, type_map = self._product_taxonomy_nodes()
        raw_category = str(payload.get("categoria", payload.get("category", "")) or "").strip()
        raw_subcategory = str(payload.get("subcat", payload.get("subcategory", "")) or "").strip()
        raw_type = str(payload.get("tipo", payload.get("type", "")) or "").strip()

        category_row = category_map.get(raw_category.casefold()) if raw_category else None
        category_label = str((category_row or {}).get("label", raw_category) or "").strip()
        category_id = str((category_row or {}).get("id", self._product_catalog_slug(category_label, "categoria")) or "").strip()

        subcategory_row = subcategory_map.get((category_label.casefold(), raw_subcategory.casefold())) if category_label and raw_subcategory else None
        subcategory_label = str((subcategory_row or {}).get("label", raw_subcategory) or "").strip()
        subcategory_id = (
            str((subcategory_row or {}).get("id", self._product_catalog_slug(f"{category_id}-{subcategory_label}", "subcategoria")) or "").strip()
            if subcategory_label
            else ""
        )

        type_row = type_map.get((category_label.casefold(), subcategory_label.casefold(), raw_type.casefold())) if category_label and subcategory_label and raw_type else None
        type_label = str((type_row or {}).get("label", raw_type) or "").strip()
        type_id = self._product_catalog_slug(f"{subcategory_id or category_id}-{type_label}", "tipo") if type_label else ""
        category_meta = dict(category_row or {})
        return {
            "categoria": category_label,
            "category_id": category_id,
            "subcat": subcategory_label,
            "subcategory_id": subcategory_id,
            "tipo": type_label,
            "type_id": type_id,
            "category_icon": str(category_meta.get("icon", "") or "").strip(),
            "category_badge": str(category_meta.get("badge", "") or "").strip(),
            "category_tone": str(category_meta.get("tone", "") or "").strip(),
        }
