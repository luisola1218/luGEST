"""Reprice owned catalog copies before a purchase note is committed."""
from dataclasses import dataclass
from typing import Any, Callable

@dataclass(frozen=True)
class PurchasePriceRules:
    parse_float: Callable
    is_material: Callable
    product_unit_price: Callable
    material_format: Callable
    now_iso: Callable
    product_price_mode: Callable
    sync_materials: Callable
    refresh_assemblies: Callable

class PurchasePricing:
    def __init__(self, rules: PurchasePriceRules):
        self.rules = rules

    def apply(self, lines, catalogs):
        material_changed = product_changed = False
        for line in lines:
            if self.rules.is_material(line.get("origem", "")):
                material_changed = self.update_material(catalogs.materials, line.get("ref", ""), line.get("preco", 0)) or material_changed
            else:
                product_changed = self.update_product(catalogs.products, line.get("ref", ""), line.get("preco", 0)) or product_changed
        if material_changed or product_changed:
            catalogs.assemblies = self.rules.refresh_assemblies(catalogs.assemblies, catalogs.products, catalogs.materials)
        for note in catalogs.notes:
            if material_changed:
                self.rules.sync_materials(note, catalogs.materials)
            if product_changed and self.sync_products(catalogs.products, note):
                note["total"] = round(sum(self.rules.parse_float(line.get("total", 0), 0) for line in note.get("linhas", [])), 2)

    def sync_products(self, products, note: dict[str, Any]) -> bool:
        changed = False
        product_map = {str(row.get("codigo", "") or "").strip(): row for row in products}
        for line in list(note.get("linhas", []) or []):
            if self.rules.is_material(line.get("origem", "Produto")):
                continue
            product = product_map.get(str(line.get("ref", "") or "").strip())
            if not product:
                continue
            new_price = round(self.rules.parse_float(self.rules.product_unit_price(product), 0), 6)
            old_price = self.rules.parse_float(line.get("preco", 0), 0)
            if abs(new_price - old_price) > 1e-9:
                line["preco"] = new_price
                qty = self.rules.parse_float(line.get("qtd", 0), 0)
                discount = max(0.0, min(100.0, self.rules.parse_float(line.get("desconto", 0), 0)))
                iva = max(0.0, min(100.0, self.rules.parse_float(line.get("iva", 23), 23)))
                base = (qty * new_price) * (1.0 - (discount / 100.0))
                line["total"] = round(base + (base * iva / 100.0), 4)
                changed = True
            new_desc = str(product.get("descricao", "") or "").strip()
            if new_desc and str(line.get("descricao", "") or "").strip() != new_desc:
                line["descricao"] = new_desc
                changed = True
        return changed


    def update_material(self, materials, materia_id: str, preco_unit: Any) -> bool:
        material = next((row for row in materials if str(row.get("id", "") or "").strip() == str(materia_id or "").strip()), None)
        if material is None:
            return False
        price_line = self.rules.parse_float(preco_unit, 0)
        old = self.rules.parse_float(material.get("p_compra", 0), 0)
        new_value = old
        formato = str(material.get("formato") or self.rules.material_format(material) or "").strip()
        if formato == "Tubo":
            metros = self.rules.parse_float(material.get("metros", 0), 0)
            if metros > 0:
                new_value = round(price_line / metros, 6)
        elif formato in ("Chapa", "Perfil"):
            peso = self.rules.parse_float(material.get("peso_unid", 0), 0)
            if peso > 0:
                new_value = round(price_line / peso, 6)
        else:
            new_value = round(price_line, 6)
        if abs(new_value - old) <= 1e-9:
            return False
        material["p_compra"] = new_value
        material["atualizado_em"] = self.rules.now_iso()
        return True


    def update_product(self, products, produto_codigo: str, preco_unit: Any) -> bool:
        code = str(produto_codigo or "").strip()
        product = next((row for row in products if str(row.get("codigo", "") or "").strip() == code), None)
        if product is None:
            return False
        price_line = self.rules.parse_float(preco_unit, 0)
        old = self.rules.parse_float(product.get("p_compra", 0), 0)
        new_value = old
        modo = self.rules.product_price_mode(product.get("categoria", ""), product.get("tipo", ""))
        if modo == "peso":
            peso = self.rules.parse_float(product.get("peso_unid", 0), 0)
            if peso > 0:
                new_value = round(price_line / peso, 6)
        elif modo == "metros":
            metros = self.rules.parse_float(product.get("metros_unidade", product.get("metros", 0)), 0)
            if metros > 0:
                new_value = round(price_line / metros, 6)
        else:
            new_value = round(price_line, 6)
        if abs(new_value - old) <= 1e-9:
            return False
        product["p_compra"] = new_value
        product["atualizado_em"] = self.rules.now_iso()
        return True
