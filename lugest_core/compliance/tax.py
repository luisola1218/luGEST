from __future__ import annotations

import base64
import json
import os
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET

from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding, rsa


SAFT_PT_AUDIT_FILE_VERSION = "1.04_01"
DEFAULT_HASH_CONTROL = "1"
DEFAULT_PRODUCT_ID = "LuGEST"
DEFAULT_PRODUCT_VERSION = "2026.03"


def _decimal(value: Any, quant: str = "0.01") -> Decimal:
    try:
        dec = Decimal(str(value if value is not None else "0"))
    except Exception:
        dec = Decimal("0")
    return dec.quantize(Decimal(quant), rounding=ROUND_HALF_UP)


def _money(value: Any) -> str:
    return f"{_decimal(value):.2f}"


def _quantity(value: Any) -> str:
    dec = _decimal(value, "0.001")
    return format(dec.normalize(), "f") if dec != dec.to_integral() else str(dec.quantize(Decimal("1")))


def _clean_text(value: Any) -> str:
    return str(value or "").replace("\r", " ").replace("\n", " ").strip()


def _date_text(value: Any) -> str:
    raw = _clean_text(value).replace("T", " ")
    if not raw:
        return ""
    try:
        return datetime.fromisoformat(raw).strftime("%Y-%m-%d")
    except Exception:
        return raw[:10]


def _datetime_text(value: Any) -> str:
    raw = _clean_text(value).replace("T", " ")
    if not raw:
        return ""
    try:
        dt = datetime.fromisoformat(raw)
    except Exception:
        try:
            dt = datetime.strptime(raw[:19], "%Y-%m-%d %H:%M:%S")
        except Exception:
            if len(raw) >= 19:
                return raw[:19].replace(" ", "T")
            if len(raw) >= 10:
                return f"{raw[:10]}T00:00:00"
            return raw
    return dt.strftime("%Y-%m-%dT%H:%M:%S")


def legal_document_number(doc_type: Any, serie_id: Any, seq_num: Any, fallback_number: Any = "") -> str:
    doc = _clean_text(doc_type).upper() or "FT"
    serie = _clean_text(serie_id)
    try:
        seq = int(Decimal(str(seq_num or "0")))
    except Exception:
        seq = 0
    if serie and seq > 0:
        return f"{doc} {serie}/{seq}"
    return _clean_text(fallback_number) or doc


def normalize_source_billing(value: Any, fallback: str = "P") -> str:
    text = _clean_text(value).upper() or fallback.upper()
    if text in {"P", "I", "M"}:
        return text
    return fallback.upper() if fallback.upper() in {"P", "I", "M"} else "P"


def invoice_status_code(*, is_void: bool) -> str:
    return "A" if bool(is_void) else "N"


def tax_code_from_rate(rate: Any) -> str:
    return "NOR" if _decimal(rate) > 0 else "NS"


def product_type_from_line(line: dict[str, Any]) -> str:
    unit = _clean_text(line.get("unit")).upper()
    reference = _clean_text(line.get("reference")).upper()
    if unit == "SV" or reference.startswith("TRANSP") or reference.startswith("SERVICO"):
        return "S"
    return "P"


def build_invoice_hash_message(
    *,
    invoice_date: Any,
    system_entry_date: Any,
    invoice_no: Any,
    gross_total: Any,
    previous_hash: Any = "",
) -> str:
    return ";".join(
        [
            _date_text(invoice_date) or "0000-00-00",
            _datetime_text(system_entry_date) or "0000-00-00T00:00:00",
            _clean_text(invoice_no) or "-",
            _money(gross_total),
            _clean_text(previous_hash),
        ]
    )


def load_or_create_signing_material(
    base_dir: str | Path,
    *,
    private_key_path: str = "",
    public_key_path: str = "",
) -> dict[str, str]:
    base = Path(base_dir)
    private_hint = _clean_text(private_key_path) or _clean_text(os.getenv("LUGEST_FISCAL_PRIVATE_KEY"))
    public_hint = _clean_text(public_key_path) or _clean_text(os.getenv("LUGEST_FISCAL_PUBLIC_KEY"))
    default_dir = base / "generated" / "compliance" / "keys"
    default_dir.mkdir(parents=True, exist_ok=True)
    private_path = Path(private_hint) if private_hint else (default_dir / "lugest_fiscal_private.pem")
    public_path = Path(public_hint) if public_hint else (default_dir / "lugest_fiscal_public.pem")
    if not private_path.is_absolute():
        private_path = base / private_path
    if not public_path.is_absolute():
        public_path = base / public_path
    private_path.parent.mkdir(parents=True, exist_ok=True)
    public_path.parent.mkdir(parents=True, exist_ok=True)

    if not private_path.exists() or not public_path.exists():
        private_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
        public_key = private_key.public_key()
        private_path.write_bytes(
            private_key.private_bytes(
                encoding=serialization.Encoding.PEM,
                format=serialization.PrivateFormat.PKCS8,
                encryption_algorithm=serialization.NoEncryption(),
            )
        )
        public_path.write_bytes(
            public_key.public_bytes(
                encoding=serialization.Encoding.PEM,
                format=serialization.PublicFormat.SubjectPublicKeyInfo,
            )
        )

    return {
        "private_key_path": str(private_path),
        "public_key_path": str(public_path),
        "private_key_pem": private_path.read_text(encoding="utf-8"),
        "public_key_pem": public_path.read_text(encoding="utf-8"),
    }


def sign_message_pkcs1_sha1(message: str, private_key_pem: str) -> str:
    key = serialization.load_pem_private_key(private_key_pem.encode("utf-8"), password=None)
    signature = key.sign(
        message.encode("utf-8"),
        padding.PKCS1v15(),
        hashes.SHA1(),
    )
    return base64.b64encode(signature).decode("ascii")


def serialize_snapshot(document: dict[str, Any]) -> str:
    return json.dumps(document, ensure_ascii=False, sort_keys=True)


def deserialize_snapshot(value: Any) -> dict[str, Any]:
    raw = _clean_text(value)
    if not raw:
        return {}
    try:
        parsed = json.loads(raw)
    except Exception:
        return {}
    return dict(parsed or {}) if isinstance(parsed, dict) else {}


def _set_text(parent: ET.Element, tag: str, value: Any, *, allow_empty: bool = False) -> ET.Element | None:
    text = _clean_text(value)
    if not text and not allow_empty:
        return None
    node = ET.SubElement(parent, tag)
    node.text = text
    return node


def _write_xml(root: ET.Element, output_path: str | Path) -> Path:
    target = Path(output_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    tree = ET.ElementTree(root)
    try:
        ET.indent(tree, space="  ")
    except Exception:
        pass
    tree.write(target, encoding="utf-8", xml_declaration=True)
    return target


def render_saft_pt_xml(export_data: dict[str, Any], output_path: str | Path) -> Path:
    header = dict(export_data.get("header", {}) or {})
    customers = list(export_data.get("customers", []) or [])
    products = list(export_data.get("products", []) or [])
    tax_table = list(export_data.get("tax_table", []) or [])
    invoices = list(export_data.get("invoices", []) or [])

    root = ET.Element("AuditFile")
    _set_text(root, "AuditFileVersion", header.get("audit_file_version", SAFT_PT_AUDIT_FILE_VERSION))

    header_node = ET.SubElement(root, "Header")
    header_map = [
        ("CompanyID", header.get("company_id")),
        ("TaxRegistrationNumber", header.get("tax_registration_number")),
        ("TaxAccountingBasis", header.get("tax_accounting_basis", "F")),
        ("CompanyName", header.get("company_name")),
        ("BusinessName", header.get("business_name")),
        ("CompanyAddress", None),
        ("FiscalYear", header.get("fiscal_year")),
        ("StartDate", header.get("start_date")),
        ("EndDate", header.get("end_date")),
        ("CurrencyCode", header.get("currency_code", "EUR")),
        ("DateCreated", header.get("date_created")),
        ("TaxEntity", header.get("tax_entity", "Global")),
        ("ProductCompanyTaxID", header.get("product_company_tax_id")),
        ("SoftwareCertificateNumber", header.get("software_certificate_number", "0")),
        ("ProductID", header.get("product_id", DEFAULT_PRODUCT_ID)),
        ("ProductVersion", header.get("product_version", DEFAULT_PRODUCT_VERSION)),
        ("HeaderComment", header.get("header_comment", "Exportação interna LuGEST SAF-T(PT).")),
    ]
    for tag, value in header_map:
        if tag == "CompanyAddress":
            address_node = ET.SubElement(header_node, "CompanyAddress")
            _set_text(address_node, "AddressDetail", header.get("company_address_detail", "-"))
            _set_text(address_node, "City", header.get("company_city", "-"))
            _set_text(address_node, "PostalCode", header.get("company_postal_code", "0000-000"))
            _set_text(address_node, "Country", header.get("company_country", "PT"))
            continue
        _set_text(header_node, tag, value)

    master_node = ET.SubElement(root, "MasterFiles")
    for customer in customers:
        customer_node = ET.SubElement(master_node, "Customer")
        _set_text(customer_node, "CustomerID", customer.get("customer_id", "CONSUMIDOR-FINAL"))
        _set_text(customer_node, "AccountID", customer.get("account_id", customer.get("customer_id", "21")))
        _set_text(customer_node, "CustomerTaxID", customer.get("tax_id", "999999990"))
        _set_text(customer_node, "CompanyName", customer.get("name", "Cliente"))
        bill_node = ET.SubElement(customer_node, "BillingAddress")
        _set_text(bill_node, "AddressDetail", customer.get("address_detail", "-"))
        _set_text(bill_node, "City", customer.get("city", "-"))
        _set_text(bill_node, "PostalCode", customer.get("postal_code", "0000-000"))
        _set_text(bill_node, "Country", customer.get("country", "PT"))

    for product in products:
        product_node = ET.SubElement(master_node, "Product")
        _set_text(product_node, "ProductType", product.get("product_type", "P"))
        _set_text(product_node, "ProductCode", product.get("product_code", "-"))
        _set_text(product_node, "ProductGroup", product.get("product_group", "GERAL"))
        _set_text(product_node, "ProductDescription", product.get("product_description", "-"))
        _set_text(product_node, "ProductNumberCode", product.get("product_number_code", product.get("product_code", "-")))

    tax_node = ET.SubElement(master_node, "TaxTable")
    for tax in tax_table:
        entry = ET.SubElement(tax_node, "TaxTableEntry")
        _set_text(entry, "TaxType", tax.get("tax_type", "IVA"))
        _set_text(entry, "TaxCountryRegion", tax.get("tax_country_region", "PT"))
        _set_text(entry, "TaxCode", tax.get("tax_code", "NOR"))
        if _decimal(tax.get("tax_percentage", 0)) > 0:
            _set_text(entry, "Description", tax.get("description", "IVA"))
            _set_text(entry, "TaxPercentage", _money(tax.get("tax_percentage", 0)))
        else:
            _set_text(entry, "Description", tax.get("description", "Nao sujeito"))
            _set_text(entry, "TaxPercentage", "0.00")

    source_documents = ET.SubElement(root, "SourceDocuments")
    sales_node = ET.SubElement(source_documents, "SalesInvoices")
    _set_text(sales_node, "NumberOfEntries", str(len(invoices)))
    _set_text(sales_node, "TotalDebit", "0.00")
    _set_text(sales_node, "TotalCredit", _money(sum(_decimal(row.get("gross_total", 0)) for row in invoices)))

    for invoice in invoices:
        invoice_node = ET.SubElement(sales_node, "Invoice")
        _set_text(invoice_node, "InvoiceNo", invoice.get("invoice_no"))
        status_node = ET.SubElement(invoice_node, "DocumentStatus")
        _set_text(status_node, "InvoiceStatus", invoice.get("invoice_status", "N"))
        _set_text(status_node, "InvoiceStatusDate", invoice.get("invoice_status_date"))
        _set_text(status_node, "SourceID", invoice.get("status_source_id", invoice.get("source_id", "Sistema")))
        _set_text(status_node, "SourceBilling", invoice.get("source_billing", "P"))
        _set_text(invoice_node, "Hash", invoice.get("hash", "0"))
        _set_text(invoice_node, "HashControl", invoice.get("hash_control", "0"))
        _set_text(invoice_node, "Period", str(invoice.get("period", "")))
        _set_text(invoice_node, "InvoiceDate", invoice.get("invoice_date"))
        _set_text(invoice_node, "InvoiceType", invoice.get("invoice_type", "FT"))
        special_node = ET.SubElement(invoice_node, "SpecialRegimes")
        _set_text(special_node, "SelfBillingIndicator", "0")
        _set_text(special_node, "CashVATSchemeIndicator", "0")
        _set_text(special_node, "ThirdPartiesBillingIndicator", "0")
        _set_text(invoice_node, "SourceID", invoice.get("source_id", "Sistema"))
        _set_text(invoice_node, "SystemEntryDate", invoice.get("system_entry_date"))
        _set_text(invoice_node, "CustomerID", invoice.get("customer_id", "CONSUMIDOR-FINAL"))
        for idx, line in enumerate(list(invoice.get("lines", []) or []), start=1):
            line_node = ET.SubElement(invoice_node, "Line")
            _set_text(line_node, "LineNumber", str(idx))
            _set_text(line_node, "ProductCode", line.get("product_code", "-"))
            _set_text(line_node, "ProductDescription", line.get("product_description", "-"))
            _set_text(line_node, "Quantity", _quantity(line.get("quantity", 0)))
            _set_text(line_node, "UnitOfMeasure", line.get("unit_of_measure", "UN"))
            _set_text(line_node, "UnitPrice", _money(line.get("unit_price", 0)))
            _set_text(line_node, "TaxPointDate", invoice.get("invoice_date"))
            _set_text(line_node, "Description", line.get("description", "-"))
            _set_text(line_node, "CreditAmount", _money(line.get("credit_amount", 0)))
            tax_entry = ET.SubElement(line_node, "Tax")
            _set_text(tax_entry, "TaxType", line.get("tax_type", "IVA"))
            _set_text(tax_entry, "TaxCountryRegion", line.get("tax_country_region", "PT"))
            _set_text(tax_entry, "TaxCode", line.get("tax_code", tax_code_from_rate(line.get("tax_percentage", 0))))
            _set_text(tax_entry, "TaxPercentage", _money(line.get("tax_percentage", 0)))
            if _clean_text(line.get("tax_exemption_reason")):
                _set_text(line_node, "TaxExemptionReason", line.get("tax_exemption_reason"))
        totals_node = ET.SubElement(invoice_node, "DocumentTotals")
        _set_text(totals_node, "TaxPayable", _money(invoice.get("tax_payable", 0)))
        _set_text(totals_node, "NetTotal", _money(invoice.get("net_total", 0)))
        _set_text(totals_node, "GrossTotal", _money(invoice.get("gross_total", 0)))

    return _write_xml(root, output_path)


def render_at_communication_preparation_xml(batch_data: dict[str, Any], output_path: str | Path) -> Path:
    header = dict(batch_data.get("header", {}) or {})
    documents = list(batch_data.get("documents", []) or [])

    root = ET.Element("ATCommunicationPreparation")
    root.set("generatedAt", _datetime_text(header.get("generated_at")) or datetime.now().strftime("%Y-%m-%dT%H:%M:%S"))
    meta_node = ET.SubElement(root, "Header")
    _set_text(meta_node, "IssuerName", header.get("issuer_name", "Emitente"))
    _set_text(meta_node, "IssuerTaxID", header.get("issuer_tax_id", "999999990"))
    _set_text(meta_node, "SoftwareCertificateNumber", header.get("software_certificate_number", "0"))
    _set_text(meta_node, "ProductID", header.get("product_id", DEFAULT_PRODUCT_ID))
    _set_text(meta_node, "ProductVersion", header.get("product_version", DEFAULT_PRODUCT_VERSION))
    _set_text(meta_node, "PreparationMode", header.get("preparation_mode", "manual"))
    docs_node = ET.SubElement(root, "Documents")
    for row in documents:
        node = ET.SubElement(docs_node, "Document")
        _set_text(node, "DocumentID", row.get("document_id"))
        _set_text(node, "InvoiceNo", row.get("invoice_no"))
        _set_text(node, "InvoiceDate", row.get("invoice_date"))
        _set_text(node, "InvoiceType", row.get("invoice_type", "FT"))
        _set_text(node, "ATCUD", row.get("atcud"))
        _set_text(node, "Hash", row.get("hash", "0"))
        _set_text(node, "HashControl", row.get("hash_control", "0"))
        _set_text(node, "CustomerTaxID", row.get("customer_tax_id", "999999990"))
        _set_text(node, "GrossTotal", _money(row.get("gross_total", 0)))
        _set_text(node, "Status", row.get("status", "Pendente"))
        _set_text(node, "SourceBilling", row.get("source_billing", "P"))
        _set_text(node, "SystemEntryDate", row.get("system_entry_date"))
    return _write_xml(root, output_path)
