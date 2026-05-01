from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from pypdf import PdfReader


ROOT = Path(__file__).resolve().parents[1]
EON_ROOT = Path(
    "/Users/silviofornaro/Library/CloudStorage/"
    "GoogleDrive-bolletteillumia.banco@gmail.com/Il mio Drive/E.ON"
)
OUTPUT = ROOT / "estrazioni_tariffe" / "eon_tariffe_2026-04.xlsx"


@dataclass(frozen=True)
class PdfSource:
    path: Path
    segment: str
    commodity: str
    offer_name: str
    offer_type: str
    variant: str = ""


SOURCES = [
    PdfSource(EON_ROOT / "CTE GAS/E.ONFlexGas_6P_D_FLX_CLSB.pdf", "RESIDENZIALE", "GAS", "E.ON Flex Gas", "VARIABILE"),
    PdfSource(EON_ROOT / "CTE GAS/E.ONFlexGasProtetta_6P_D_P12_CLSE.pdf", "RESIDENZIALE", "GAS", "E.ON Flex Gas Protetta", "VARIABILE"),
    PdfSource(EON_ROOT / "CTE GAS/E.ONGas50FlexPer12_6P_D_FLX_50_CLSB.pdf", "RESIDENZIALE", "GAS", "E.ON Gas 50 Flex Per12", "VARIABILE"),
    PdfSource(EON_ROOT / "CTE GAS/E.ONGas50TuaPer12_6P_D_50_V1_CLSA (2).pdf", "RESIDENZIALE", "GAS", "E.ON Gas 50 TuaPer12", "FISSA"),
    PdfSource(EON_ROOT / "CTE GAS/E.ONGasTua_6P_D_V1_CLSA (2).pdf", "RESIDENZIALE", "GAS", "E.ON Gas Tua", "FISSA"),
    PdfSource(EON_ROOT / "CTE GAS/E.ONGasVerdeProtetta_6P_D_P12_V1_CLSE.pdf", "RESIDENZIALE", "GAS", "E.ON GasVerde Protetta", "FISSA"),
    PdfSource(EON_ROOT / "CTE LUCE/E.ONFlexLuce_6P_D_FLX_CLSC.pdf", "RESIDENZIALE", "EE", "E.ON Flex Luce", "VARIABILE"),
    PdfSource(EON_ROOT / "CTE LUCE/E.ONFlexLuceCasa_6P_D_ACR_V1_CLSE.pdf", "RESIDENZIALE", "EE", "E.ON Flex Luce Casa", "VARIABILE"),
    PdfSource(EON_ROOT / "CTE LUCE/E.ONFlexLucePer24_6P_D_RI_CLSD.pdf", "RESIDENZIALE", "EE", "E.ON Flex Luce Per24", "VARIABILE"),
    PdfSource(EON_ROOT / "CTE LUCE/E.ONLuce50FlexPer12_6P_D_FLX_50_CLSC.pdf", "RESIDENZIALE", "EE", "E.ON Luce 50 Flex Per12", "VARIABILE"),
    PdfSource(EON_ROOT / "CTE LUCE/E.ONLuce50FlexPer24_6P_D_RI_50_CLSD.pdf", "RESIDENZIALE", "EE", "E.ON Luce 50 Flex Per24", "VARIABILE"),
    PdfSource(EON_ROOT / "CTE LUCE/E.ONLuce50TuaPer12_6P_D_50_V1_CLSA (1).pdf", "RESIDENZIALE", "EE", "E.ON Luce 50 TuaPer12", "FISSA"),
    PdfSource(EON_ROOT / "CTE LUCE/E.ONLuce50TuaPer24_6P_D_50_V1_CLSC (2).pdf", "RESIDENZIALE", "EE", "E.ON Luce 50 TuaPer24", "FISSA"),
    PdfSource(EON_ROOT / "CTE LUCE/E.ONLuceTuaPer24_6P_D_V1_CLSC (3).pdf", "RESIDENZIALE", "EE", "E.ON Luce Tua Per 24", "FISSA"),
    PdfSource(EON_ROOT / "CTE LUCE/E.ONLuceVerdeCasa_6P_D_ACR_V1_CLSE.pdf", "RESIDENZIALE", "EE", "E.ON LuceVerde Casa", "FISSA"),
    PdfSource(EON_ROOT / "CTE Microbusiness/E.ON Gas Impresa_6P_M_CLSC.pdf", "BUSINESS", "GAS", "E.ON Gas Impresa CLSC", "VARIABILE", "CLSC"),
    PdfSource(EON_ROOT / "CTE Microbusiness/E.ON Gas Impresa_6P_M_CLSE.pdf", "BUSINESS", "GAS", "E.ON Gas Impresa CLSE", "VARIABILE", "CLSE"),
    PdfSource(EON_ROOT / "CTE Microbusiness/E.ON LuceDinamica ECO_6P_M_CLSC.pdf", "BUSINESS", "EE", "E.ON LuceDinamica ECO CLSC", "VARIABILE", "CLSC"),
    PdfSource(EON_ROOT / "CTE Microbusiness/E.ON LuceDinamica ECO_6P_M_CLSE.pdf", "BUSINESS", "EE", "E.ON LuceDinamica ECO CLSE", "VARIABILE", "CLSE"),
]


def read_pdf_text(path: Path) -> str:
    reader = PdfReader(str(path))
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def number(value: str | None) -> float | None:
    if value is None:
        return None
    return float(value.replace(".", "").replace(",", "."))


def date_it(value: str | None) -> str:
    if not value:
        return ""
    return datetime.strptime(value, "%d/%m/%Y").date().isoformat()


def first(patterns: list[str], text: str, flags: int = re.IGNORECASE | re.DOTALL) -> str | None:
    for pattern in patterns:
        match = re.search(pattern, text, flags)
        if match:
            return match.group(1)
    return None


def extract_common(text: str) -> dict[str, Any]:
    return {
        "codice_offerta": first([r"Codice Offerta\s+([0-9A-Z]+)"], text) or "",
        "valid_to": date_it(first([r"Scadenza\s+(\d{2}/\d{2}/\d{4})"], text)),
        "valid_from": date_it(
            first(
                [
                    r"alla data del\s+(\d{2}/\d{2}/\d{4})",
                    r"alla data\s+(\d{2}/\d{2}/\d{4})",
                ],
                text,
            )
        ),
        "bonus": number(
            first(
                [
                    r"Bonus(?:\s+in\s+bolletta)?\s+di\s+([0-9]+,[0-9]+)\s*€",
                    r"Bonus\s+([0-9]+,[0-9]+)\s*€",
                ],
                text,
            )
        ),
    }


def extract_energy_values(src: PdfSource, text: str) -> dict[str, Any]:
    values: dict[str, Any] = {}
    if src.commodity == "GAS":
        values["omega"] = number(
            first(
                [
                    r"(?:Componente Omega|Ω|OMEGA).{0,220}?pari a\s+([0-9]+,[0-9]+)\s*€/Smc",
                    r"corrispettivo variabile.{0,220}?pari a\s+([0-9]+,[0-9]+)\s*€/Smc",
                ],
                text,
            )
        )
        values["prezzo_fisso"] = number(
            first(
                [
                    r"somministrazione\s+è\s+pari\s+a\s+([0-9]+,[0-9]+)\s*€/Smc",
                    r"Materia Prima Gas\*.*?pari\s+a\s+([0-9]+,[0-9]+)\s*€/Smc",
                ],
                text,
            )
        )
        values["ccv_quota_fissa"] = number(
            first([r"quota fissa\s+è\s+pari\s+a\s+([0-9]+,[0-9]+)\s*€/P[Dd]R/anno"], text)
        )
        values["ccv_quota_variabile"] = number(
            first([r"quota variabile\s+è\s+pari\s+a\s+([0-9]+,[0-9]+)\s*€/Smc"], text)
        )
        values["gestione_energetica_fissa"] = number(
            first([r"gestione energetica\s+è\s+pari\s+a\s+([0-9]+(?:,[0-9]+)?)\s*€/P[Dd]R/anno"], text)
        )
        values["oneri_integrativi"] = number(
            first([r"oneri integrativi di vendita\s+è\s+pari\s+a\s+([0-9]+,[0-9]+)\s*€/Smc"], text)
        )
        return values

    values["omega"] = number(
        first(
            [
                r"Componente Omega\).*?pari a\s+([0-9]+,[0-9]+)\s*€/kWh",
                r"parametro .omega.*?pari a\s+([0-9]+,[0-9]+)\s*€/kWh",
                r"corrispettivo variabile.*?pari a\s+([0-9]+,[0-9]+)\s*€/kWh",
            ],
            text,
        )
    )
    values["prezzo_mono"] = number(
        first(
            [
                r"prezzo monorario.*?pari a\s+([0-9]+,[0-9]+)\s*€/kWh",
                r"Opzione Monoraria:\s*([0-9]+,[0-9]+)\s*€/kWh",
            ],
            text,
        )
    )
    values["prezzo_f1"] = number(first([r"Prezzo Fascia F1:\s*([0-9]+,[0-9]+)\s*€/kWh"], text))
    values["prezzo_f23"] = number(first([r"Prezzo Fascia F2 e F3:\s*([0-9]+,[0-9]+)\s*€/kWh"], text))
    values["ccv_quota_fissa"] = number(
        first([r"commercializzazione e vendita\s+è\s+pari\s+a\s+([0-9]+,[0-9]+)\s*€/POD/anno"], text)
    )
    values["gestione_energetica_fissa"] = number(
        first([r"gestione energetica\s+è\s+pari\s+a\s+([0-9]+(?:,[0-9]+)?)\s*€/POD/anno"], text)
    )
    values["dispbt"] = number(first([r"DispBT.*?pari\s+a\s+([0-9]+,[0-9]+)\s*€/POD/anno"], text))
    return values


def row(src: PdfSource, common: dict[str, Any], component: str, uom: str, value: float | None, notes: str) -> dict[str, Any]:
    return {
        "provider": "E.ON",
        "segmento": src.segment,
        "commodity": src.commodity,
        "offer_type": src.offer_type,
        "offer_name": src.offer_name,
        "valid_from": common["valid_from"],
        "valid_to": common["valid_to"],
        "component": component,
        "uom": uom,
        "value": value,
        "index_name": "PUN" if src.commodity == "EE" and src.offer_type == "VARIABILE" else "PSV" if src.commodity == "GAS" and src.offer_type == "VARIABILE" else "",
        "notes": notes,
        "codice_offerta": common["codice_offerta"],
        "source_pdf": src.path.name,
    }


def build_tariff_rows() -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    tariff_rows: list[dict[str, Any]] = []
    summary_rows: list[dict[str, Any]] = []

    for src in SOURCES:
        if not src.path.exists():
            raise FileNotFoundError(src.path)
        text = read_pdf_text(src.path)
        common = extract_common(text)
        values = extract_energy_values(src, text)
        fixed_total = (values.get("ccv_quota_fissa") or 0.0) + (values.get("gestione_energetica_fissa") or 0.0)

        if src.commodity == "GAS" and src.offer_type == "VARIABILE":
            tariff_rows.append(row(src, common, "fee_energia", "€/Smc", values.get("omega"), "Omega su PSV DA."))
            tariff_rows.append(row(src, common, "ccv_quota_fissa", "€/anno", fixed_total or None, "Quota fissa annua; per business include anche gestione energetica fissa."))
            tariff_rows.append(row(src, common, "ccv_quota_variabile", "€/Smc", values.get("ccv_quota_variabile"), "Commercializzazione al dettaglio quota variabile."))
            if values.get("oneri_integrativi") is not None:
                tariff_rows.append(row(src, common, "bilanciamento", "€/Smc", values.get("oneri_integrativi"), "Oneri integrativi di vendita E.ON."))
        elif src.commodity == "GAS":
            tariff_rows.append(row(src, common, "prezzo_fisso", "€/Smc", values.get("prezzo_fisso"), "Materia Prima Gas fissa."))
            tariff_rows.append(row(src, common, "ccv_quota_fissa", "€/anno", fixed_total or None, "Quota fissa annua."))
            tariff_rows.append(row(src, common, "ccv_quota_variabile", "€/Smc", values.get("ccv_quota_variabile"), "Commercializzazione al dettaglio quota variabile."))
        elif src.commodity == "EE" and src.offer_type == "VARIABILE":
            tariff_rows.append(row(src, common, "fee_energia", "€/kWh", values.get("omega"), "Omega su PUN. PUN da considerare con perdite di rete 10%."))
            tariff_rows.append(row(src, common, "perdite_rete", "%", 0.10, "Perdite rete bassa tensione: fattore 1,100."))
            tariff_rows.append(row(src, common, "ccv_quota_fissa", "€/anno", fixed_total or None, "Quota fissa annua; per business include anche gestione energetica fissa."))
            if values.get("dispbt") is not None:
                tariff_rows.append(row(src, common, "dispbt", "€/anno", values.get("dispbt"), "Componente DispBT."))
        else:
            tariff_rows.append(row(src, common, "prezzo_mono", "€/kWh", values.get("prezzo_mono"), "Prezzo monorario."))
            if values.get("prezzo_f1") is not None:
                tariff_rows.append(row(src, common, "prezzo_f1", "€/kWh", values.get("prezzo_f1"), "Prezzo fascia F1."))
            if values.get("prezzo_f23") is not None:
                tariff_rows.append(row(src, common, "prezzo_f23", "€/kWh", values.get("prezzo_f23"), "Prezzo fasce F2/F3."))
            tariff_rows.append(row(src, common, "ccv_quota_fissa", "€/anno", fixed_total or None, "Quota fissa annua."))
            if values.get("dispbt") is not None:
                tariff_rows.append(row(src, common, "dispbt", "€/anno", values.get("dispbt"), "Componente DispBT."))

        if common.get("bonus") is not None:
            tariff_rows.append(row(src, common, "sconto_bonus", "€ una tantum", -abs(common["bonus"]), "Bonus commerciale E.ON, distinto dal Bonus Sociale statale."))

        summary_rows.append(
            {
                "provider": "E.ON",
                "segmento": src.segment,
                "commodity": src.commodity,
                "offer_type": src.offer_type,
                "offer_name": src.offer_name,
                "valid_from": common["valid_from"],
                "valid_to": common["valid_to"],
                "codice_offerta": common["codice_offerta"],
                "omega": values.get("omega"),
                "prezzo_fisso": values.get("prezzo_fisso"),
                "prezzo_mono": values.get("prezzo_mono"),
                "prezzo_f1": values.get("prezzo_f1"),
                "prezzo_f23": values.get("prezzo_f23"),
                "ccv_quota_fissa_totale_annua": fixed_total or None,
                "ccv_quota_variabile": values.get("ccv_quota_variabile"),
                "dispbt": values.get("dispbt"),
                "oneri_integrativi": values.get("oneri_integrativi"),
                "bonus_commerciale": -abs(common["bonus"]) if common.get("bonus") is not None else None,
                "source_pdf": src.path.name,
            }
        )
    return tariff_rows, summary_rows


def write_sheet(ws, rows: list[dict[str, Any]], headers: list[str]) -> None:
    ws.append(headers)
    for item in rows:
        ws.append([item.get(h, "") for h in headers])
    header_fill = PatternFill("solid", fgColor="DDEBF7")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for idx, header in enumerate(headers, start=1):
        width = max(len(str(header)) + 2, 12)
        for value in [ws.cell(row=r, column=idx).value for r in range(2, min(ws.max_row, 80) + 1)]:
            width = min(max(width, len(str(value)) + 2 if value is not None else width), 48)
        ws.column_dimensions[get_column_letter(idx)].width = width


def main() -> None:
    tariff_rows, summary_rows = build_tariff_rows()
    wb = Workbook()
    ws_tariffe = wb.active
    ws_tariffe.title = "tariffe"
    tariff_headers = [
        "provider",
        "segmento",
        "commodity",
        "offer_type",
        "offer_name",
        "valid_from",
        "valid_to",
        "component",
        "uom",
        "value",
        "index_name",
        "notes",
        "codice_offerta",
        "source_pdf",
    ]
    write_sheet(ws_tariffe, tariff_rows, tariff_headers)

    ws_summary = wb.create_sheet("riepilogo_offerte")
    summary_headers = [
        "provider",
        "segmento",
        "commodity",
        "offer_type",
        "offer_name",
        "valid_from",
        "valid_to",
        "codice_offerta",
        "omega",
        "prezzo_fisso",
        "prezzo_mono",
        "prezzo_f1",
        "prezzo_f23",
        "ccv_quota_fissa_totale_annua",
        "ccv_quota_variabile",
        "dispbt",
        "oneri_integrativi",
        "bonus_commerciale",
        "source_pdf",
    ]
    write_sheet(ws_summary, summary_rows, summary_headers)

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT)
    print(f"Creato: {OUTPUT}")
    print(f"Offerte: {len(summary_rows)}")
    print(f"Righe tariffarie: {len(tariff_rows)}")


if __name__ == "__main__":
    main()
