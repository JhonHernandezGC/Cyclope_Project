from __future__ import annotations
import argparse
import json
import re
import unicodedata
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any

import pandas as pd
import pdfplumber


# ============================================================================
# CAMBIAR RUTAS AQUI
# ============================================================================
# 1) PDF_INPUT_PATH:
#    Ruta de donde el script va a tomar los PDFs.
#    Puede ser una carpeta con varios .pdf o un PDF individual.
#
# 2) OUTPUT_DIR:
#    Ruta donde el script va a guardar los resultados:
#    - camara_extraction.json
#    - camara_extraction.xlsx
#
# Cuando muevas este archivo al portatil corporativo, reemplaza estas dos rutas
# por las carpetas reales de ese equipo.
# ============================================================================
PDF_INPUT_PATH = r"C:\Users\jhern\OneDrive\Desktop\Claude Code\Cyclope_Project\Camara_PDF"
OUTPUT_DIR = r"C:\Users\jhern\OneDrive\Desktop\Claude Code\Cyclope_Project\Output"


DOC_ID_RE = re.compile(r"\b(C\.?\s*C\.?|C\.?\s*E\.?|P\.?\s*P\.?)\s*No\.?\s*([A-Z0-9.\-]+)", re.I)

STOP_HEADINGS = {
    "CAMARA DE COMERCIO DE BOGOTA",
    "SEDE VIRTUAL",
    "CERTIFICADO DE EXISTENCIA Y REPRESENTACION LEGAL",
    "CERTIFICADO DE INSCRIPCION DE DOCUMENTOS",
    "FECHA EXPEDICION",
    "RECIBO NO",
    "VALOR",
    "CODIGO DE VERIFICACION",
    "VERIFIQUE EL CONTENIDO",
    "PAGINA",
    "NOMBRAMIENTOS",
    "ORGANO DE ADMINISTRACION",
    "JUNTA DIRECTIVA",
    "REPRESENTANTES LEGALES",
    "REFORMAS DE ESTATUTOS",
    "REVISORES FISCALES",
    "PODERES",
    "DOCUMENTO INSCRIPCION",
    "CAPITAL",
    "MATRICULA",
    "UBICACION",
}


@dataclass
class Appointment:
    section: str
    group: str
    cargo: str
    nombre: str
    identificacion_tipo: str
    identificacion_numero: str
    page: int


def strip_accents(value: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", value) if not unicodedata.combining(c)
    )


def norm(value: str) -> str:
    value = strip_accents(value).upper()
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def clean_space(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def extract_pages(pdf_path: Path) -> list[dict[str, Any]]:
    pages: list[dict[str, Any]] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for index, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            words = page.extract_words(x_tolerance=1, y_tolerance=3, keep_blank_chars=False)
            pages.append({"page": index, "text": text, "words": words})
    return pages


def extract_basic_info(full_text: str) -> dict[str, str]:
    patterns = {
        "razon_social": r"Raz[oó]n social:\s*(.+)",
        "nit": r"\bNit:\s*([0-9.\-\s]+)",
        "domicilio_principal": r"Domicilio principal:\s*(.+)",
        "direccion_domicilio_principal": r"Direcci[oó]n\s+del\s+domicilio\s+principal\s*:\s*(.*?)(?:\n|$)", #se cruza con linea 337
        "Teléfono comercial 1:": r"Tel[eé]fono comercial 1:\s*(.+)", # se cruza con linea 337
        "ciiu": r"C[oó]digo\s+CIIU\s*[:\-]?\s*(\d{4})"
    }
    result: dict[str, str] = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, full_text, flags=re.I)
        result[key] = clean_space(match.group(1)) if match else ""
    result["nit"] = re.sub(r"\s+", "-", result.get("nit", "")).strip("-")
    return result


def normalized_index(original: str) -> tuple[str, list[int]]:
    normalized_chars: list[str] = []
    index_map: list[int] = []
    for original_index, char in enumerate(original):
        decomposed = unicodedata.normalize("NFKD", char)
        for part in decomposed:
            if not unicodedata.combining(part):
                normalized_chars.append(part.upper())
                index_map.append(original_index)
    return "".join(normalized_chars), index_map


def section_between(full_text: str, start_terms: list[str], end_terms: list[str]) -> str:
    normalized, index_map = normalized_index(full_text)
    starts = [(normalized.find(term), term) for term in start_terms]
    starts = [(pos, term) for pos, term in starts if pos >= 0]
    if not starts:
        return ""
    start_pos, start_term = min(starts, key=lambda item: item[0])
    start_original = index_map[start_pos + len(start_term) - 1] + 1

    end_candidates: list[int] = []
    for term in end_terms:
        pos = normalized.find(term, start_pos + len(start_term))
        if pos >= 0:
            end_candidates.append(index_map[pos])
    end_original = min(end_candidates) if end_candidates else len(full_text)
    return clean_space(full_text[start_original:end_original])


def extract_faculties(full_text: str) -> dict[str, Any]:
    faculties_text = section_between(
        full_text,
        ["FACULTADES Y LIMITACIONES DEL REPRESENTANTE LEGAL", "FACULTADES Y LIMITACIONES"],
        ["NOMBRAMIENTOS", "REFORMAS DE ESTATUTOS", "RECURSOS", "CERTIFICA"],
    )
    limitations_area = norm(faculties_text if faculties_text else full_text)
    judicial_area = norm(full_text)

    limitation_patterns = [
        r"AUTORIZACION EXPRESA",
        r"CUANTIA",
        r"SUPERIOR A",
        r"NECESITAR[A-Z]* AUTORIZACION",
        r"LIMITACION",
        r"NO PODRA",
        r"JUNTA DE (MIEMBROS|SOCIOS|ACCIONISTAS|DIRECTIVA)",
    ]
    judicial_patterns = [
        r"REPRESENTANTES? JUDICIALES?",
        r"APODERAD[OA]S?",
        r"PODER GENERAL",
        r"ACTUACIONES JUDICIALES",
        r"AUDIENCIAS? DE CONCILIACION",
    ]

    limitations_found = sorted(
        {pattern for pattern in limitation_patterns if re.search(pattern, limitations_area)}
    )
    judicial_found = sorted(
        {pattern for pattern in judicial_patterns if re.search(pattern, judicial_area)}
    )
    return {
        "facultades_texto": faculties_text,
        "tiene_limitaciones": bool(limitations_found),
        "reglas_limitaciones_detectadas": limitations_found,
        "tiene_indicios_representacion_judicial": bool(judicial_found),
        "reglas_representacion_judicial_detectadas": judicial_found,
    }


def group_words_by_line(words: list[dict[str, Any]]) -> list[list[dict[str, Any]]]:
    lines: dict[int, list[dict[str, Any]]] = {}
    for word in words:
        top = round(float(word["top"]) / 3) * 3
        lines.setdefault(top, []).append(word)
    return [sorted(lines[top], key=lambda item: item["x0"]) for top in sorted(lines)]


def line_text(line: list[dict[str, Any]]) -> str:
    return clean_space(" ".join(str(word["text"]) for word in line))


def line_columns(line: list[dict[str, Any]]) -> tuple[str, str, str]:
    cargo_words: list[str] = []
    name_words: list[str] = []
    id_words: list[str] = []
    for word in line:
        x0 = float(word["x0"])
        text = str(word["text"])
        if x0 < 205:
            cargo_words.append(text)
        elif x0 < 370:
            name_words.append(text)
        else:
            id_words.append(text)
    return clean_space(" ".join(cargo_words)), clean_space(" ".join(name_words)), clean_space(" ".join(id_words))


def is_stop_line(raw: str) -> bool:
    normalized = norm(raw)
    if not normalized:
        return True
    if normalized.startswith("PAGINA "):
        return True
    if normalized.startswith(("POR ACTA", "POR DOCUMENTO", "POR ESCRITURA", "INSCRITA ")):
        return True
    return any(normalized.startswith(heading) for heading in STOP_HEADINGS)


def extract_appointments(pages: list[dict[str, Any]]) -> list[Appointment]:
    appointments: list[Appointment] = []
    active_table = False
    section = ""
    group = ""
    current: Appointment | None = None

    def flush_current() -> None:
        nonlocal current
        if current:
            current.cargo = clean_space(current.cargo)
            current.nombre = clean_space(current.nombre)
            appointments.append(current)
            current = None

    for page in pages:
        page_no = int(page["page"])
        for line in group_words_by_line(page["words"]):
            raw = line_text(line)
            normalized = norm(raw)

            if normalized in {"REPRESENTANTES LEGALES", "REPRESENTANTE LEGAL"}:
                flush_current()
                section = "REPRESENTANTES LEGALES"
                group = ""
                active_table = False
                continue
            if normalized in {"ORGANO DE ADMINISTRACION", "JUNTA DIRECTIVA"}:
                flush_current()
                section = "ORGANO DE ADMINISTRACION"
                active_table = False
                continue
            if normalized in {"REVISORES FISCALES", "PODERES"}:
                flush_current()
                section = ""
                group = ""
                active_table = False
                continue
            if normalized in {"PRINCIPALES", "SUPLENTES"}:
                flush_current()
                group = normalized
                active_table = False
                continue
            if "CARGO" in normalized and "NOMBRE" in normalized and "IDENTIFICACION" in normalized:
                flush_current()
                active_table = section in {"REPRESENTANTES LEGALES", "ORGANO DE ADMINISTRACION"}
                continue

            if not active_table:
                continue

            if is_stop_line(raw):
                flush_current()
                active_table = False
                continue

            cargo, name, ident = line_columns(line)
            id_match = DOC_ID_RE.search(ident)
            if id_match:
                flush_current()
                current = Appointment(
                    section=section or "NOMBRAMIENTOS",
                    group=group,
                    cargo=cargo,
                    nombre=name,
                    identificacion_tipo=clean_space(id_match.group(1).replace(" ", "")),
                    identificacion_numero=clean_space(id_match.group(2)),
                    page=page_no,
                )
                continue

            if current and not is_stop_line(raw):
                if cargo:
                    current.cargo = clean_space(current.cargo + " " + cargo)
                if name:
                    current.nombre = clean_space(current.nombre + " " + name)

        flush_current()
        active_table = False

    flush_current()
    return appointments


def extract_pdf(pdf_path: Path) -> dict[str, Any]:
    pages = extract_pages(pdf_path)
    full_text = "\n".join(page["text"] for page in pages)
    basic = extract_basic_info(full_text)
    appointments = extract_appointments(pages)
    return {
        "archivo": str(pdf_path),
        **basic,
        **extract_faculties(full_text),
        "nombramientos": [asdict(item) for item in appointments],
    }


def write_outputs(records: list[dict[str, Any]], output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "camara_extraction.json").write_text(
        json.dumps(records, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    summary_rows = []
    appointment_rows = []
    for record in records:
        summary_rows.append(
            {
                "archivo": Path(record["archivo"]).name,
                "razon_social": record["razon_social"],
                "nit": record["nit"],
                "domicilio_principal": record["domicilio_principal"],
                "direccion_domicilio_principal": record["direccion_domicilio_principal"],  # nueva columna localizada en linea 102
                "Teléfono comercial 1:": record["Teléfono comercial 1:"],  # nueva columna localizada en linea 103
                "Actividad principal Código CIIU": record.get("ciiu", ""),
                "tiene_limitaciones": record["tiene_limitaciones"],
                "tiene_indicios_representacion_judicial": record[
                    "tiene_indicios_representacion_judicial"
                ],
            }
        )
        for item in record["nombramientos"]:
            appointment_rows.append({"archivo": Path(record["archivo"]).name, **item})

    excel_path = output_dir / "camara_extraction.xlsx"
    try:
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Resumen", index=False)
            pd.DataFrame(appointment_rows).to_excel(writer, sheet_name="Nombramientos", index=False)
    except PermissionError as exc:
        raise SystemExit(
            f"No pude escribir el Excel porque parece estar abierto o bloqueado: {excel_path}. "
            "Cierra el archivo y vuelve a correr el script."
        ) from exc


def find_pdfs(inputs: list[str]) -> list[Path]:
    pdfs: list[Path] = []
    for raw in inputs:
        path = Path(raw)
        if path.is_dir():
            pdfs.extend(sorted(path.glob("*.pdf")))
        elif path.suffix.lower() == ".pdf":
            pdfs.append(path)
    return pdfs


def main() -> None:
    parser = argparse.ArgumentParser(description="Extrae datos clave de certificados CCB digitales.")
    parser.add_argument(
        "inputs",
        nargs="*",
        help=(
            "Uno o mas PDFs o carpetas con PDFs. "
            "Si no se informa, usa PDF_INPUT_PATH configurado al inicio del archivo."
        ),
    )
    parser.add_argument(
        "--output-dir",
        default=OUTPUT_DIR,
        help=(
            "Carpeta de salida para JSON y Excel. "
            "Si no se informa, usa OUTPUT_DIR configurado al inicio del archivo."
        ),
    )
    args = parser.parse_args()

    input_paths = args.inputs or [PDF_INPUT_PATH]
    pdfs = find_pdfs(input_paths)
    if not pdfs:
        raise SystemExit(
            "No se encontraron PDFs para procesar. "
            "Revisa la ruta PDF_INPUT_PATH al inicio del archivo."
        )

    records = [extract_pdf(pdf) for pdf in pdfs]
    write_outputs(records, Path(args.output_dir))

    for record in records:
        print(
            f"{Path(record['archivo']).name}: "
            f"{record['razon_social']} | NIT {record['nit']} | "
            f"{len(record['nombramientos'])} nombramientos"
        )
    print(f"Salida: {Path(args.output_dir).resolve()}")


if __name__ == "__main__":
    main()
