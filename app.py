from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import StreamingResponse
import pandas as pd
import io
import re

app = FastAPI()

def parse_export_block(text: str):
    sections = {}
    current_section = None
    in_export = False

    for line in text.splitlines():
        line = line.strip()

        if line == "[EXCEL_EXPORT]":
            in_export = True
            continue
        if line == "[/EXCEL_EXPORT]":
            break
        if not in_export:
            continue

        section_match = re.match(r"\[(.+?)\]", line)
        if section_match:
            current_section = section_match.group(1)
            sections[current_section] = []
            continue

        if "|" in line and current_section:
            parts = [col.strip() for col in line.split("|")]
            sections[current_section].append(parts)

    return sections


def append_rows(df, rows):
    for row in rows:
        if len(row) != len(df.columns):
            row = row[:len(df.columns)] + [""] * (len(df.columns) - len(row))
        df.loc[len(df)] = row
    return df


@app.post("/update_excel")
async def update_excel(
    export_block: str = Form(...),
    excel_file: UploadFile = File(...)
):

    content = await excel_file.read()
    sections = parse_export_block(export_block)

    file_stream = io.BytesIO(content)
    excel = pd.ExcelFile(file_stream)

    infos = pd.read_excel(excel, sheet_name="Infos Patient")
    histo = pd.read_excel(excel, sheet_name="Historique Séances")
    analyse = pd.read_excel(excel, sheet_name="Analyse Clinique")
    therapeu = pd.read_excel(excel, sheet_name="Thérapeutiques")
    litho = pd.read_excel(excel, sheet_name="Lithothérapie")
    suivi = pd.read_excel(excel, sheet_name="Suivi")

    if "Historique Séances" in sections:
        histo = append_rows(histo, sections["Historique Séances"])

    if "Analyse Clinique" in sections:
        analyse = append_rows(analyse, sections["Analyse Clinique"])

    if "Thérapeutiques" in sections:
        therapeu = append_rows(therapeu, sections["Thérapeutiques"])

    if "Lithothérapie" in sections:
        litho = append_rows(litho, sections["Lithothérapie"])

    if "Suivi" in sections:
        suivi = append_rows(suivi, sections["Suivi"])

    # Update Infos Patient B13 à B16
    update_map = {
        "Demande": 12,
        "Famille": 13,
        "Sante": 14,
        "Situation": 15
    }

    for field, row_idx in update_map.items():
        if field in sections and sections[field]:
            value = sections[field][0][-1]
            infos.iloc[row_idx, 1] = value

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        infos.to_excel(writer, sheet_name="Infos Patient", index=False)
        histo.to_excel(writer, sheet_name="Historique Séances", index=False)
        analyse.to_excel(writer, sheet_name="Analyse Clinique", index=False)
        therapeu.to_excel(writer, sheet_name="Thérapeutiques", index=False)
        litho.to_excel(writer, sheet_name="Lithothérapie", index=False)
        suivi.to_excel(writer, sheet_name="Suivi", index=False)

    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=updated.xlsx"}
    )
