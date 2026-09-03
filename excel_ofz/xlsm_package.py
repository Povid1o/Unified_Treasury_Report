"""Превращение .xlsx, собранного openpyxl, в .xlsm с макросами и кнопками.

openpyxl не умеет ни создавать проект VBA, ни рисовать фигуры с назначенным
макросом, поэтому пакет OOXML дособирается здесь вручную: добавляются части
`xl/vbaProject.bin` и `xl/drawings/drawing1.xml`, правятся типы содержимого и
связи.
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path
from typing import List, NamedTuple

EMU_PER_PX = 9525

NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

CT_WORKBOOK_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
CT_WORKBOOK_XLSM = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
CT_VBA = "application/vnd.ms-office.vbaProject"
CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
REL_VBA = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"


class Button(NamedTuple):
    text: str
    macro: str
    width_px: int
    color: str


def _escape(text: str) -> str:
    return (text.replace("&", "&amp;").replace("<", "&lt;")
                .replace(">", "&gt;").replace('"', "&quot;"))


def build_drawing_xml(buttons: List[Button], first_row: int, height_px: int = 30,
                      col_off_px: int = 4) -> str:
    """Фигуры-кнопки. Атрибут macro у xdr:sp — это и есть «Назначить макрос».

    По одной кнопке на строку: Excel обрезает colOff по фактической ширине
    столбца, поэтому раскладывать их в ряд одним смещением нельзя — кнопки
    схлопываются друг на друга.
    """
    parts = [
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<xdr:wsDr xmlns:xdr="{NS_XDR}" xmlns:a="{NS_A}">'
    ]
    for i, button in enumerate(buttons):
        parts.append(
            f'<xdr:oneCellAnchor>'
            f'<xdr:from><xdr:col>0</xdr:col><xdr:colOff>{col_off_px * EMU_PER_PX}</xdr:colOff>'
            f'<xdr:row>{first_row + i}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
            f'<xdr:ext cx="{button.width_px * EMU_PER_PX}" cy="{height_px * EMU_PER_PX}"/>'
            f'<xdr:sp macro="[0]!{_escape(button.macro)}" textlink="">'
            f'<xdr:nvSpPr>'
            f'<xdr:cNvPr id="{i + 2}" name="btn{_escape(button.macro)}"/>'
            f'<xdr:cNvSpPr/>'
            f'</xdr:nvSpPr>'
            f'<xdr:spPr>'
            f'<a:xfrm><a:off x="0" y="0"/>'
            f'<a:ext cx="{button.width_px * EMU_PER_PX}" cy="{height_px * EMU_PER_PX}"/></a:xfrm>'
            f'<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{button.color}"/></a:solidFill>'
            f'<a:ln><a:noFill/></a:ln>'
            f'</xdr:spPr>'
            f'<xdr:txBody>'
            f'<a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/>'
            f'<a:p><a:pPr algn="ctr"/>'
            f'<a:r><a:rPr lang="ru-RU" sz="1100" b="1">'
            f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr>'
            f'<a:t>{_escape(button.text)}</a:t></a:r></a:p>'
            f'</xdr:txBody>'
            f'</xdr:sp>'
            f'<xdr:clientData/>'
            f'</xdr:oneCellAnchor>'
        )
    parts.append("</xdr:wsDr>")
    return "".join(parts)


def _add_content_types(xml: str) -> str:
    xml = xml.replace(CT_WORKBOOK_XLSX, CT_WORKBOOK_XLSM)
    if 'Extension="bin"' not in xml:
        xml = re.sub(
            r"(<Types[^>]*>)",
            rf'\1<Default Extension="bin" ContentType="{CT_VBA}"/>',
            xml, count=1)
    if "/xl/drawings/drawing1.xml" not in xml:
        xml = xml.replace(
            "</Types>",
            f'<Override PartName="/xl/drawings/drawing1.xml" ContentType="{CT_DRAWING}"/></Types>')
    return xml


def _add_workbook_rel(xml: str) -> str:
    if "vbaProject.bin" in xml:
        return xml
    rel_id = _next_rel_id(xml)
    return xml.replace(
        "</Relationships>",
        f'<Relationship Id="{rel_id}" Type="{REL_VBA}" Target="vbaProject.bin"/></Relationships>')


def _next_rel_id(xml: str) -> str:
    used = {int(m) for m in re.findall(r'Id="rId(\d+)"', xml)}
    return f"rId{max(used, default=0) + 1}"


def _sheet_rels_with_drawing() -> str:
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{NS_PKG_REL}">'
            f'<Relationship Id="rIdDrawing1" Type="{REL_DRAWING}" Target="../drawings/drawing1.xml"/>'
            f'</Relationships>')


def _attach_drawing(sheet_xml: str) -> str:
    """Ссылка на рисунок обязана стоять в самом конце — таков порядок в схеме."""
    if "<drawing " in sheet_xml:
        return sheet_xml
    if f'xmlns:r="{NS_R}"' not in sheet_xml:
        sheet_xml = sheet_xml.replace("<worksheet ", f'<worksheet xmlns:r="{NS_R}" ', 1)
    return sheet_xml.replace("</worksheet>", '<drawing r:id="rIdDrawing1"/></worksheet>')


def to_xlsm(xlsx_path: Path, xlsm_path: Path, vba_project: bytes,
            buttons: List[Button], button_sheet: str = "xl/worksheets/sheet1.xml",
            first_button_row: int = 26) -> Path:
    """Пересобирает .xlsx в .xlsm: проект VBA + кнопки на указанном листе."""
    sheet_rels = button_sheet.replace("worksheets/", "worksheets/_rels/") + ".rels"

    with zipfile.ZipFile(xlsx_path) as src:
        names = src.namelist()
        if button_sheet not in names:
            raise ValueError(f"В книге нет части {button_sheet}: {names}")
        payload = {name: src.read(name) for name in names}

    payload["[Content_Types].xml"] = _add_content_types(
        payload["[Content_Types].xml"].decode("utf-8")).encode("utf-8")
    payload["xl/_rels/workbook.xml.rels"] = _add_workbook_rel(
        payload["xl/_rels/workbook.xml.rels"].decode("utf-8")).encode("utf-8")
    payload[button_sheet] = _attach_drawing(
        payload[button_sheet].decode("utf-8")).encode("utf-8")
    payload[sheet_rels] = _sheet_rels_with_drawing().encode("utf-8")
    payload["xl/drawings/drawing1.xml"] = build_drawing_xml(buttons, first_button_row).encode("utf-8")
    payload["xl/vbaProject.bin"] = vba_project

    xlsm_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = xlsm_path.with_suffix(".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        # [Content_Types].xml по спецификации OPC должен идти первой частью.
        dst.writestr("[Content_Types].xml", payload.pop("[Content_Types].xml"))
        for name, data in payload.items():
            dst.writestr(name, data)
    shutil.move(str(tmp), str(xlsm_path))
    return xlsm_path
