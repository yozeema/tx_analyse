#!/usr/bin/env python3
"""
Convert the `data.data_string.data.series` payload inside a JSON file into
an XLSX workbook that contains both English field names and their Chinese
labels in the first two rows. The XLSX file is generated directly using the
OOXML specification so no third-party dependency is required.
"""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
import json
from pathlib import Path
from typing import Iterable, Sequence
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

FIELD_LABELS = {
    "timeMinute": "时刻",
    "commentCnt": "评论数",
    "commentCntRank": "互动高峰",
    "commentUcnt": "评论人数",
    "costDiamonds": "消耗钻石数",
    "consumeUcnt": "送礼人数",
    "earnScore": "音浪",
    "earnScoreRank": "送礼高峰",
    "expectMinute": "预计扶持时长",
    "expectWatchCnt": "预计进房数量",
    "followUcnt": "关注人数",
    "likeCnt": "点赞数",
    "likeUcnt": "点赞人数",
    "lottery": "抽奖",
    "luckymoneyCnt": "福袋钻石数",
    "operatorID": "操作员ID",
    "operatorName": "操作员名字",
    "pcuTotal": "在线人数",
    "realMinute": "实际扶持时长",
    "realWatchCnt": "实际进房数量",
    "watchUcnt": "进入直播间人数",
    "watchUcntRank": "在线观众高峰",
    "keyEvent": "关键事件",
}


def _load_data_string_series(json_path: Path) -> list[dict]:
    with json_path.open(encoding="utf-8") as fh:
        payload = json.load(fh)

    try:
        data_string = payload["data"]["data_string"]
    except KeyError as exc:  # pragma: no cover - defensive guard
        raise KeyError("Missing data.data_string in the provided JSON file") from exc

    try:
        data_payload = json.loads(data_string)
    except json.JSONDecodeError as exc:  # pragma: no cover - defensive guard
        raise ValueError("data_string is not valid JSON") from exc

    try:
        series = data_payload["data"]["series"]
    except KeyError as exc:  # pragma: no cover - defensive guard
        raise KeyError("data_string does not contain data.series") from exc

    if not isinstance(series, list):  # pragma: no cover - defensive guard
        raise TypeError("data.series must be a list")

    return series


def _ordered_fields(series: Iterable[dict]) -> list[str]:
    seen: set[str] = set()
    ordered: list[str] = []
    for entry in series:
        for key in entry:
            if key not in seen:
                ordered.append(key)
                seen.add(key)
    result: list[str] = []

    def _pop(field: str) -> str | None:
        if field in ordered:
            ordered.remove(field)
            return field
        return None

    time_minute = _pop("timeMinute")
    if time_minute:
        result.append(time_minute)

    _pop("groupPlay")  # drop unwanted column entirely
    watch_ucnt = _pop("watchUcnt")
    pcu_total = _pop("pcuTotal")

    result.extend(ordered)

    if watch_ucnt:
        result.append(watch_ucnt)
    if pcu_total:
        result.append(pcu_total)

    return result


def _column_name(index: int) -> str:
    """Convert a 1-based column index into Excel-style column letters."""
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _escape_cell_value(value: object) -> str:
    text = str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = escape(text)
    return text.replace("\n", "&#10;")


def _sheet_xml(rows: Sequence[Sequence[object]]) -> str:
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        "  <sheetData>",
    ]
    for row_index, row in enumerate(rows, start=1):
        lines.append(f'    <row r="{row_index}">')
        for column_index, value in enumerate(row, start=1):
            if value in ("", None):
                continue
            cell_ref = f"{_column_name(column_index)}{row_index}"
            text = _escape_cell_value(value)
            lines.append(
                f'      <c r="{cell_ref}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'
            )
        lines.append("    </row>")
    lines.append("  </sheetData>")
    lines.append("</worksheet>")
    return "\n".join(lines)


def _content_types_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


def _root_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""


def _workbook_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="data" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"""


def _workbook_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""


def _styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
  </fonts>
  <fills count="2">
    <fill>
      <patternFill patternType="none"/>
    </fill>
    <fill>
      <patternFill patternType="gray125"/>
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
"""


def _app_props_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>json_to_xlsx</Application>
</Properties>
"""


def _core_props_xml() -> str:
    timestamp = datetime.now(timezone.utc).replace(microsecond=False).isoformat()
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>json_to_xlsx</dc:creator>
  <cp:lastModifiedBy>json_to_xlsx</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:modified>
</cp:coreProperties>
"""


def _write_xlsx(rows: list[list[object]], xlsx_path: Path) -> None:
    sheet_xml = _sheet_xml(rows)
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)

    with ZipFile(xlsx_path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _content_types_xml())
        archive.writestr("_rels/.rels", _root_rels_xml())
        archive.writestr("docProps/core.xml", _core_props_xml())
        archive.writestr("docProps/app.xml", _app_props_xml())
        archive.writestr("xl/workbook.xml", _workbook_xml())
        archive.writestr("xl/_rels/workbook.xml.rels", _workbook_rels_xml())
        archive.writestr("xl/styles.xml", _styles_xml())
        archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert data.data_string JSON payload into an XLSX file."
    )
    parser.add_argument(
        "json_path",
        type=Path,
        help="Path to the source JSON file.",
    )
    parser.add_argument(
        "output_dir",
        type=Path,
        help="Directory where the XLSX file should be written.",
    )
    parser.add_argument(
        "--output-name",
        type=str,
        default=None,
        help="Optional XLSX file name. Defaults to <json stem>.xlsx",
    )
    return parser.parse_args()


def main() -> None:
    args = _parse_args()
    series = _load_data_string_series(args.json_path)

    if not series:
        raise SystemExit("data.series is empty, nothing to write.")

    fields = _ordered_fields(series)
    if "keyEvent" in fields:
        fields.remove("keyEvent")
    fields.append("keyEvent")  # always keep the key event column last
    chinese_labels = [FIELD_LABELS.get(field, field) for field in fields]
    data_rows = [[entry.get(field, "") for field in fields] for entry in series]

    output_name = args.output_name or (args.json_path.stem + ".xlsx")
    xlsx_path = args.output_dir / output_name

    _write_xlsx([fields, chinese_labels, *data_rows], xlsx_path)
    print(f"Wrote {xlsx_path}")


if __name__ == "__main__":
    main()
