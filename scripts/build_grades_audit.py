"""Rebuild Grades_Audit_Dashboard.pbix report layout to mirror the Excel Grade Audit report.

Usage: python build_grades_audit.py <input.pbix> <output.pbix>

Only Report/Layout is rewritten; the data model (DataModel) is copied untouched.
SecurityBindings is dropped (required when the layout changes; Desktop regenerates it).
"""
import json
import sys
import uuid
import zipfile

T = "VW_GRADES_AUDIT_REPORT"
SRC = "v"

# Row fields in the same order as the Excel columns A..Q (fields missing in the view are skipped)
ROW_FIELDS = [
    ("CAMPUS_NAME", "Campus Name"),
    ("HCT_ID", "HCT Id"),
    ("STUDENT_NAME", "Full Name"),
    ("CATALOG_TERM", "Catalog Term"),
    ("TOTAL_REGISTERED_CREDIT", "Total Registered Credits"),
    ("PROGRAM_CREDITS_COMPLETED", "Program Credits Completed"),
    ("PROGRAM_CREDITS_REQUIRED", "Program Credits Required"),
    ("PENDING_CREDITS", "Pending Credits"),
    ("ACADEMIC_STANDING", "Academic Standing"),
]
# Power BI freezes every matrix row header, so wide text fields (Advisor is always NULL in the view,
# Major/Program are slicers, Comments is long) go into the cell tooltip to leave room for course columns.
TOOLTIP_FIELDS = [
    ("COMMENTS", "Comments"),
    ("PROGRAM_DESCRIPTION", "Program Description"),
    ("ADVISOR", "Advisor"),
]
SLICERS = [
    ("CAMPUS_NAME", "Campus"),
    ("MAJOR_CODE", "Major Code"),
    ("PROGRAM", "Program"),
    ("CATALOG_TERM", "Catalog Term"),
]
PAGE_W, PAGE_H = 1600, 900


def lit(v):
    return {"expr": {"Literal": {"Value": v}}}


def s(text):
    return lit("'" + text.replace("'", "''") + "'")


def col(prop):
    return {"Column": {"Expression": {"SourceRef": {"Source": SRC}}, "Property": prop}}


def qref(prop):
    return f"{T}.{prop}"


def container(cfg, x, y, w, h, z):
    cfg["layouts"] = [{"id": 0, "position": {"x": x, "y": y, "z": z, "width": w, "height": h, "tabOrder": z}}]
    return {
        "x": x, "y": y, "z": z, "width": w, "height": h,
        "config": json.dumps(cfg), "filters": "[]",
    }


def title_box(x, y, w, h, z):
    cfg = {
        "name": uuid.uuid4().hex[:20],
        "singleVisual": {
            "visualType": "textbox",
            "drillFilterOtherVisuals": True,
            "objects": {"general": [{"properties": {"paragraphs": [{"textRuns": [{
                "value": "GRADE AUDIT REPORT",
                "textStyle": {"fontWeight": "bold", "fontSize": "18pt"},
            }]}]}}]},
        },
    }
    return container(cfg, x, y, w, h, z)


def slicer(prop, label, x, y, w, h, z):
    cfg = {
        "name": uuid.uuid4().hex[:20],
        "singleVisual": {
            "visualType": "slicer",
            "projections": {"Values": [{"queryRef": qref(prop), "active": True}]},
            "prototypeQuery": {
                "Version": 2,
                "From": [{"Name": SRC, "Entity": T, "Type": 0}],
                "Select": [{**col(prop), "Name": qref(prop), "NativeReferenceName": label}],
            },
            "columnProperties": {qref(prop): {"displayName": label}},
            "drillFilterOtherVisuals": True,
            "objects": {
                "data": [{"properties": {"mode": s("Dropdown")}}],
                "selection": [{"properties": {"selectAllCheckboxEnabled": lit("true")}}],
                "header": [{"properties": {"text": s(label)}}],
            },
        },
    }
    return container(cfg, x, y, w, h, z)


def grade_matrix(x, y, w, h, z):
    rows = [{"queryRef": qref(p), "active": True} for p, _ in ROW_FIELDS]
    grade_ref = f"Min({T}.GRADE)"
    select = [{**col(p), "Name": qref(p), "NativeReferenceName": lbl} for p, lbl in ROW_FIELDS]
    select.append({**col("COURSE_NAME"), "Name": qref("COURSE_NAME"), "NativeReferenceName": "Course"})
    # Function 3 = Min (shown as "First" for text) -> the grade for that student/course
    select.append({"Aggregation": {"Expression": col("GRADE"), "Function": 3},
                   "Name": grade_ref, "NativeReferenceName": "Grade"})
    col_props = {qref(p): {"displayName": lbl} for p, lbl in ROW_FIELDS}
    col_props[grade_ref] = {"displayName": "Grade"}
    tips = []
    for p, lbl in TOOLTIP_FIELDS:
        ref = f"Min({T}.{p})"
        select.append({"Aggregation": {"Expression": col(p), "Function": 3}, "Name": ref, "NativeReferenceName": lbl})
        col_props[ref] = {"displayName": lbl}
        tips.append({"queryRef": ref})
    header_bg = {"solid": {"color": lit("'#F3F3F3'")}}
    cfg = {
        "name": uuid.uuid4().hex[:20],
        "singleVisual": {
            "visualType": "pivotTable",
            "projections": {
                "Rows": rows,
                "Columns": [{"queryRef": qref("COURSE_NAME"), "active": True}],
                "Values": [{"queryRef": grade_ref}],
                "Tooltips": tips,
            },
            "prototypeQuery": {
                "Version": 2,
                "From": [{"Name": SRC, "Entity": T, "Type": 0}],
                "Select": select,
                "OrderBy": [{"Direction": 1, "Expression": col("STUDENT_NAME")}],
            },
            "columnProperties": col_props,
            "drillFilterOtherVisuals": True,
            "objects": {
                "rowHeaders": [{"properties": {
                    "stepped": lit("false"),
                    "showExpandCollapseButtons": lit("false"),
                    "backColor": header_bg,
                    "wordWrap": lit("false"),
                }}],
                "columnHeaders": [{"properties": {"backColor": header_bg, "wordWrap": lit("false")}}],
                "subTotals": [{"properties": {"rowSubtotals": lit("false"), "columnSubtotals": lit("false")}}],
                "values": [{"properties": {"fontSize": lit("9D")}}],
                "grid": [{"properties": {"gridVertical": lit("true"), "gridHorizontal": lit("true")}}],
            },
            "vcObjects": {"title": [{"properties": {"show": lit("false")}}]},
        },
    }
    return container(cfg, x, y, w, h, z)


def build_layout(layout):
    sec = layout["sections"][0]
    sec["displayName"] = "Grade Audit"
    sec["width"], sec["height"] = PAGE_W, PAGE_H
    sec["displayOption"] = 1
    vcs = [title_box(16, 8, 500, 48, 0)]
    sw, gap = 220, 12
    for i, (p, lbl) in enumerate(SLICERS):
        vcs.append(slicer(p, lbl, 16 + i * (sw + gap), 60, sw, 64, i + 1))
    vcs.append(grade_matrix(16, 136, PAGE_W - 32, PAGE_H - 152, len(SLICERS) + 1))
    sec["visualContainers"] = vcs
    return layout


def main(src, dst):
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w") as zout:
        for item in zin.infolist():
            if item.filename == "SecurityBindings":
                continue
            data = zin.read(item.filename)
            if item.filename == "Report/Layout":
                layout = build_layout(json.loads(data.decode("utf-16-le")))
                data = json.dumps(layout, separators=(",", ":")).encode("utf-16-le")
            elif item.filename == "[Content_Types].xml":
                data = data.replace(b'<Override PartName="/SecurityBindings" ContentType="" />', b"")
            zout.writestr(item, data, compress_type=item.compress_type)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
