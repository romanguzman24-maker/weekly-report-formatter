#!/usr/bin/env python3
"""Weekly Report Formatter v9.15 — Village at First RR uses simple 13-col format (no set-aside/sub/tenant cols)"""
from flask import Flask, request, send_file, render_template_string, jsonify
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io, re, os, traceback
from datetime import datetime
from collections import OrderedDict


def read_expiring_rows(file_storage_or_path):
    """Read raw Yardi Expiring Leases export. Returns list of dict rows, sorted oldest -> newest."""
    from openpyxl import load_workbook
    wb = load_workbook(file_storage_or_path, data_only=True, read_only=True)
    ws = wb.active
    headers, rows = None, []
    for r in ws.iter_rows(values_only=True):
        if headers is None:
            if r and r[0] and str(r[0]).strip().lower().startswith("lease expires"):
                headers = [str(h).strip() if h is not None else "" for h in r]
            continue
        if not r or all(v is None or (isinstance(v, str) and not v.strip()) for v in r):
            continue
        first = r[0]
        if isinstance(first, str) and any(k in first.lower() for k in ("total","grand","summary","village at")):
            continue
        if not isinstance(first, datetime):
            continue
        rows.append({headers[i]: r[i] if i < len(r) else None for i in range(len(headers))})
    wb.close()
    rows.sort(key=lambda x: x.get("Lease Expires") or datetime.max)
    return rows


def build_monthly_counts(rows):
    """Returns list of (month_label, count) sorted oldest -> newest."""
    counts = OrderedDict()
    for r in rows:
        d = r.get("Lease Expires")
        if not isinstance(d, datetime):
            continue
        key = d.strftime("%b %Y")
        sk = (d.year, d.month)
        if key not in counts:
            counts[key] = {"sort": sk, "count": 0}
        counts[key]["count"] += 1
    return [(k, v["count"]) for k, v in sorted(counts.items(), key=lambda kv: kv[1]["sort"])]


def format_expiring_leases_tab(out_wb, source_file, sheet_name, property_name):
    """Build the formatted 'Expiring Leases (120 days)' tab. Sorted oldest to newest."""
    rows = read_expiring_rows(source_file)
    out_cols = ["Lease Expires","Unit","Resident","Market Rent","Current Rent",
                "Loss to Lease","Current Lease Term","Months At Property",
                "MTM?","Appr Status","Comments"]
    n_cols = len(out_cols)
    last_letter = get_column_letter(n_cols)
    if sheet_name in out_wb.sheetnames:
        del out_wb[sheet_name]
    ws = out_wb.create_sheet(sheet_name)
    thin_black = Side(style="thin", color="000000")
    thin_gray = Side(style="thin", color="CCCCCC")
    box = Border(top=thin_black, bottom=thin_black, left=thin_black, right=thin_black)
    box_gray = Border(top=thin_gray, bottom=thin_gray, left=thin_gray, right=thin_gray)
    ws.merge_cells(f"A1:{last_letter}1"); ws.merge_cells(f"A2:{last_letter}2"); ws.merge_cells(f"A3:{last_letter}3")
    ws["A1"] = property_name
    ws["A2"] = "Expiring Leases (120 days)"
    ws["A3"] = f"Report Date, {datetime.now().strftime('%m.%d.%y')}"
    title_fill = PatternFill("solid", fgColor=GREEN[2:])
    title_font_bold = Font(name="Calibri", size=9, bold=True, color="505050")
    title_font = Font(name="Calibri", size=9, color="505050")
    center = Alignment(horizontal="center", vertical="center")
    for ri, fnt in [(1, title_font_bold),(2, title_font_bold),(3, title_font)]:
        for col in range(1, n_cols+1):
            c = ws.cell(row=ri, column=col); c.fill=title_fill; c.font=fnt; c.alignment=center
    head_fill = PatternFill("solid", fgColor="D9D9D9")
    head_font = Font(name="Calibri", size=9, bold=True)
    for col in range(1, n_cols+1):
        c = ws.cell(row=4, column=col); c.fill=head_fill; c.font=head_font; c.alignment=center; c.border=box
    for i, label in enumerate(out_cols, start=1):
        c = ws.cell(row=5, column=i, value=label); c.fill=head_fill; c.font=head_font; c.alignment=center; c.border=box
    data_font = Font(name="Calibri", size=9)
    left_a = Alignment(horizontal="left", vertical="center")
    right_a = Alignment(horizontal="right", vertical="center")
    center_a = Alignment(horizontal="center", vertical="center")
    start_row = 6
    for i, row in enumerate(rows):
        r = start_row + i
        vals = [row.get("Lease Expires"), row.get("Unit"), row.get("Resident"),
                row.get("Market Rent"), row.get("Current Rent"), row.get("Loss to Lease"),
                row.get("Current Lease Term"), row.get("Months At Property"),
                row.get("MTM?"), row.get("Appr Status"),
                row.get("Comments ") or row.get("Comments") or None]
        for ci, v in enumerate(vals, start=1):
            c = ws.cell(row=r, column=ci, value=v); c.font=data_font; c.border=box_gray
            label = out_cols[ci-1]
            if label == "Lease Expires":
                c.number_format="mm/dd/yyyy"; c.alignment=center_a
            elif label in ("Market Rent","Current Rent","Loss to Lease"):
                c.number_format="#,##0"; c.alignment=right_a
            elif label in ("Unit","Current Lease Term","Months At Property","MTM?","Appr Status"):
                c.alignment=center_a
            else:
                c.alignment=left_a
    last_data_row = start_row + len(rows) - 1
    total_row = last_data_row + 1 if rows else start_row
    total_font = Font(name="Calibri", size=9, bold=True)
    ws.cell(row=total_row, column=1, value="Total")
    if rows:
        ws.cell(row=total_row, column=3, value=f"=COUNTA(C{start_row}:C{last_data_row})")
        for ci, label in enumerate(out_cols, start=1):
            if label in ("Market Rent","Current Rent","Loss to Lease"):
                cl = get_column_letter(ci)
                cell = ws.cell(row=total_row, column=ci, value=f"=SUM({cl}{start_row}:{cl}{last_data_row})")
                cell.number_format="#,##0"; cell.alignment=right_a
    for col in range(1, n_cols+1):
        c = ws.cell(row=total_row, column=col); c.font=total_font; c.fill=head_fill; c.border=box
        if c.alignment is None or c.alignment.horizontal is None:
            c.alignment = right_a
    widths = {"Lease Expires":13,"Unit":9,"Resident":28,"Market Rent":13,"Current Rent":13,
              "Loss to Lease":13,"Current Lease Term":13,"Months At Property":13,
              "MTM?":8,"Appr Status":13,"Comments":40}
    for i, label in enumerate(out_cols, start=1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(label, 12)
    ws.freeze_panes = "A6"
    return ws


def write_expiring_summary_block(ws, top_left_row, top_left_col, monthly_counts):
    """Writes a 2-col Month/Count block with Total row. Returns last row used."""
    BLUE = "BDD7EE"
    thin_black = Side(style="thin", color="000000")
    thin_gray_b = Side(style="thin", color="BFBFBF")
    box = Border(top=thin_black, bottom=thin_black, left=thin_black, right=thin_black)
    box_gray = Border(top=thin_gray_b, bottom=thin_gray_b, left=thin_gray_b, right=thin_gray_b)
    r0, c0, c1 = top_left_row, top_left_col, top_left_col+1
    last_letter = get_column_letter(c1)
    ws.merge_cells(start_row=r0, start_column=c0, end_row=r0, end_column=c1)
    h = ws.cell(row=r0, column=c0, value="Expiring Leases (120 days)")
    h.fill = PatternFill("solid", fgColor=BLUE)
    h.font = Font(name="Calibri", size=9, bold=True)
    h.alignment = Alignment(horizontal="center", vertical="center")
    h.border = box
    ws.cell(row=r0, column=c1).border = box
    for col, lbl in [(c0,"Month"),(c1,"Count")]:
        c = ws.cell(row=r0+1, column=col, value=lbl)
        c.fill = PatternFill("solid", fgColor=BLUE)
        c.font = Font(name="Calibri", size=9, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = box
    data_start = r0 + 2
    for i, (m, cnt) in enumerate(monthly_counts):
        r = data_start + i
        cm = ws.cell(row=r, column=c0, value=m)
        cm.font = Font(name="Calibri", size=9)
        cm.alignment = Alignment(horizontal="left", vertical="center")
        cm.border = box_gray
        cc = ws.cell(row=r, column=c1, value=cnt)
        cc.font = Font(name="Calibri", size=9)
        cc.alignment = Alignment(horizontal="center", vertical="center")
        cc.border = box_gray
        cc.number_format = "0"
    total_r = data_start + len(monthly_counts)
    last_data_r = total_r - 1 if monthly_counts else data_start
    tl = ws.cell(row=total_r, column=c0, value="Total")
    tl.fill = PatternFill("solid", fgColor=BLUE)
    tl.font = Font(name="Calibri", size=9, bold=True)
    tl.alignment = Alignment(horizontal="center", vertical="center")
    tl.border = box
    if monthly_counts:
        tc = ws.cell(row=total_r, column=c1, value=f"=SUM({last_letter}{data_start}:{last_letter}{last_data_r})")
    else:
        tc = ws.cell(row=total_r, column=c1, value=0)
    tc.fill = PatternFill("solid", fgColor=BLUE)
    tc.font = Font(name="Calibri", size=9, bold=True)
    tc.alignment = Alignment(horizontal="center", vertical="center")
    tc.border = box
    tc.number_format = "0"
    return total_r
