#!/usr/bin/env python3
"""Weekly Report Formatter v9.8 — Rent Roll ST fix + AR credits/prepays grouping"""
from flask import Flask, request, send_file, render_template_string, jsonify
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io, re, os, traceback

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

GREEN='FF7AD694'; GRAY_HDR='FFBFBFBF'; GRAY_AR='FFD9D9D9'
KG_RED='FFF28E86'; KG_YEL='FFFDD868'; KG_BLUE='FF8CB5F9'
WHITE='FFFFFFFF'; BLACK='FF000000'; DARKGRAY='FF505050'; RED_FONT='FFFF0000'
BLUE_IN='FF8CB5F9'

PROPERTY_UNITS = {
    'Village at Madrone (fka Village at Morgan Hill) (x93)': 249,
    'Village at First': 120,
    'Village at Santa Teresa': 100,
}

def get_total_units(prop):
    for key, val in PROPERTY_UNITS.items():
        if key.lower() in prop.lower() or prop.lower() in key.lower():
            return val
    return 249

def gfill(h): return PatternFill(fill_type='solid', fgColor=h)
def gfont(bold=False,sz=9,color='FF000000'): return Font(name='Calibri',size=sz,bold=bold,color=color)
def galign(h='left',v='center',wrap=False): return Alignment(horizontal=h,vertical=v,wrap_text=wrap)
T=Side(style='thin',color='FF000000'); TG=Side(style='thin',color='FFCCCCCC')
def bblack(): return Border(top=T,bottom=T,left=T,right=T)
def bgray():  return Border(top=TG,bottom=TG,left=TG,right=TG)

def parse_ua(ws):
    out=[]; status='Occupied'
    rows=list(ws.iter_rows(values_only=True))
    has_kg=False; unit_col=0; unit_pat=r'^\d{3,5}$'
    for row in rows:
        if not row: continue
        v0=str(row[0] or '').strip(); v1=str(row[1] or '').strip() if len(row)>1 else ''
        if re.match(r'^\d{2}-\d{3}',v0): unit_col=0; unit_pat=r'^\d{2}-\d{3}'; break
        if re.match(r'^\d{2}-\d{3}',v1): unit_col=1; unit_pat=r'^\d{2}-\d{3}'; break
        if re.match(r'^\d{3,5}$',v0): unit_col=0; unit_pat=r'^\d{3,5}$'; break
    for row in rows[:10]:
        if any('kg' in str(v or '').lower() for v in row):
            has_kg=True; break
    for row in rows:
        if not row or all(v is None or str(v).strip()=='' for v in row): continue
        c0=str(row[0] or '').strip()
        c1=str(row[1] or '').strip() if len(row)>1 else ''
        hdr=c0 if c0 else c1
        lo=hdr.lower()
        if any(x in lo for x in ['- vacant','- notice','- occupied','- past','- current']):
            if   'occupied' in lo or ('current' in lo and 'notice' not in lo): status='Occupied'
            elif 'vacant' in lo: status='Vacant'
            elif 'notice' in lo: status='Notice'
            continue
        if re.match(r'^(Unit Availability|Showing|Group|As Of|Total|Property)',hdr,re.I): continue
        if hdr in ('Unit',): continue
        if unit_col==1 and str(row[0] or '').strip() in ('Vacant','Notice','Occupied'):
            status=str(row[0]).strip()
        unit=str(row[unit_col] or '').strip() if len(row)>unit_col else ''
        if not re.match(unit_pat, unit): continue
        o=unit_col
        if has_kg:
            out.append({'status':status,'unit':unit,
                'res_id':str(row[o+1] or '').strip() if len(row)>o+1 else '',
                'name':str(row[o+2] or '').strip() if len(row)>o+2 else '',
                'kg_app':row[o+3] if len(row)>o+3 else None,
                'kg_pend':row[o+4] if len(row)>o+4 else None,
                'site_pend':row[o+5] if len(row)>o+5 else None,
                'res_rent':row[o+6] if len(row)>o+6 else None,
                'unit_rent':row[o+7] if len(row)>o+7 else None,
                'res_dep':row[o+8] if len(row)>o+8 else None,
                'unit_dep':row[o+9] if len(row)>o+9 else None,
                'yardi_st':str(row[o+10] or '').strip() if len(row)>o+10 else '',
                'days':row[o+11] if len(row)>o+11 else None,
                'make_rdy':row[o+12] if len(row)>o+12 else None,
                'move_in':row[o+13] if len(row)>o+13 else None,
                'hold':str(row[o+14] or '').strip() if len(row)>o+14 else '',
                'hold_until':row[o+15] if len(row)>o+15 else None,
                'notice':row[o+16] if len(row)>o+16 else None,
                'move_out':row[o+17] if len(row)>o+17 else None,
                'lease_sgn':row[o+18] if len(row)>o+18 else None,
                'lease_from':row[o+19] if len(row)>o+19 else None,
                'lease_to':row[o+20] if len(row)>o+20 else None,
                'has_kg':True})
        else:
            out.append({'status':status,'unit':unit,
                'res_id':str(row[o+1] or '').strip() if len(row)>o+1 else '',
                'name':str(row[o+2] or '').strip() if len(row)>o+2 else '',
                'kg_app':None,'kg_pend':None,'site_pend':None,
                'res_rent':row[o+3] if len(row)>o+3 else None,
                'unit_rent':row[o+4] if len(row)>o+4 else None,
                'res_dep':row[o+5] if len(row)>o+5 else None,
                'unit_dep':row[o+6] if len(row)>o+6 else None,
                'yardi_st':str(row[o+7] or '').strip() if len(row)>o+7 else '',
                'days':row[o+8] if len(row)>o+8 else None,
                'make_rdy':row[o+9] if len(row)>o+9 else None,
                'move_in':row[o+10] if len(row)>o+10 else None,
                'hold':str(row[o+11] or '').strip() if len(row)>o+11 else '',
                'hold_until':row[o+12] if len(row)>o+12 else None,
                'notice':row[o+13] if len(row)>o+13 else None,
                'move_out':row[o+14] if len(row)>o+14 else None,
                'lease_sgn':row[o+15] if len(row)>o+15 else None,
                'lease_from':row[o+16] if len(row)>o+16 else None,
                'lease_to':row[o+17] if len(row)>o+17 else None,
                'has_kg':False})
    return out

def fmt_ua(wb_out, raw_bytes, date, prop):
    wb_r=openpyxl.load_workbook(io.BytesIO(raw_bytes),data_only=True,keep_vba=False,read_only=True)
    data=parse_ua(wb_r.active)
    wb_r.close()
    V=[r for r in data if r['status']=='Vacant']
    N=[r for r in data if r['status']=='Notice']
    O=[r for r in data if r['status']=='Occupied']
    tab=f'Unit Availability {date}'
    if tab in wb_out.sheetnames: del wb_out[tab]
    ws=wb_out.create_sheet(tab); MC=21
    for ti,(text,bold) in enumerate(zip(
        ['Unit Availability Details ',prop,f'As Of: {date}','Showing Pre-Leased: Yes','Showing Occupied: Yes'],
        [True,False,True,False,False])):
        r=ti+1
        for c in range(1,MC+1): ws.cell(r,c).fill=gfill(GREEN); ws.cell(r,c).font=gfont(bold=bold); ws.cell(r,c).alignment=Alignment(horizontal='left',vertical='center',wrap_text=False)
        ws.cell(r,1).value=text
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=MC)
    h6=[None,'Unit','Resident','Name','KG Approved','KG Pend','Site Pending','Resident','Unit','Resident','Unit','Status','Days','Make','Move','Hold','Notice','Move','Lease','Lease','Lease']
    h7=[' ',' ',' ',' ',' ',' ',' ','Rent','Rent','Deposit','Deposit',' ','Vacant','Ready','In',' ',' ','Out','Sign','From','To']
    for c in range(1,MC+1):
        for r,hdr in [(6,h6),(7,h7)]:
            cell=ws.cell(r,c); cell.value=hdr[c-1] if hdr[c-1] is not None else ''
            cell.font=gfont(bold=True); cell.fill=gfill(GRAY_HDR)
            cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=False)
    DATE_COLS={14,15,17,18,19,20,21}
    rn=8; blank=False
    for p in V+N+O:
        isVN=p['status'] in ('Vacant','Notice')
        if p['status']=='Occupied' and not blank:
            for c in range(1,MC+1): ws.cell(rn,c).fill=gfill(WHITE)
            ws.row_dimensions[rn].height=15.0; rn+=1; blank=True
        def sc(col,val,bg=WHITE,h='left',_rn=rn):
            cell=ws.cell(_rn,col); cell.value=val; cell.font=gfont(); cell.fill=gfill(bg)
            cell.alignment=Alignment(horizontal=h,vertical='center',wrap_text=False)
            if col in DATE_COLS and val is not None:
                cell.number_format='MM/DD/YY'
        sc(1,p['status']); sc(2,p['unit']); sc(3,p['res_id'] or None)
        sc(4,None if p['status']=='Vacant' else (p['name'] or None))
        sc(5,None,KG_RED if isVN else WHITE); sc(6,None,KG_YEL if isVN else WHITE); sc(7,None,KG_BLUE if isVN else WHITE)
        sc(8,p['res_rent'] if p['res_rent'] is not None else 0,h='right')
        sc(9,p['unit_rent'] if p['unit_rent'] is not None else 0,h='right')
        sc(10,p['res_dep'] if p['res_dep'] is not None else 0,h='right')
        sc(11,p['unit_dep'] if p['unit_dep'] is not None else 0,h='right')
        sc(12,p['yardi_st'] or None); sc(13,p['days'] if p['days'] is not None else None,h='right')
        sc(14,p['make_rdy'] if p['make_rdy'] is not None else None,h='center')
        sc(15,p['move_in'] if p['move_in'] is not None else None,h='center')
        sc(16,p['hold'] or None)
        sc(17,p['hold_until'] if p['hold_until'] is not None else None,h='center')
        sc(18,p['move_out'] if p['move_out'] is not None else None,h='center')
        sc(19,p['lease_sgn'] if p['lease_sgn'] is not None else None,h='center')
        sc(20,p['lease_from'] if p['lease_from'] is not None else None,h='center')
        sc(21,p['lease_to'] if p['lease_to'] is not None else None,h='center')
        ws.row_dimensions[rn].height=15.0; rn+=1
    for col,w in {'A':14.28,'B':11.42,'C':21.42,'D':15,'E':12,'F':10,'G':12,'H':10,'I':11.42,'J':10,'K':10,'L':7.14,'M':11.42,'N':10,'O':10,'P':8,'Q':10,'R':10,'S':10,'T':10,'U':10}.items():
        ws.column_dimensions[col].width=w
    ws.row_dimensions[1].height=15.75
    for r in range(2,8): ws.row_dimensions[r].height=15.0
    ws.freeze_panes='A10'
    return ws, len(V), len(N), len(O)

def get_notes(wb, prefix):
    notes={}
    tabs=[n for n in wb.sheetnames if prefix.lower() in n.lower()]
    if not tabs: return notes
    rows=list(wb[tabs[-1]].iter_rows(values_only=True))
    hi=next((i for i,r in enumerate(rows[:10]) if any(str(c or '').lower()=='resident' for c in r)),-1)
    if hi==-1: return notes
    hdr=rows[hi]; ri=next((i for i,c in enumerate(hdr) if str(c or '').lower()=='resident'),-1); ni=len(hdr)-1
    for row in rows[hi+1:]:
        id_=str(row[ri] or '').strip() if ri>=0 else ''; note=str(row[ni] or '').strip() if len(row)>ni else ''
        if id_ and note: notes[id_]=note
    return notes

def fmt_ar(wb_out, raw_bytes, date, prev_notes, is_sub):
    wb_r=openpyxl.load_workbook(io.BytesIO(raw_bytes),data_only=True,keep_vba=False,read_only=True)
    rr=list(wb_r.active.iter_rows(values_only=True)); wb_r.close()
    NL='Comments'; tc=BLACK if is_sub else DARKGRAY
    tab=f'{"SUB AR" if is_sub else "Tenant AR"} {date}'
    if tab in wb_out.sheetnames: del wb_out[tab]
    ws=wb_out.create_sheet(tab); MC=13
    for ti in range(3):
        for c in range(1,MC+1):
            ws.cell(ti+1,c).fill=gfill(GREEN); ws.cell(ti+1,c).font=gfont(color=tc)
            ws.cell(ti+1,c).alignment=Alignment(horizontal='center',vertical='center',wrap_text=False)
        ws.cell(ti+1,1).value=str(rr[ti][0] if ti<len(rr) and rr[ti] else '')
    h4=['','','','','Total','','','','','','','','']
    h5=['','','','','Unpaid','0-30','31-60','61-90','Over 90','','','','']
    h6=['Unit','Resident','Status','Name','Charges','days','days','days','days','Prepays','Suspense','Balance','Comments']
    for c in range(1,MC+1):
        ra=5<=c<=12
        for ri,hdr in [(4,h4),(5,h5),(6,h6)]:
            cell=ws.cell(ri,c); cell.value=hdr[c-1]; cell.font=gfont(bold=True)
            cell.fill=gfill(GRAY_AR)
            cell.alignment=Alignment(horizontal='right' if ra else ('center' if c==13 else 'left'),vertical='center',wrap_text=False)
    for r in range(1,4): ws.cell(r,13).fill=gfill(GREEN); ws.cell(r,13).font=gfont(color=tc)
    hi=next((i for i,r in enumerate(rr[:10]) if r and any(str(c or '').lower()=='unit' for c in r) and any(str(c or '').lower()=='resident' for c in r)),5)
    ev,cu,no,cr=[],[],[],[]
    for row in rr[hi+1:]:
        if not row or all(c is None or c=='' for c in row): continue
        st=str(row[2] or '').strip().lower(); u=str(row[0] or '').strip()
        if re.search(r'subtotal|village at|^total$',u,re.I): continue
        if not u or not re.match(r'^\d{2}',u): continue
        # A row is a credit if Total Unpaid Charges < 0 OR Prepays < 0
        charges=row[4]; prepays=row[9] if len(row)>9 else None
        try: charge_val=float(str(charges or 0).replace(',',''))
        except: charge_val=0
        try: prepay_val=float(str(prepays or 0).replace(',',''))
        except: prepay_val=0
        if charge_val < 0 or prepay_val < 0:
            cr.append(row)
        elif st in ('eviction','past'): ev.append(row)
        elif st=='notice': no.append(row)
        else: cu.append(row)
    def sk(r):
        try: return -(float(str(r[4] or 0).replace(',','')))
        except: return 0
    ev.sort(key=sk); cu.sort(key=sk); no.sort(key=sk)
    def write_row(rn, row, rc):
        rid=str(row[1] or '').strip(); note=prev_notes.get(rid,'')
        for c in range(1,13):
            v=row[c-1]; sv=str(v if v is not None else '').strip()
            try: num=float(sv.replace(',',''))
            except: num=None
            isn=5<=c<=12 and num is not None and sv!=''
            cell=ws.cell(rn,c); cell.value=num if isn else (sv or None)
            cell.font=gfont(color=rc); cell.fill=gfill(WHITE)
            cell.alignment=Alignment(horizontal='right' if c>=5 else 'left',vertical='center',wrap_text=False)
            if isn: cell.number_format='#,##0.00'
        nc=ws.cell(rn,13); nc.value=None; nc.font=gfont(color=BLACK)
        nc.fill=gfill(WHITE); nc.alignment=Alignment(horizontal='center',vertical='center',wrap_text=False)
        nc.number_format='@'
    data_start=7; rn=data_start
    for row in ev+cu+no:
        st=str(row[2] or '').strip().lower()
        rc=RED_FONT if (not is_sub and st in ('notice','eviction','past')) else BLACK
        write_row(rn,row,rc); rn+=1
    tb=bblack()
    pos_end=rn-1
    def write_total(rn, start, end, label='Total'):
        ws.cell(rn,1).value=label; ws.cell(rn,1).font=gfont(bold=True); ws.cell(rn,1).fill=gfill(WHITE); ws.cell(rn,1).border=tb
        for c in range(5,13):
            cell=ws.cell(rn,c)
            cell.value=f'=SUM({get_column_letter(c)}{start}:{get_column_letter(c)}{end})'
            cell.font=gfont(bold=True); cell.fill=gfill(WHITE); cell.border=tb; cell.number_format='#,##0.00'
        for c in [2,3,4,13]: ws.cell(rn,c).fill=gfill(WHITE); ws.cell(rn,c).font=gfont(bold=True); ws.cell(rn,c).border=tb
    write_total(rn,data_start,pos_end); rn+=1
    if cr:
        rn+=1
        for c in range(1,MC+1):
            cell=ws.cell(rn,c); cell.fill=gfill(GRAY_AR); cell.font=gfont(bold=True)
            cell.alignment=Alignment(horizontal='left',vertical='center',wrap_text=False)
        ws.cell(rn,1).value='Credits / Prepays'; ws.cell(rn,4).value='Name'
        ws.cell(rn,5).value='Charges'; ws.cell(rn,10).value='Prepays'; ws.cell(rn,11).value='Suspense'; ws.cell(rn,12).value='Balance'
        for c in [5,10,11,12]: ws.cell(rn,c).alignment=Alignment(horizontal='right',vertical='center',wrap_text=False)
        rn+=1
        cr_start=rn
        for row in cr:
            write_row(rn,row,BLACK); rn+=1
        cr_end=rn-1
        write_total(rn,cr_start,cr_end,'Credits Total'); rn+=1
    for i,w in enumerate([9,13,10,24,12,10,10,10,10,10,10,12,38],1): ws.column_dimensions[get_column_letter(i)].width=w
    ws.freeze_panes='A7'
    return ws, len(ev), len(cu), len(no), pos_end, data_start

def fmt_rr(wb_out, raw_bytes, date, prop):
    wb_r=openpyxl.load_workbook(io.BytesIO(raw_bytes),data_only=True,keep_vba=False,read_only=True)
    rr=list(wb_r.active.iter_rows(values_only=True)); wb_r.close()
    tab=f'Rent Roll {date}'
    if tab in wb_out.sheetnames: del wb_out[tab]
    ws=wb_out.create_sheet(tab)

    import datetime as _dt

    MC=15
    title1=str(rr[0][0] or 'FPI Rent Roll') if rr else 'FPI Rent Roll'
    title2=str(rr[1][0] or prop) if len(rr)>1 else prop
    title3=str(rr[2][0] or f'As of Date: {date}') if len(rr)>2 else f'As of Date: {date}'
    title4=str(rr[3][0] or '') if len(rr)>3 else ''
    for r,txt,bold in [(1,title1,True),(2,title2,False),(3,title3,True),(4,title4,False)]:
        for c in range(1,MC+1):
            ws.cell(r,c).fill=gfill(GREEN)
            ws.cell(r,c).font=gfont(bold=bold)
            ws.cell(r,c).alignment=Alignment(horizontal='center',vertical='center',wrap_text=False)
        ws.cell(r,1).value=txt
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=MC)

    h5=['Unit','Unit Type','Unit Set Aside','Resident Name','Sq Ft',
        'Market Rent','Loss/Gain\nto Lease','Sub Rent','Tenant Rent',
        'Lease Rent','Vacancy','Deposit','Move In','Lease From','Lease To']
    for c,h in enumerate(h5,1):
        cell=ws.cell(5,c); cell.value=h; cell.font=gfont(bold=True)
        cell.fill=gfill(GRAY_HDR)
        cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.row_dimensions[5].height=28

    hi=next((i for i,r in enumerate(rr[:10]) if r and any(str(c or '').lower()=='unit' for c in r) and any('type' in str(c or '').lower() for c in r)),4)

    # Auto-detect unit pattern and column offset
    rr_offset=0; unit_pat=r'^\d{2}-\d{3}'
    for row in rr[hi+1:]:
        if not row: continue
        v0=str(row[0] or '').strip()
        v1=str(row[1] or '').strip() if len(row)>1 else ''
        if re.match(r'^\d{2}-\d{3}',v0): rr_offset=0; unit_pat=r'^\d{2}-\d{3}'; break
        if re.match(r'^\d{2}-\d{3}',v1): rr_offset=1; unit_pat=r'^\d{2}-\d{3}'; break
        if re.match(r'^\d{3,5}$',v0): rr_offset=0; unit_pat=r'^\d{3,5}'; break
        if re.match(r'^\d{3,5}$',v1): rr_offset=1; unit_pat=r'^\d{3,5}'; break

    # Detect if this is a Santa Teresa / 4-digit property with extra header cols
    # Santa Teresa raw export has: Unit, Unit Type, Unit Set Aside (%), flag col, Resident Name, ...
    # vs Madrone: Unit, Unit Type, Resident Name, ...
    # We detect by checking if col 2 (index) in first data row looks like a percentage or set-aside value
    is_st_format = False
    for row in rr[hi+1:]:
        if not row: continue
        unit_v = str(row[rr_offset] or '').strip()
        if not re.match(unit_pat, unit_v): continue
        # Check col rr_offset+2 — if it looks like a % (30%/50%/60%) or numeric set-aside, it's ST format
        col2_v = str(row[rr_offset+2] or '').strip()
        if re.match(r'^\d{1,3}%?$', col2_v) or col2_v in ('30%','50%','60%','30','50','60'):
            is_st_format = True
        break

    V,O=[],[]
    for row in rr[hi+1:]:
        if not row or all(c is None or c=='' for c in row): continue
        if len(row) <= rr_offset: continue
        unit=str(row[rr_offset] or '').strip()
        if not re.match(unit_pat,unit): continue
        o=rr_offset
        # For ST format: name is at o+4 (after unit, unit_type, set_aside%, flag_col)
        # For Madrone format: name is at o+2
        name_idx = o+4 if is_st_format else o+2
        rname=str(row[name_idx] or '').strip() if len(row)>name_idx else ''
        if rname.strip().upper() in ('VACANT',' VACANT') or not rname.strip(): V.append((row,o))
        else: O.append((row,o))
    V.sort(key=lambda x:str(x[0][x[1]] or '').strip())
    O.sort(key=lambda x:str(x[0][x[1]] or '').strip())

    def set_aside_from_col(val):
        """Parse set-aside directly from a column value (ST format)."""
        s = str(val or '').strip().replace('%','')
        if s == '30': return '30%'
        if s == '50': return '50%'
        if s == '60': return '60%'
        return str(val or '').strip() if val else ''

    def set_aside(unit_type):
        """Decode set-aside from unit type code (Madrone format)."""
        code=str(unit_type or '').strip()
        if not code: return ''
        m=re.search(r'(1|2)(50|60)',code)
        if m:
            pct=m.group(2)
            if pct=='50': return '50%'
            if pct=='60': return '60%'
        last=code[-1].upper()
        if last=='3': return '30%'
        if last=='5': return '50%'
        if last=='6': return '60%'
        if last=='M': return 'Exempt Unit'
        return ''

    def write_rr_row(rn, row, o, fc):
        unit=str(row[o] or '').strip()
        ut=str(row[o+1] or '').strip() if len(row)>o+1 else ''

        if is_st_format:
            # ST columns: unit@o, unit_type@o+1, set_aside@o+2, flag@o+3, name@o+4,
            #             sq_ft@o+5, market_rent@o+6, loss_gain@o+7, sub_rent@o+8,
            #             tenant_rent@o+9, lease_rent@o+10, vacancy@o+11, deposit@o+12,
            #             move_in@o+13, lease_from@o+14, lease_to@o+15
            sa_raw = row[o+2] if len(row)>o+2 else None
            sa_display = set_aside_from_col(sa_raw)
            rname = str(row[o+4] or '').strip() if len(row)>o+4 else ''
            sq    = row[o+5]  if len(row)>o+5  else None
            mr    = row[o+6]  if len(row)>o+6  else None
            lg    = row[o+7]  if len(row)>o+7  else None
            sr    = row[o+8]  if len(row)>o+8  else None
            tr    = row[o+9]  if len(row)>o+9  else None
            lr    = row[o+10] if len(row)>o+10 else None
            vac_r = row[o+11] if len(row)>o+11 else None
            dep   = row[o+12] if len(row)>o+12 else None
            mi    = row[o+13] if len(row)>o+13 else None
            lf    = row[o+14] if len(row)>o+14 else None
            lt    = row[o+15] if len(row)>o+15 else None
        else:
            # Madrone columns: unit@o, unit_type@o+1, name@o+2, sq@o+3, mr@o+4,
            #                  lg@o+5, sr@o+6, tr@o+7, lr@o+8, vac@o+9, dep@o+10,
            #                  mi@o+11, lf@o+12, lt@o+13
            sa_display = set_aside(ut)
            rname = str(row[o+2] or '').strip() if len(row)>o+2 else ''
            sq    = row[o+3]  if len(row)>o+3  else None
            mr    = row[o+4]  if len(row)>o+4  else None
            lg    = row[o+5]  if len(row)>o+5  else None
            sr    = row[o+6]  if len(row)>o+6  else None
            tr    = row[o+7]  if len(row)>o+7  else None
            lr    = row[o+8]  if len(row)>o+8  else None
            vac_r = row[o+9]  if len(row)>o+9  else None
            dep   = row[o+10] if len(row)>o+10 else None
            mi    = row[o+11] if len(row)>o+11 else None
            lf    = row[o+12] if len(row)>o+12 else None
            lt    = row[o+13] if len(row)>o+13 else None

        isvac = rname.strip().upper() in ('VACANT',' VACANT') or not rname.strip()

        def sc(col,val,h='left',fmt=None):
            cell=ws.cell(rn,col); cell.value=val; cell.font=gfont(color=fc); cell.fill=gfill(WHITE)
            cell.alignment=Alignment(horizontal=h,vertical='center',wrap_text=False)
            if fmt: cell.number_format=fmt

        sc(1,unit)
        sc(2,ut)
        sc(3,sa_display,'center')
        sc(4,' VACANT' if isvac else rname)
        sc(5,sq or 0,'center','#,##0')
        sc(6,mr or 0,'right','#,##0.00')
        try: lg_val=0 if isvac else float(str(lg or 0).replace(',',''))
        except: lg_val=0
        sc(7,lg_val,'right','#,##0.00')
        try: sr_val=float(str(sr or 0).replace(',',''))
        except: sr_val=0
        sc(8,sr_val,'right','#,##0.00')
        try: tr_val=float(str(tr or 0).replace(',',''))
        except: tr_val=0
        sc(9,tr_val,'right','#,##0.00')
        try: lr_val=float(str(lr or 0).replace(',',''))
        except: lr_val=0
        sc(10,lr_val,'right','#,##0.00')
        try: vac_val=-(float(str(mr or 0).replace(',',''))) if isvac else 0
        except: vac_val=0
        sc(11,vac_val,'right','#,##0.00')
        try: dep_val=float(str(dep or 0).replace(',',''))
        except: dep_val=0
        sc(12,dep_val,'right','#,##0.00')
        for col,dv in [(13,mi),(14,lf),(15,lt)]:
            cell=ws.cell(rn,col); cell.value=dv; cell.font=gfont(color=fc); cell.fill=gfill(WHITE)
            cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=False)
            if dv is not None and (isinstance(dv,_dt.datetime) or (isinstance(dv,(int,float)) and dv>40000)):
                cell.number_format='MM/DD/YY'
        ws.row_dimensions[rn].height=15.0

    rn=6
    for row,o in V: write_rr_row(rn,row,o,RED_FONT); rn+=1
    for row,o in O: write_rr_row(rn,row,o,BLACK); rn+=1
    data_end=rn-1

    for c in range(1,MC+1):
        ws.cell(rn,c).fill=gfill(GRAY_HDR); ws.cell(rn,c).font=gfont(bold=True)
    for c in range(5,13):
        ws.cell(rn,c).value=f'=SUM({chr(64+c)}6:{chr(64+c)}{data_end})'
        ws.cell(rn,c).number_format='#,##0.00'
        ws.cell(rn,c).alignment=Alignment(horizontal='right',vertical='center')
    rn+=1

    rn+=1
    for c in range(1,MC+1): ws.cell(rn,c).fill=gfill(GREEN)
    ws.cell(rn,1).value='Non-Revenue Units'; ws.cell(rn,1).font=gfont(bold=True,color='FF000000')
    ws.merge_cells(start_row=rn,start_column=1,end_row=rn,end_column=MC)
    rn+=1
    ws.cell(rn,4).value='No Data Available'; ws.cell(rn,4).font=gfont()
    rn+=2

    occ_sq=occ_mr=occ_sr=occ_tr=occ_dep=0; occ_cnt=0
    vac_sq=vac_mr=0; vac_cnt=len(V)
    for row,o in O:
        sq_idx  = o+5 if is_st_format else o+3
        mr_idx  = o+6 if is_st_format else o+4
        sr_idx  = o+8 if is_st_format else o+6
        tr_idx  = o+9 if is_st_format else o+7
        dep_idx = o+12 if is_st_format else o+10
        try: occ_sq+=float(row[sq_idx] or 0)
        except: pass
        try: occ_mr+=float(str(row[mr_idx] or 0).replace(',',''))
        except: pass
        try: occ_sr+=float(str(row[sr_idx] or 0).replace(',',''))
        except: pass
        try: occ_tr+=float(str(row[tr_idx] or 0).replace(',',''))
        except: pass
        try: occ_dep+=float(str(row[dep_idx] or 0).replace(',',''))
        except: pass
        occ_cnt+=1
    for row,o in V:
        sq_idx = o+5 if is_st_format else o+3
        mr_idx = o+6 if is_st_format else o+4
        try: vac_sq+=float(row[sq_idx] or 0)
        except: pass
        try: vac_mr+=float(str(row[mr_idx] or 0).replace(',',''))
        except: pass
    tot_sq=occ_sq+vac_sq; tot_mr=occ_mr+vac_mr; tot_cnt=occ_cnt+vac_cnt

    ws.cell(rn,1).value='Total Market Rent :'; ws.cell(rn,1).font=gfont(bold=True)
    ws.cell(rn,2).value=tot_mr; ws.cell(rn,2).font=gfont(); ws.cell(rn,2).number_format='#,##0.00'
    ws.cell(rn,7).value='Total Potential Rent :'; ws.cell(rn,7).font=gfont(bold=True)
    ws.cell(rn,9).value=occ_tr; ws.cell(rn,9).font=gfont(); ws.cell(rn,9).number_format='#,##0.00'
    rn+=2

    sum_hdrs=['Square\nFootage','Market\nRent','Sub rent','Actual\nRent','Security\nDeposit','Other\nDeposit','# of\nUnits','Occupancy']
    for i,h in enumerate(sum_hdrs,5):
        ws.cell(rn,i).value=h; ws.cell(rn,i).font=gfont(bold=True)
        ws.cell(rn,i).fill=gfill(GRAY_HDR)
        ws.cell(rn,i).alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.row_dimensions[rn].height=28; rn+=1

    tot_pct=f'{occ_cnt/tot_cnt*100:.2f}%' if tot_cnt else '0%'
    vac_pct=f'{vac_cnt/tot_cnt*100:.2f}%' if tot_cnt else '0%'
    for label,sq,mr,sr,tr,dep,cnt,pct in [
        ('Occupied Units',occ_sq,occ_mr,occ_sr,occ_tr,occ_dep,occ_cnt,tot_pct),
        ('Vacant Units',  vac_sq,vac_mr,0,0,0,vac_cnt,vac_pct),
        ('Totals:',       tot_sq,tot_mr,occ_sr,occ_tr,occ_dep,tot_cnt,'100%'),
    ]:
        bold=(label=='Totals:')
        ws.cell(rn,4).value=label; ws.cell(rn,4).font=gfont(bold=bold)
        for ci,val in [(5,sq),(6,mr),(7,sr),(8,tr),(9,dep),(10,0)]:
            ws.cell(rn,ci).value=val or 0; ws.cell(rn,ci).font=gfont(bold=bold)
            ws.cell(rn,ci).number_format='#,##0.00'
            ws.cell(rn,ci).alignment=Alignment(horizontal='right',vertical='center')
        ws.cell(rn,11).value=cnt; ws.cell(rn,11).font=gfont(bold=bold)
        ws.cell(rn,12).value=pct; ws.cell(rn,12).font=gfont(bold=bold)
        ws.cell(rn,12).alignment=Alignment(horizontal='right',vertical='center')
        for ci in range(4,13): ws.cell(rn,ci).border=Border(top=T,bottom=T,left=T,right=T)
        rn+=1

    for col,w in {'A':9,'B':12,'C':13,'D':22,'E':7,'F':11,'G':11,'H':9,'I':10,'J':10,'K':10,'L':10,'M':10,'N':10,'O':10}.items():
        ws.column_dimensions[col].width=w
    for r in range(1,5): ws.row_dimensions[r].height=14
    ws.freeze_panes='A6'
    return ws, len(V), len(O)

def build_weekly_summary(wb_out, wb_ro, date, prop, ua_ws=None, tar_ws=None, sar_ws=None, tar_total=0, sar_total=0, rr_ws=None):
    total_units = get_total_units(prop)
    ws_name=next((n for n in wb_ro.sheetnames if 'weekly summary' in n.lower()),None)
    if not ws_name: return
    ws_src=wb_ro[ws_name]
    src_vals={}
    for row in ws_src.iter_rows(values_only=False):
        for cell in row:
            if cell.value is not None:
                src_vals[(cell.row,cell.column)]=cell.value
    if ws_name in wb_out.sheetnames: del wb_out[ws_name]
    ws=wb_out.create_sheet(ws_name)
    f9=gfont(sz=9); f9b=gfont(bold=True,sz=9)
    NTV_BLUE='FFBDD7EE'; HDR_BLUE='FFB8CCE4'
    def cell(r,c,val=None,bg=None,bold=False,fmt=None,h='left',bdr=None,wrap=False):
        cell=ws.cell(r,c)
        if val is not None: cell.value=val
        if bg: cell.fill=gfill(bg)
        cell.font=gfont(bold=bold,sz=9)
        cell.alignment=Alignment(horizontal=h,vertical='center',wrap_text=wrap)
        if fmt: cell.number_format=fmt
        if bdr: cell.border=bdr
        return cell
    AB=bblack()
    for r in range(1,4):
        for c in range(2,8):
            ws.cell(r,c).fill=gfill(HDR_BLUE); ws.cell(r,c).font=f9b
            ws.cell(r,c).alignment=galign('center')
            ws.cell(r,c).border=Border(top=T if r==1 else None,bottom=T if r==3 else None,left=T if c==2 else None,right=T if c==7 else None)
    ws.cell(1,2).value=src_vals.get((1,2), prop.split('(')[0].strip())
    ws.cell(2,2).value=src_vals.get((2,2),'Occupancy & Delinquency Summary')
    ws.cell(3,2).value=date; ws.cell(3,2).fill=gfill(BLUE_IN)
    for r in range(1,4): ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=7)
    for r in range(4,23):
        for c in range(2,8):
            ws.cell(r,c).border=AB; ws.cell(r,c).font=f9
    cell(5,2,total_units,h='center'); cell(5,3,'=B5/$B$5',fmt='0.00%',h='center'); cell(5,4,'Total Units',h='left')
    occ=[(6,'Subtract',False,'Physically Vacant'),(7,'Add',True,'Applications - Approved @ KG'),(8,'Add',True,'Applications - Pending Not Approved @ KG'),(9,'Add',True,'Applications - Site Processing - Not Sent to KG'),(10,'Subtract',False,'Notices to Vacate Not at Legal'),(11,'Subtract',False,'Notices to Vacate @ Legal'),(12,None,False,'NET LEASED ')]
    for r,lbl,is_blue,desc in occ:
        if lbl: ws.cell(r,1).value=lbl; ws.cell(r,1).font=f9; ws.cell(r,1).alignment=galign('center')
        if r==12: ws.cell(r,2).value='=B5+B6+B7+B8+B9+B10+B11'
        else:
            if is_blue: ws.cell(r,2).fill=gfill(BLUE_IN)
        ws.cell(r,2).font=f9; ws.cell(r,2).alignment=galign('center')
        ws.cell(r,3).value=f'=B{r}/$B$5'; ws.cell(r,3).font=f9; ws.cell(r,3).number_format='0.00%'; ws.cell(r,3).alignment=galign('center')
        ws.cell(r,4).value=desc; ws.cell(r,4).font=f9; ws.cell(r,4).alignment=galign('left')
    for r in [6,10,11]: ws.cell(r,2).fill=gfill(BLUE_IN)
    ws.cell(14,2).fill=gfill(BLUE_IN); ws.cell(14,2).font=f9; ws.cell(14,2).alignment=galign('center'); ws.cell(14,2).border=AB
    ws.cell(14,3).value='=B14/B5'; ws.cell(14,3).font=f9; ws.cell(14,3).number_format='0.00%'; ws.cell(14,3).alignment=galign('center'); ws.cell(14,3).border=AB
    ws.cell(14,4).value='# of tenants owing prev. full month rent, including'; ws.cell(14,4).font=f9; ws.cell(14,4).alignment=galign('left'); ws.cell(14,4).border=AB
    ws.cell(14,5).border=AB; ws.cell(14,5).font=f9
    ws.cell(14,6).font=f9; ws.cell(14,6).alignment=galign('center'); ws.cell(14,6).border=AB
    ws.cell(14,7).value='@ legal'; ws.cell(14,7).font=f9; ws.cell(14,7).border=AB
    ws.cell(16,2).font=f9; ws.cell(16,2).alignment=galign('center'); ws.cell(16,2).border=AB
    ws.cell(16,3).fill=gfill(BLUE_IN); ws.cell(16,3).font=f9; ws.cell(16,3).alignment=galign('center'); ws.cell(16,3).number_format='#,##0_);(#,##0)'; ws.cell(16,3).border=AB
    ws.cell(16,4).value='# Physically Occupied and Total Leased Rent'; ws.cell(16,4).font=f9; ws.cell(16,4).alignment=galign('left'); ws.cell(16,4).border=AB
    for c in [5,6,7]: ws.cell(16,c).border=AB; ws.cell(16,c).font=f9
    AR_FMT='_([$$-409]* #,##0.00_);_([$$-409]* \\(#,##0.00\\);_([$$-409]* "-"??_);_(@_)'
    for r,desc in [(18,'Tenant Accounts Receivable (AR)'),(19,'Subsidy Accounts Receivable (AR) '),(20,'Total  AR')]:
        for c in range(2,8): ws.cell(r,c).border=AB; ws.cell(r,c).font=f9
        if r in [18,19]: ws.cell(r,3).fill=gfill(BLUE_IN)
        ws.cell(r,3).alignment=Alignment(horizontal='center',vertical='center',wrap_text=False); ws.cell(r,3).number_format=AR_FMT
        ws.cell(r,4).value=desc; ws.cell(r,4).alignment=Alignment(horizontal='left',vertical='center',wrap_text=False)
        if r in [18,19]: ws.cell(r,5).value=f'=C{r}/C16'; ws.cell(r,5).number_format='0.00%'; ws.cell(r,5).alignment=Alignment(horizontal='right',vertical='center',wrap_text=False)
        if r==20: ws.cell(r,5).value='=SUM(E18:E19)'; ws.cell(r,5).number_format='0.00%'; ws.cell(r,5).alignment=Alignment(horizontal='right',vertical='center',wrap_text=False)
    ws.cell(20,3).value='=C18+C19'
    ws.cell(22,2).value='* AR to include current month delinquency beginning 10th of each month'
    for c in range(2,8): ws.cell(22,c).border=AB; ws.cell(22,c).font=f9
    ws.cell(25,2).value='NTV'; ws.cell(25,2).font=f9b; ws.cell(25,2).fill=gfill(NTV_BLUE)
    ws.cell(25,2).border=Border(top=T,bottom=T,left=T,right=T); ws.cell(25,3).border=Border(top=T,bottom=T,right=T)
    ws.merge_cells('B25:C25')
    ws.cell(26,2).value='Unit #'; ws.cell(26,2).font=f9b; ws.cell(26,2).fill=gfill(NTV_BLUE)
    ws.cell(26,2).border=Border(bottom=T,left=T,right=T); ws.cell(26,2).alignment=galign('center')
    ws.cell(26,3).value='Move-in year'; ws.cell(26,3).font=f9b; ws.cell(26,3).fill=gfill(NTV_BLUE)
    ws.cell(26,3).border=Border(bottom=T,left=T,right=T); ws.cell(26,3).alignment=galign('center')
    for r in range(27,50):
        v2=src_vals.get((r,2)); v3=src_vals.get((r,3))
        ws.cell(r,2).border=AB; ws.cell(r,2).font=f9
        ws.cell(r,3).border=AB; ws.cell(r,3).font=f9
        if v2: ws.cell(r,2).value=v2; ws.cell(r,2).alignment=galign('left')
        if v3:
            ws.cell(r,3).value=v3; ws.cell(r,3).number_format='MM/DD/YY'
            ws.cell(r,3).alignment=Alignment(horizontal='center',vertical='center',wrap_text=False)
    ws['B3']=date
    occ_count=0
    if ua_ws:
        V=N=kA=kP=sP=leased=0
        for r in range(8,ua_ws.max_row+1):
            st=str(ua_ws.cell(r,1).value or '').strip()
            if st=='Vacant': V+=1
            if st=='Notice': N+=1
            if st in ('Vacant','Notice'):
                if ua_ws.cell(r,5).value: kA+=1
                if ua_ws.cell(r,6).value: kP+=1
                if ua_ws.cell(r,7).value: sP+=1
            if st=='Occupied': occ_count+=1
            if st in ('Occupied','Notice'):
                try: leased+=float(ua_ws.cell(r,9).value or 0)
                except: pass
        ws['B6']=-V; ws['B7']=kA; ws['B8']=kP; ws['B9']=sP; ws['B10']=-N
        ws.cell(16,2).value=occ_count
        if rr_ws:
            rr_leased=0
            for rr_r in range(6,rr_ws.max_row+1):
                rr_name=str(rr_ws.cell(rr_r,4).value or '').strip()
                if 'VACANT' in rr_name.upper() or not rr_name: continue
                if rr_ws.cell(rr_r,1).value is None: continue
                try: rr_leased+=float(rr_ws.cell(rr_r,10).value or 0)
                except: pass
            ws['C16']=rr_leased
        else:
            ws['C16']=leased
    if tar_ws:
        ev=b14=f14=0
        for r in range(7,tar_ws.max_row+1):
            if str(tar_ws.cell(r,1).value or '').lower()=='total': break
            st=str(tar_ws.cell(r,3).value or '').strip().lower()
            if st in ('eviction','past'): ev+=1; f14+=1
            try:
                p31=float(tar_ws.cell(r,7).value or 0); p61=float(tar_ws.cell(r,8).value or 0)
                p90=float(tar_ws.cell(r,9).value or 0); cur=float(tar_ws.cell(r,6).value or 0)
                if p31+p61+p90>0 and cur>0 and p31+p61+p90>=cur: b14+=1
            except: pass
        ws['B11']=-ev; ws['B14']=b14; ws['F14']=f14
    def getT(aw):
        if not aw: return 0
        for r in range(7,aw.max_row+1):
            if str(aw.cell(r,1).value or '').lower()=='total':
                try: return float(aw.cell(r,5).value or 0)
                except: return 0
        return 0
    if tar_total==0 and tar_ws: tar_total=getT(tar_ws)
    if sar_total==0 and sar_ws: sar_total=getT(sar_ws)
    ws['C18']=tar_total; ws['C19']=sar_total
    ws.cell(20,3).value=tar_total+sar_total
    if ws['C16'].value and float(ws['C16'].value or 0)>0:
        leased_val=float(ws['C16'].value)
        ws.cell(18,5).value=tar_total/leased_val if leased_val else 0
        ws.cell(19,5).value=sar_total/leased_val if leased_val else 0
        ws.cell(20,5).value=(tar_total+sar_total)/leased_val if leased_val else 0
    for col,w in {'A':16.57,'B':14.43,'C':15.43,'D':41.43,'E':9.29,'F':6.0,'G':7.86}.items():
        ws.column_dimensions[col].width=w
    for r,h in [(1,12),(2,12),(3,12),(4,12),(5,12),(6,12),(7,15.75),(8,15.75),(9,15.75),(10,15.75),(11,15.75),(12,15.75),(13,15.75),(14,12),(15,15.75),(16,16.5),(17,15.75),(18,12),(19,15.75),(20,12),(21,12),(22,12),(25,16.5)]:
        ws.row_dimensions[r].height=h

@app.route('/health')
def health():
    return jsonify({'status':'ok','version':'9.8'})

@app.route('/')
def index():
    return render_template_string(PAGE)

@app.route('/format', methods=['POST'])
def format_report():
    try:
        date=request.form.get('date','').strip()
        prop=request.form.get('prop','').strip()
        if not date or not prop:
            return jsonify({'error':'Missing date or property'}),400
        wb_file=request.files.get('wb')
        if not wb_file:
            return jsonify({'error':'Working workbook is required'}),400
        wb_bytes=wb_file.read()
        wb_ro=openpyxl.load_workbook(io.BytesIO(wb_bytes),data_only=True,keep_vba=False,read_only=True)
        pTAR=get_notes(wb_ro,'Tenant AR')
        pSAR=get_notes(wb_ro,'Sub AR')
        pSAR.update(get_notes(wb_ro,'SUB AR'))
        wb_out=openpyxl.Workbook(); wb_out.remove(wb_out.active)
        ua_ws=tar_ws=sar_ws=None
        tar_total=sar_total=0
        ua_f=request.files.get('ua')
        if ua_f: ua_ws,*_=fmt_ua(wb_out,ua_f.read(),date,prop)
        tar_f=request.files.get('tar')
        if tar_f:
            tar_ws,ev,cu,no,pos_end,data_start=fmt_ar(wb_out,tar_f.read(),date,pTAR,False)
            for r in range(data_start,pos_end+1):
                try: tar_total+=float(tar_ws.cell(r,5).value or 0)
                except: pass
        sar_f=request.files.get('sar')
        if sar_f:
            sar_ws,ev,cu,no,pos_end,data_start=fmt_ar(wb_out,sar_f.read(),date,pSAR,True)
            for r in range(data_start,pos_end+1):
                try: sar_total+=float(sar_ws.cell(r,5).value or 0)
                except: pass
        rr_ws=None
        rr_f=request.files.get('rr')
        if rr_f: rr_ws,*_=fmt_rr(wb_out,rr_f.read(),date,prop)
        build_weekly_summary(wb_out,wb_ro,date,prop,ua_ws,tar_ws,sar_ws,tar_total,sar_total,rr_ws)
        wb_ro.close()
        def find_tab(names, prefix):
            return next((n for n in names if n.strip().lower().startswith(prefix.lower())), None)
        current=list(wb_out.sheetnames)
        ws_tab=find_tab(current,'weekly summary'); ua_tab=find_tab(current,'unit availability')
        rr_tab=find_tab(current,'rent roll'); tar_tab=find_tab(current,'tenant ar'); sar_tab=find_tab(current,'sub ar')
        desired=[t for t in [ws_tab,ua_tab,rr_tab,tar_tab,sar_tab] if t]
        remaining=[t for t in current if t not in desired]
        ordered=desired+remaining
        for target_i,name in enumerate(ordered):
            current=list(wb_out.sheetnames); current_i=current.index(name)
            if current_i!=target_i: wb_out.move_sheet(name,offset=target_i-current_i)
        import zipfile, re as _re
        raw=io.BytesIO(); wb_out.save(raw); raw.seek(0)
        patched=io.BytesIO()
        with zipfile.ZipFile(raw,'r') as zin, zipfile.ZipFile(patched,'w',zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data=zin.read(item)
                if item=='xl/workbook.xml':
                    txt=data.decode('utf-8')
                    txt=_re.sub(r'<calcPr[^/]*/>',"<calcPr calcId=\"191\" refMode=\"A1\"/>",txt)
                    data=txt.encode('utf-8')
                zout.writestr(item,data)
        patched.seek(0)
        prefix=prop.split('(')[0].strip().replace(' ','_')
        fname=f'{prefix}_Weekly_{date.replace(".","")}_Formatted.xlsx'
        return send_file(patched,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',as_attachment=True,download_name=fname)
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error':str(e)}),500

PAGE='''<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Weekly Report Formatter</title>
<style>
@import url("https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Sora:wght@300;400;600;700&display=swap");
:root{--g:#7AD694;--gd:#4fa868;--bg:#f5f6f8;--card:#fff;--bdr:#e0e2e8;--ink:#1a1c24;--mut:#7a7e90;}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:"Sora",sans-serif;background:var(--bg);color:var(--ink);}
.hdr{background:linear-gradient(135deg,#1a2e1e,#2c4a31);padding:24px 40px;display:flex;align-items:center;gap:16px;border-bottom:3px solid var(--g);}
.hi{width:40px;height:40px;background:var(--g);border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:20px;}
.hdr h1{font-size:18px;font-weight:700;color:#fff;}
.hdr p{font-size:11px;color:rgba(255,255,255,.45);font-family:"DM Mono",monospace;margin-top:2px;}
.hv{margin-left:auto;background:rgba(122,214,148,.15);border:1px solid rgba(122,214,148,.3);color:var(--g);padding:4px 11px;border-radius:20px;font-size:11px;font-family:"DM Mono",monospace;}
.main{max-width:800px;margin:0 auto;padding:28px 20px 60px;}
.card{background:var(--card);border:1px solid var(--bdr);border-radius:13px;padding:20px 24px;margin-bottom:14px;position:relative;}
.sn{position:absolute;top:-1px;left:20px;background:var(--g);color:#1a2e1e;font-size:10px;font-weight:700;font-family:"DM Mono",monospace;padding:3px 10px;border-radius:0 0 7px 7px;letter-spacing:1px;}
.ct{font-size:14px;font-weight:700;margin:6px 0 4px;}
.cd{font-size:12px;color:var(--mut);margin-bottom:14px;line-height:1.6;}
select,input[type=text]{font-family:"DM Mono",monospace;font-size:13px;border:1.5px solid var(--bdr);border-radius:8px;padding:8px 12px;background:#fff;color:var(--ink);transition:border-color .2s;outline:none;}
select:focus,input:focus{border-color:var(--g);}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;}
.slot{border:1.5px dashed var(--bdr);border-radius:10px;padding:12px 14px;background:#fafbfc;position:relative;transition:all .2s;cursor:pointer;}
.slot:hover{border-color:var(--g);background:rgba(122,214,148,.04);}
.slot.on{border-style:solid;border-color:var(--gd);background:#f0fbf3;}
.slot.full{grid-column:1/-1;}
.slot input[type=file]{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%;}
.sh{display:flex;align-items:center;gap:7px;margin-bottom:3px;}
.dot{width:8px;height:8px;border-radius:50%;flex-shrink:0;}
.sl{font-size:12px;font-weight:700;}
.ss{font-size:10.5px;color:var(--mut);line-height:1.4;}
.sn2{font-size:10.5px;color:var(--gd);font-family:"DM Mono",monospace;margin-top:4px;min-height:14px;}
.btn{display:block;width:100%;background:linear-gradient(135deg,#2c4a31,#1f3824);color:#fff;border:none;border-radius:10px;padding:14px;font-family:"Sora",sans-serif;font-size:14px;font-weight:700;cursor:pointer;transition:all .2s;margin-top:6px;}
.btn:hover:not(:disabled){transform:translateY(-1px);box-shadow:0 8px 24px rgba(44,74,49,.35);}
.btn:disabled{opacity:.5;cursor:not-allowed;}
.pg{height:4px;background:var(--bdr);border-radius:2px;overflow:hidden;margin-top:10px;display:none;}
.pg.on{display:block;}
.pf{height:100%;background:var(--g);border-radius:2px;transition:width .4s;}
.log{background:#1a1c24;border-radius:9px;padding:12px;font-family:"DM Mono",monospace;font-size:11px;color:#8aff9a;max-height:180px;overflow-y:auto;display:none;line-height:1.7;margin-top:10px;}
.log.on{display:block;}
.le{color:#f28e86;}.lw{color:#fdd868;}.li{color:#8cb5f9;}
.dl{display:none;background:#f0fbf3;border:2px solid var(--g);border-radius:11px;padding:18px;text-align:center;margin-top:12px;}
.dl.on{display:block;}
.dl h3{font-size:14px;font-weight:700;color:var(--gd);margin-bottom:5px;}
.dl p{font-size:12px;color:var(--mut);margin-bottom:12px;}
.dlb{display:inline-flex;align-items:center;gap:7px;background:var(--gd);color:#fff;border:none;border-radius:8px;padding:10px 22px;font-family:"Sora",sans-serif;font-size:13px;font-weight:700;cursor:pointer;text-decoration:none;}
.dlb:hover{background:#3d8a53;}
@media(max-width:600px){.hdr{padding:16px;}.main{padding:16px 12px 50px;}.grid{grid-template-columns:1fr;}.slot.full{grid-column:1;}}
</style></head><body>
<div class="hdr"><div class="hi">&#127970;</div><div><h1>Weekly Report Formatter</h1><p>Occupancy &amp; Delinquency &middot; FPI Management</p></div><div class="hv">v9.8</div></div>
<div class="main">
  <div class="card"><div class="sn">STEP 01</div><div class="ct">Select Property &amp; Enter Date</div><div class="cd">Choose the property and enter this week\'s report date.</div>
    <select id="prop" style="width:100%;margin-bottom:10px;"><option value="Village at Madrone (fka Village at Morgan Hill) (x93)">Village at Madrone (x93)</option><option value="Village at First">Village at First</option><option value="Village at Santa Teresa">Village at Santa Teresa</option></select>
    <div style="display:flex;align-items:center;gap:10px;"><input type="text" id="date" placeholder="04.06.26" maxlength="8" style="width:120px;"/><span style="font-size:11px;color:var(--mut);font-family:\'DM Mono\',monospace;">MM.DD.YY</span></div>
  </div>
  <div class="card"><div class="sn">STEP 02</div><div class="ct">Upload Working Workbook</div><div class="cd">The master workbook with Weekly Summary and prior AR history.</div>
    <div class="grid"><div class="slot full" id="s-wb"><input type="file" id="f-wb" accept=".xlsx,.xls,.xlsm"/><div class="sh"><div class="dot" style="background:#8CB5F9;"></div><span class="sl">&#128210; Weekly Workbook</span></div><div class="ss">Master file — Weekly Summary + history</div><div class="sn2" id="n-wb">Click or drag file here</div></div></div>
  </div>
  <div class="card"><div class="sn">STEP 03</div><div class="ct">Upload Yardi Exports</div><div class="cd">Upload each Yardi export. Leave empty anything you don\'t have — it will be skipped.</div>
    <div class="grid">
      <div class="slot" id="s-ua"><input type="file" id="f-ua" accept=".xlsx,.xls,.xlsm"/><div class="sh"><div class="dot" style="background:#7AD694;"></div><span class="sl">Unit Availability</span></div><div class="ss">Onsite &rarr; Analytics &rarr; Unit Availability Details</div><div class="sn2" id="n-ua">Click or drag file here</div></div>
      <div class="slot" id="s-tar"><input type="file" id="f-tar" accept=".xlsx,.xls,.xlsm"/><div class="sh"><div class="dot" style="background:#F28E86;"></div><span class="sl">Tenant AR</span></div><div class="ss">Analytics &rarr; Receivable Aging (Excl. HUD)</div><div class="sn2" id="n-tar">Click or drag file here</div></div>
      <div class="slot" id="s-sar"><input type="file" id="f-sar" accept=".xlsx,.xls,.xlsm"/><div class="sh"><div class="dot" style="background:#8CB5F9;"></div><span class="sl">Subsidy AR</span></div><div class="ss">Analytics &rarr; Receivable Aging (HUD Only)</div><div class="sn2" id="n-sar">Click or drag file here</div></div>
      <div class="slot" id="s-rr"><input type="file" id="f-rr" accept=".xlsx,.xls,.xlsm"/><div class="sh"><div class="dot" style="background:#C4A0F5;"></div><span class="sl">Rent Roll</span></div><div class="ss">Onsite &rarr; Analytics &rarr; Rent Roll</div><div class="sn2" id="n-rr">Click or drag file here</div></div>
    </div>
  </div>
  <div class="card"><div class="sn">STEP 04</div><div class="ct">Format &amp; Download</div><div class="cd">Formats all reports with exact colors, structure, and formulas. Downloads the final workbook.</div>
    <button class="btn" id="btn" onclick="run()">&#9889; Format Report</button>
    <div class="pg" id="pg"><div class="pf" id="pf" style="width:0%"></div></div>
    <div class="log" id="log"></div>
    <div class="dl" id="dl"><h3>&#10003; Done!</h3><p>Your formatted workbook is ready.</p><a class="dlb" id="dlb" href="#">&#8595; Download Formatted Workbook</a></div>
  </div>
</div>
<script>
["wb","ua","tar","sar","rr"].forEach(k=>{
  document.getElementById("f-"+k).addEventListener("change",function(){
    if(this.files[0]){document.getElementById("s-"+k).classList.add("on");document.getElementById("n-"+k).textContent="✓ "+this.files[0].name;}
  });
});
function L(m,c=""){const el=document.getElementById("log");el.classList.add("on");const d=document.createElement("div");if(c)d.className=c;d.textContent="> "+m;el.appendChild(d);el.scrollTop=el.scrollHeight;}
function P(p){document.getElementById("pg").classList.add("on");document.getElementById("pf").style.width=p+"%";}
async function run(){
  const btn=document.getElementById("btn");btn.disabled=true;
  document.getElementById("log").innerHTML="";document.getElementById("log").classList.remove("on");
  document.getElementById("dl").classList.remove("on");document.getElementById("pg").classList.remove("on");
  const date=document.getElementById("date").value.trim();const prop=document.getElementById("prop").value;
  if(!date){alert("Please enter the report date.");btn.disabled=false;return;}
  if(!document.getElementById("f-wb").files[0]){alert("Please upload the working workbook.");btn.disabled=false;return;}
  const form=new FormData();
  form.append("date",date);form.append("prop",prop);
  ["wb","ua","tar","sar","rr"].forEach(k=>{const f=document.getElementById("f-"+k).files[0];if(f)form.append(k,f);});
  L("Uploading and formatting...");P(20);
  try{
    const resp=await fetch("/format",{method:"POST",body:form});
    P(80);
    if(!resp.ok){
      let msg="Server error";
      try{const e=await resp.json();msg=e.error||msg;}catch(ex){}
      L("Error: "+msg,"le");btn.disabled=false;return;
    }
    const blob=await resp.blob();P(100);
    L("Done!","li");
    const url=URL.createObjectURL(blob);
    const prefix=prop.split("(")[0].trim().replace(/ /g,"_");
    const fname=prefix+"_Weekly_"+date.replace(/\\./g,"")+"_Formatted.xlsx";
    const a=document.getElementById("dlb");a.href=url;a.download=fname;
    document.getElementById("dl").classList.add("on");
  }catch(e){L("Error: "+e.message,"le");}
  btn.disabled=false;
}
</script></body></html>'''

if __name__=='__main__':
    port=int(os.environ.get('PORT',5000))
    app.run(host='0.0.0.0',port=port,debug=False)
