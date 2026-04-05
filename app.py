#!/usr/bin/env python3
"""Weekly Report Formatter v8 — Production"""
from flask import Flask, request, send_file, render_template_string, jsonify
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy
import io, re, os, traceback

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

GREEN='FF7AD694'; GRAY_HDR='FFBFBFBF'; GRAY_AR='FFD9D9D9'
KG_RED='FFF28E86'; KG_YEL='FFFDD868'; KG_BLUE='FF8CB5F9'
WHITE='FFFFFFFF'; BLACK='FF000000'; DARKGRAY='FF505050'; RED_FONT='FFFF0000'

def gfill(h): return PatternFill(fill_type='solid', fgColor=h)
def gfont(bold=False,sz=9,color='FF000000'): return Font(name='Calibri',size=sz,bold=bold,color=color)
def galign(h='left',v='center',wrap=False): return Alignment(horizontal=h,vertical=v,wrap_text=wrap)
T=Side(style='thin',color='FF000000'); TG=Side(style='thin',color='FFCCCCCC')
def bblack(): return Border(top=T,bottom=T,left=T,right=T)
def bgray():  return Border(top=TG,bottom=TG,left=TG,right=TG)

def parse_ua(ws):
    out=[]; status='Occupied'
    for row in ws.iter_rows(values_only=True):
        c0=str(row[0] or '').strip()
        if not c0: continue
        lo=c0.lower()
        if 'village at' in lo and not (row[1] or '') and not (row[2] or ''):
            if   re.search(r'- occupied',lo): status='Occupied'
            elif re.search(r'- vacant',lo):   status='Vacant'
            elif re.search(r'- notice',lo):   status='Notice'
            continue
        if c0 in ('Unit','Total') or re.match(r'^(Unit Availability|Showing|Group|As Of)',c0): continue
        if not re.match(r'^\d{2}-\d{3}',c0): continue
        out.append({'status':status,'unit':c0,'res_id':str(row[1] or '').strip(),
            'name':str(row[2] or '').strip(),'res_rent':row[3],'unit_rent':row[4],
            'res_dep':row[5],'unit_dep':row[6],'yardi_st':str(row[7] or '').strip(),
            'days':row[8],'make_rdy':row[9],'move_in':row[10],'hold':str(row[11] or '').strip(),
            'hold_until':row[12],'notice':row[13],'move_out':row[14],
            'lease_sgn':row[15],'lease_from':row[16],'lease_to':row[17]})
    return out

def fmt_ua(wb_out, raw_bytes, date, prop):
    wb_r=openpyxl.load_workbook(io.BytesIO(raw_bytes),data_only=True,keep_vba=False)
    data=parse_ua(wb_r.active)
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
        for c in range(1,MC+1):
            ws.cell(r,c).fill=gfill(GREEN); ws.cell(r,c).font=gfont(bold=bold); ws.cell(r,c).alignment=galign()
        ws.cell(r,1).value=text
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=MC)
    h6=[None,'Unit','Resident','Name','KG Approved','KG Pend','Site Pending','Resident','Unit','Resident','Unit','Status','Days','Make','Move','Hold','Notice','Move','Lease','Lease','Lease']
    h7=[' ',' ',' ',' ',' ',' ',' ','Rent','Rent','Deposit','Deposit',' ','Vacant','Ready','In',' ',' ','Out','Sign','From','To']
    for c in range(1,MC+1):
        for r,hdr in [(6,h6),(7,h7)]:
            cell=ws.cell(r,c); cell.value=hdr[c-1] if hdr[c-1] is not None else ''
            cell.font=gfont(bold=True); cell.fill=gfill(GRAY_HDR); cell.alignment=galign('center')
    for c in range(1,MC+1): ws.cell(7,c).border=Border(bottom=T)
    rn=8; blank=False
    for p in V+N+O:
        isVN=p['status'] in ('Vacant','Notice')
        if p['status']=='Occupied' and not blank:
            for c in range(1,MC+1): ws.cell(rn,c).fill=gfill(WHITE)
            ws.row_dimensions[rn].height=15.0; rn+=1; blank=True
        def sc(col,val,bg=WHITE,h='left'):
            cell=ws.cell(rn,col); cell.value=val; cell.font=gfont(); cell.fill=gfill(bg); cell.alignment=galign(h)
        sc(1,p['status']); sc(2,p['unit']); sc(3,p['res_id'] or None)
        sc(4,None if p['status']=='Vacant' else (p['name'] or None))
        sc(5,None,KG_RED if isVN else WHITE); sc(6,None,KG_YEL if isVN else WHITE); sc(7,None,KG_BLUE if isVN else WHITE)
        sc(8,p['res_rent'] if p['res_rent'] is not None else 0,h='right')
        sc(9,p['unit_rent'] if p['unit_rent'] is not None else 0,h='right')
        sc(10,p['res_dep'] if p['res_dep'] is not None else 0,h='right')
        sc(11,p['unit_dep'] if p['unit_dep'] is not None else 0,h='right')
        sc(12,p['yardi_st'] or None); sc(13,p['days'] if p['days'] is not None else None,h='right')
        sc(14,p['make_rdy'] if p['make_rdy'] is not None else None,h='right')
        sc(15,p['move_in'] if p['move_in'] is not None else None,h='right')
        sc(16,p['hold'] or None); sc(17,p['hold_until'] if p['hold_until'] is not None else None)
        sc(18,p['move_out'] if p['move_out'] is not None else None)
        sc(19,p['lease_sgn'] if p['lease_sgn'] is not None else None)
        sc(20,p['lease_from'] if p['lease_from'] is not None else None)
        sc(21,p['lease_to'] if p['lease_to'] is not None else None)
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
    rr=list(openpyxl.load_workbook(io.BytesIO(raw_bytes),data_only=True,keep_vba=False).active.iter_rows(values_only=True))
    NL='Comments' if is_sub else 'Delinquency notes'; tc=BLACK if is_sub else DARKGRAY
    tab=f'{"SUB AR" if is_sub else "Tenant AR"} {date}'
    if tab in wb_out.sheetnames: del wb_out[tab]
    ws=wb_out.create_sheet(tab); MC=13
    for ti in range(3):
        for c in range(1,MC+1): ws.cell(ti+1,c).fill=gfill(GREEN); ws.cell(ti+1,c).font=gfont(color=tc); ws.cell(ti+1,c).alignment=galign('center')
        ws.cell(ti+1,1).value=str(rr[ti][0] if ti<len(rr) and rr[ti] else '')
    h4=['','','','','Total','','','','','','','',NL]; h5=['','','','','Unpaid','0-30','31-60','61-90','Over 90','','','','']
    h6=['Unit','Resident','Status','Name','Charges','days','days','days','days','Prepays','Suspense','Balance',NL]
    bd=bblack()
    for c in range(1,MC+1):
        ra=5<=c<=12
        for ri,hdr in [(4,h4),(5,h5),(6,h6)]:
            cell=ws.cell(ri,c); cell.value=hdr[c-1]; cell.font=gfont(bold=True); cell.fill=gfill(GRAY_AR); cell.border=bd
            cell.alignment=galign('right' if ra else ('center' if c==13 else 'left'))
    for r in range(1,4): ws.cell(r,13).fill=gfill(GREEN); ws.cell(r,13).font=gfont(color=tc)
    hi=next((i for i,r in enumerate(rr[:10]) if r and any(str(c or '').lower()=='unit' for c in r) and any(str(c or '').lower()=='resident' for c in r)),5)
    ev,cu,no=[],[],[]
    for row in rr[hi+1:]:
        if not row or all(c is None or c=='' for c in row): continue
        st=str(row[2] or '').strip().lower(); u=str(row[0] or '').strip()
        if re.search(r'subtotal|village at',u,re.I): continue
        if not u or not re.match(r'^\d{2}',u): continue
        if st in ('eviction','past'): ev.append(row)
        elif st=='notice': no.append(row)
        else: cu.append(row)
    def sk(r):
        try: return -(float(str(r[4] or 0).replace(',','')))
        except: return 0
    ev.sort(key=sk); cu.sort(key=sk); no.sort(key=sk)
    rn=7; gb=bgray()
    for row in ev+cu+no:
        rid=str(row[1] or '').strip(); note=prev_notes.get(rid,'')
        st=str(row[2] or '').strip().lower()
        rc=RED_FONT if (not is_sub and st in ('notice','eviction','past')) else BLACK
        for c in range(1,13):
            v=row[c-1]; sv=str(v if v is not None else '').strip()
            try: num=float(sv.replace(',',''))
            except: num=None
            isn=5<=c<=12 and num is not None and sv!=''
            cell=ws.cell(rn,c); cell.value=num if isn else (sv or None)
            cell.font=gfont(color=rc); cell.fill=gfill(WHITE); cell.alignment=galign('right' if c>=5 else 'left'); cell.border=gb
            if isn: cell.number_format='#,##0.00'
        nc=ws.cell(rn,13); nc.value=note or None; nc.font=gfont(color=BLACK); nc.fill=gfill(WHITE); nc.alignment=galign('left',wrap=True); nc.border=gb; nc.number_format='@'
        rn+=1
    tb=bblack()
    ws.cell(rn,1).value='Total'; ws.cell(rn,1).font=gfont(bold=True); ws.cell(rn,1).fill=gfill(WHITE); ws.cell(rn,1).border=tb
    for c in range(5,13):
        cell=ws.cell(rn,c); cell.value=f'=SUM({get_column_letter(c)}7:{get_column_letter(c)}{rn-1})'
        cell.font=gfont(bold=True); cell.fill=gfill(WHITE); cell.border=tb; cell.number_format='#,##0.00'
    for c in [2,3,4,13]: ws.cell(rn,c).fill=gfill(WHITE); ws.cell(rn,c).font=gfont(bold=True); ws.cell(rn,c).border=tb
    for i,w in enumerate([9,13,10,24,12,10,10,10,10,10,10,12,38],1): ws.column_dimensions[get_column_letter(i)].width=w
    ws.freeze_panes='A7'
    return ws, len(ev), len(cu), len(no)

def fmt_rr(wb_out, raw_bytes, date, prop):
    rr=list(openpyxl.load_workbook(io.BytesIO(raw_bytes),data_only=True,keep_vba=False).active.iter_rows(values_only=True))
    tab=f'Rent Roll {date}'
    if tab in wb_out.sheetnames: del wb_out[tab]
    ws=wb_out.create_sheet(tab); MC=14
    for ti,(text,bold) in enumerate(zip(['Rent Roll',prop,f'As Of = {date}','Month Year'],[True,False,True,False])):
        for c in range(1,MC+1): ws.cell(ti+1,c).fill=gfill(GREEN); ws.cell(ti+1,c).font=gfont(bold=bold)
        ws.cell(ti+1,1).value=text
    h5=['Unit','Unit Type','Unit','Resident','Name','Market','Actual','Resident','Other','Move In','Lease','Move Out','Balance','Comments']
    h6=['\xa0','\xa0','Sq Ft','\xa0','\xa0','Rent','Rent','Deposit','Deposit','\xa0','Expiration','\xa0','\xa0','']
    for c in range(1,MC+1):
        c5=ws.cell(5,c); c5.value=h5[c-1]; c5.font=gfont(bold=True); c5.fill=gfill(GRAY_HDR); c5.alignment=galign('center'); c5.border=Border(bottom=T)
        c6=ws.cell(6,c); c6.value=h6[c-1]; c6.font=gfont(bold=True); c6.fill=gfill(GRAY_HDR); c6.alignment=galign('center'); c6.border=Border(top=T,bottom=T)
    hi=next((i for i,r in enumerate(rr[:10]) if r and any(str(c or '').lower()=='unit' for c in r) and any('type' in str(c or '').lower() for c in r)),4)
    V,O=[],[]
    for row in rr[hi+1:]:
        if not row or all(c is None or c=='' for c in row): continue
        unit=str(row[1] or '').strip()
        if not re.match(r'^\d{2}-\d{3}',unit): continue
        rname=str(row[3] or '').strip().upper()
        if 'VACANT' in rname or not rname: V.append(row)
        else: O.append(row)
    V.sort(key=lambda r:str(r[1] or '')); O.sort(key=lambda r:str(r[1] or ''))
    rn=7
    for rows,fc in [(V,RED_FONT),(O,BLACK)]:
        for row in rows:
            unit=str(row[1] or '').strip(); ut=str(row[2] or '').strip()
            rname=str(row[3] or '').strip(); sq=row[4]; mr=row[5]; tr=row[8]; dep=row[11]; mi=row[12]; lt=row[14]
            isvac='VACANT' in rname.upper() or not rname
            def sc(col,val,h='left',fmt=None):
                cell=ws.cell(rn,col); cell.value=val; cell.font=gfont(color=fc); cell.fill=gfill(WHITE); cell.alignment=galign(h)
                if fmt: cell.number_format=fmt
            sc(1,unit); sc(2,ut)
            ws.cell(rn,3).value=sq or 0; ws.cell(rn,3).font=gfont(color=fc); ws.cell(rn,3).fill=gfill(WHITE); ws.cell(rn,3).alignment=galign('right'); ws.cell(rn,3).number_format='#,##0'
            sc(4,'VACANT' if isvac else rname); sc(5,'VACANT' if isvac else rname)
            for col,val in [(6,mr),(7,tr),(8,dep)]:
                ws.cell(rn,col).value=val or 0; ws.cell(rn,col).font=gfont(color=fc); ws.cell(rn,col).fill=gfill(WHITE); ws.cell(rn,col).alignment=galign('right'); ws.cell(rn,col).number_format='#,##0.00'
            sc(9,0,h='right')
            for col,dv in [(10,mi),(11,lt),(12,None)]:
                cell=ws.cell(rn,col); cell.value=dv; cell.font=gfont(color=fc); cell.fill=gfill(WHITE); cell.alignment=galign('center')
                if isinstance(dv,(int,float)) and dv: cell.number_format='m/d/yyyy'
            ws.cell(rn,13).value=0; ws.cell(rn,13).font=gfont(color=fc); ws.cell(rn,13).fill=gfill(WHITE); ws.cell(rn,13).alignment=galign('right'); ws.cell(rn,13).number_format='#,##0.00'
            sc(14,'')
            ws.row_dimensions[rn].height=15.0; rn+=1
    for col,w in {'A':9,'B':12,'C':7,'D':13,'E':22,'F':11,'G':11,'H':11,'I':11,'J':11,'K':13,'L':11,'M':10,'N':49}.items():
        ws.column_dimensions[col].width=w
    ws.freeze_panes='A7'
    return ws, len(V), len(O)

def copy_ws(wb_src, wb_out, date, ua_ws=None, tar_ws=None, sar_ws=None):
    ws_name=next((n for n in wb_src.sheetnames if 'weekly summary' in n.lower()),None)
    if not ws_name: return
    ws_s=wb_src[ws_name]
    if ws_name in wb_out.sheetnames: del wb_out[ws_name]
    ws=wb_out.create_sheet(ws_name)
    for row in ws_s.iter_rows():
        for cell in row:
            nc=ws.cell(row=cell.row,column=cell.column,value=cell.value)
            if cell.has_style:
                nc.font=copy(cell.font); nc.fill=copy(cell.fill)
                nc.border=copy(cell.border); nc.alignment=copy(cell.alignment)
                nc.number_format=cell.number_format
    for m in ws_s.merged_cells.ranges: ws.merge_cells(str(m))
    for col in ws_s.column_dimensions: ws.column_dimensions[col].width=ws_s.column_dimensions[col].width
    for r in ws_s.row_dimensions: ws.row_dimensions[r].height=ws_s.row_dimensions[r].height
    ws['B3']=date
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
            if st in ('Occupied','Notice'):
                try: leased+=float(ua_ws.cell(r,9).value or 0)
                except: pass
        ws['B6']=-V; ws['B7']=kA; ws['B8']=kP; ws['B9']=sP; ws['B10']=-N; ws['C16']=leased
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
    ws['C18']=getT(tar_ws); ws['C19']=getT(sar_ws)

@app.route('/health')
def health():
    return jsonify({'status':'ok','version':'8.0'})

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
        wb_src=openpyxl.load_workbook(io.BytesIO(wb_file.read()),data_only=True,keep_vba=False)
        pTAR=get_notes(wb_src,'Tenant AR')
        pSAR=get_notes(wb_src,'Sub AR')
        pSAR.update(get_notes(wb_src,'SUB AR'))
        wb_out=openpyxl.Workbook(); wb_out.remove(wb_out.active)
        ua_ws=tar_ws=sar_ws=None
        ua_f=request.files.get('ua')
        if ua_f: ua_ws,*_=fmt_ua(wb_out,ua_f.read(),date,prop)
        tar_f=request.files.get('tar')
        if tar_f: tar_ws,*_=fmt_ar(wb_out,tar_f.read(),date,pTAR,False)
        sar_f=request.files.get('sar')
        if sar_f: sar_ws,*_=fmt_ar(wb_out,sar_f.read(),date,pSAR,True)
        rr_f=request.files.get('rr')
        if rr_f: fmt_rr(wb_out,rr_f.read(),date,prop)
        copy_ws(wb_src,wb_out,date,ua_ws,tar_ws,sar_ws)
        out=io.BytesIO(); wb_out.save(out); out.seek(0)
        prefix=prop.split('(')[0].strip().replace(' ','_')
        fname=f'{prefix}_Weekly_{date.replace(".","")}_Formatted.xlsx'
        return send_file(out,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',as_attachment=True,download_name=fname)
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
<div class="hdr"><div class="hi">&#127970;</div><div><h1>Weekly Report Formatter</h1><p>Occupancy &amp; Delinquency &middot; FPI Management</p></div><div class="hv">v8.0</div></div>
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
    <div class="dl" id="dl"><h3>&#10003; Done!</h3><p>Your formatted workbook is ready to open in Excel.</p><a class="dlb" id="dlb" href="#">&#8595; Download Formatted Workbook</a></div>
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
      try{const e=await resp.json();msg=e.error||msg;}catch(e){}
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

