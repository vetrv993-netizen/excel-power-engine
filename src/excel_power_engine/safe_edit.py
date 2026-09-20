from __future__ import annotations
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from shutil import copy2
from zipfile import ZipFile, ZIP_DEFLATED
import json, re
from xml.sax.saxutils import escape
from .inspect import inspect_workbook
from .models import SafeEditReport

def _excel_serial(v): return (v.date() if isinstance(v,datetime) else v)-date(1899,12,30)
def _find_cell(row_xml, ref):
    q=re.escape(ref)
    for pat in (rf'<c\b[^>]*\br="{q}"[^>]*/>',rf'<c\b[^>]*\br="{q}"[^>]*>.*?</c>'):
        m=re.search(pat,row_xml,re.S)
        if m:return m

def _attrs(cell_xml):
    end=cell_xml.find(">")
    attrs=cell_xml[:end+1] if end>=0 else cell_xml
    attrs=re.sub(r'\s+t="[^"]*"','',attrs)
    return attrs.replace('/>','>')

def _replace_cell(row_xml, ref, body, t=None):
    m=_find_cell(row_xml,ref)
    if m:
        attrs=_attrs(m.group(0))
        if t: attrs=attrs[:-1]+f' t="{t}">'
        new=attrs+body+'</c>'
        return row_xml[:m.start()]+new+row_xml[m.end():]
    extra=f' t="{t}"' if t else ''
    return row_xml.replace('</row>',f'<c r="{ref}"{extra}>{body}</c></row>',1)

def _set_value(row_xml,ref,v):
    if v is None:return _replace_cell(row_xml,ref,'')
    if isinstance(v,bool):return _replace_cell(row_xml,ref,f'<v>{1 if v else 0}</v>','b')
    if isinstance(v,(datetime,date)):return _replace_cell(row_xml,ref,f'<v>{_excel_serial(v).days}</v>')
    if isinstance(v,(int,float)) and not isinstance(v,bool):return _replace_cell(row_xml,ref,f'<v>{v}</v>')
    return _replace_cell(row_xml,ref,f'<is><t>{escape(str(v))}</t></is>','inlineStr')

def _set_formula(row_xml,ref,formula):return _replace_cell(row_xml,ref,f'<f>{escape(formula.lstrip("="))}</f>')

@dataclass(slots=True)
class EditOperation:
    cell:str
    value:object=None
    formula:str|None=None

class SafeEditor:
    def edit_cells(self,source,sheet,edits,output=None,create_backup=True):
        source=Path(source); output=Path(output) if output else source.with_name(source.stem+'_SAFE_EDIT'+source.suffix)
        backup=source.with_name(source.stem+'_BACKUP'+source.suffix) if create_backup else None
        if source.resolve()==output.resolve():raise ValueError('SafeEdit requires a different output path')
        profile=inspect_workbook(source,deep=True)
        if not profile.is_ooxml:raise ValueError('SafeEdit currently supports OOXML files only')
        if sheet not in profile.sheet_parts:raise KeyError(f'Worksheet not found: {sheet}')
        part=profile.sheet_parts[sheet]
        if backup:copy2(source,backup)
        with ZipFile(source,'r') as zin: original={i.filename:zin.read(i.filename) for i in zin.infolist()}
        target=original[part]; xml=target.decode('utf-8',errors='strict')
        rows={int(m.group(1)):m.group(0) for m in re.finditer(r'<row\b[^>]*\br="(\d+)"[^>]*>.*?</row>',xml,re.S)}
        updates={}; edited=[]
        for op in edits:
            m=re.fullmatch(r'([A-Z]+)(\d+)',op.cell.upper())
            if not m:raise ValueError(f'Invalid cell reference: {op.cell}')
            r=int(m.group(2))
            if r not in rows:raise ValueError(f'Row {r} does not exist in {sheet}')
            cur=updates.get(r,rows[r])
            cur=_set_formula(cur,op.cell.upper(),op.formula) if op.formula is not None else _set_value(cur,op.cell.upper(),op.value)
            updates[r]=cur; edited.append(op.cell.upper())
        outxml=[]; last=0
        for m in re.finditer(r'<row\b[^>]*\br="(\d+)"[^>]*>.*?</row>',xml,re.S):
            r=int(m.group(1))
            if r not in updates:continue
            outxml.extend([xml[last:m.start()],updates[r]]); last=m.end()
        outxml.append(xml[last:]); edited_bytes=''.join(outxml).encode('utf-8')
        if output.exists():output.unlink()
        with ZipFile(output,'w',ZIP_DEFLATED) as zout:
            for name,data in original.items():zout.writestr(name,edited_bytes if name==part else data)
        with ZipFile(output,'r') as zout:new={i.filename:zout.read(i.filename) for i in zout.infolist()}
        changed=[n for n in original if original[n]!=new.get(n,b'')]
        protected=['xl/vbaProject.bin','xl/styles.xml','xl/workbook.xml','xl/_rels/workbook.xml.rels','[Content_Types].xml']
        protected_ok=all(original.get(n)==new.get(n) for n in protected)
        vba_ok=(not profile.has_vba) or original.get('xl/vbaProject.bin')==new.get('xl/vbaProject.bin')
        x14_before=b'x14:dataValidation' in target; x14_after=b'x14:dataValidation' in new.get(part,b'')
        return SafeEditReport(source,output,backup,sheet,edited,changed,protected_ok,vba_ok,(not x14_before) or x14_after)
    @staticmethod
    def save_report(report,path):Path(path).write_text(json.dumps(report.as_dict(),ensure_ascii=False,indent=2),encoding='utf-8')
