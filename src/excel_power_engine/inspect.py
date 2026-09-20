from __future__ import annotations
import hashlib, re, xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile
from .models import WorkbookProfile

NS_MAIN="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_DOC="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL="http://schemas.openxmlformats.org/package/2006/relationships"

def _sha256(data: bytes)->str: return hashlib.sha256(data).hexdigest()
def _normalize_part_target(target: str, base: str="xl")->str:
    target=target.replace("\\", "/")
    if target.startswith("/"): return target.lstrip("/")
    if target.startswith("xl/"): return target
    return f"{base}/{target}"

def _read_rels(zf: ZipFile, rels_part: str)->dict[str,str]:
    if rels_part not in zf.namelist(): return {}
    root=ET.fromstring(zf.read(rels_part))
    return {r.get("Id",""):r.get("Target","") for r in root.findall(f"{{{NS_REL}}}Relationship")}

def _sheet_stats(xml: bytes)->dict[str,object]:
    text=xml.decode("utf-8",errors="ignore")
    dim=re.search(r'<dimension\b[^>]*ref="([^"]+)"', text)
    return {
        "rows":len(re.findall(r"<row\b",text)),
        "cells":len(re.findall(r"<c\b",text)),
        "formulas":len(re.findall(r"<f(?:\s|>)",text)),
        "merged_ranges":len(re.findall(r"<mergeCell\b",text)),
        "conditional_formatting":len(re.findall(r"<conditionalFormatting\b",text)),
        "data_validations":len(re.findall(r"<dataValidation\b",text)),
        "x14_data_validations":len(re.findall(r"<x14:dataValidation\b",text)),
        "sheet_protection":"<sheetProtection" in text,
        "auto_filter":"<autoFilter" in text,
        "dimension":dim.group(1) if dim else None,
    }

def inspect_workbook(path: str|Path, deep: bool=False)->WorkbookProfile:
    p=Path(path); suffix=p.suffix.lower()
    profile=WorkbookProfile(p,suffix,suffix in {".xlsx",".xlsm",".xltx",".xltm"},suffix==".xlsm")
    if not profile.is_ooxml: profile.features.add("non-ooxml"); return profile
    with ZipFile(p) as zf:
        profile.internal_parts=zf.namelist(); profile.part_sizes={n:zf.getinfo(n).file_size for n in profile.internal_parts}
        profile.has_vba="xl/vbaProject.bin" in profile.internal_parts
        if profile.has_vba: profile.features.add("vba")
        names=[n.lower() for n in profile.internal_parts]
        rules=[
            ("customxml","customXml"),("pivot","pivot"),("drawing|chart","drawings/charts"),
            ("externallink","external-links"),("tables/","tables"),("slicer","slicers"),
            ("vml","vml"),("ctrlprop|activex|controls","controls"),
            ("comment|threadedcomment","comments"),("connections","connections"),("querytable","query-tables"),
        ]
        for needle,feature in rules:
            if any(re.search(needle,n) for n in names): profile.features.add(feature)
        if "xl/sharedstrings.xml" in names: profile.features.add("shared-strings")
        if "xl/styles.xml" in names: profile.features.add("styles")
        if "xl/workbook.xml" in profile.internal_parts:
            root=ET.fromstring(zf.read("xl/workbook.xml")); rels=_read_rels(zf,"xl/_rels/workbook.xml.rels")
            for sh in root.findall(f"{{{NS_MAIN}}}sheets/{{{NS_MAIN}}}sheet"):
                name=sh.get("name",""); rid=sh.get(f"{{{NS_DOC}}}id",""); profile.sheet_parts[name]=_normalize_part_target(rels.get(rid,""))
            names_parent=root.find(f"{{{NS_MAIN}}}definedNames")
            if names_parent is not None:
                profile.defined_names=[n.get("name","") for n in names_parent]
                if profile.defined_names: profile.features.add("defined-names")
        if deep:
            found_x14 = False
            for name,part in profile.sheet_parts.items():
                if part not in profile.internal_parts: continue
                sheet_bytes = zf.read(part)
                stats=_sheet_stats(sheet_bytes); rels=_read_rels(zf,f"xl/worksheets/_rels/{Path(part).name}.rels")
                if stats.get("x14_data_validations", 0):
                    found_x14 = True
                stats["relationships"]=len(rels)
                stats["table_relationships"]=sum("table" in t.lower() for t in rels.values())
                stats["drawing_relationships"]=sum("drawing" in t.lower() for t in rels.values())
                profile.sheet_profiles[name]=stats
            if found_x14:
                profile.features.add("x14")
            profile.integrity={"part_count":len(profile.internal_parts),"package_sha256":_package_hash(zf),"vba_sha256":_sha256(zf.read("xl/vbaProject.bin")) if profile.has_vba else None}
    return profile

def _package_hash(zf:ZipFile)->str:
    h=hashlib.sha256()
    for name in sorted(zf.namelist()): h.update(name.encode()); h.update(zf.read(name))
    return h.hexdigest()
