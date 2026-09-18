import zipfile, re, os
import xml.etree.ElementTree as ET

# Read template
template_path = "sample-data/Best Project Placeholder.pptx"
with zipfile.ZipFile(template_path, "r") as z:
    all_files = {name: z.read(name) for name in z.namelist()}

pres_xml_str = all_files["ppt/presentation.xml"].decode("utf-8")
pres_rels_str = all_files["ppt/_rels/presentation.xml.rels"].decode("utf-8")
content_types_str = all_files["[Content_Types].xml"].decode("utf-8")
slide1_xml_str = all_files["ppt/slides/slide1.xml"].decode("utf-8")
slide1_rels_bytes = all_files["ppt/slides/_rels/slide1.xml.rels"]

# Slide dimensions
sld_sz_match = re.search(r'<p:sldSz\s+cx="(\d+)"\s+cy="(\d+)"', pres_xml_str)
slide_w = int(sld_sz_match.group(1)) if sld_sz_match else 10693400

students = [
    ("Daniel Seddik", "Advanced Programming Applications"),
    ("Lina Mohamed", "Advanced Artificial Intelligence"),
    ("Abd Allah Hegazi", "Distributed Systems Security"),
    ("Mohamed Medhat Abdelmotal", "Malware Analysis and Reverse Engineering"),
    ("Judy Zein-Eldeen", "Deep Learning")
]

# Find max rId in pres_rels
existing_rids = [int(m) for m in re.findall(r"rId(\d+)", pres_rels_str)]
next_rid = max(existing_rids) + 1

# Parse presentation.xml.rels
root_rels = ET.fromstring(pres_rels_str)
ns_rels = "http://schemas.openxmlformats.org/package/2006/relationships"

# Filter out existing slide relationships
for rel in list(root_rels.findall(f"{{{ns_rels}}}Relationship")):
    if "relationships/slide" in rel.get("Type", ""):
        root_rels.remove(rel)

# Filter out existing slide overrides in [Content_Types].xml
ct_root = ET.fromstring(content_types_str)
ns_ct = "http://schemas.openxmlformats.org/package/2006/content-types"
for ov in list(ct_root.findall(f"{{{ns_ct}}}Override")):
    if ov.get("PartName", "").startswith("/ppt/slides/slide"):
        ct_root.remove(ov)

out_files = dict(all_files)
for k in list(out_files.keys()):
    if k.startswith("ppt/slides/slide") and (k.endswith(".xml") or k.endswith(".xml.rels")):
        del out_files[k]

sld_id_lst_elements = []
sld_id_base = 256

for idx, (st_name, c_name) in enumerate(students, start=1):
    r_id = f"rId{next_rid}"
    next_rid += 1
    sld_id = sld_id_base + idx - 1
    
    # 1. Relationship in presentation.xml.rels
    rel_elem = ET.SubElement(root_rels, f"{{{ns_rels}}}Relationship")
    rel_elem.set("Id", r_id)
    rel_elem.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide")
    rel_elem.set("Target", f"slides/slide{idx}.xml")
    
    # 2. Content types override
    ov_elem = ET.SubElement(ct_root, f"{{{ns_ct}}}Override")
    ov_elem.set("PartName", f"/ppt/slides/slide{idx}.xml")
    ov_elem.set("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")
    
    # 3. Slide XML manipulation
    def xml_esc(s):
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&apos;")
    
    s_xml = slide1_xml_str
    s_xml = s_xml.replace("{Student_Name}", xml_esc(st_name))
    s_xml = s_xml.replace("{Course_Name}", xml_esc(c_name))
    
    if len(st_name) > 17:
        new_w = int(round(slide_w * 0.90))
        new_x = int(round((slide_w - new_w) / 2))
        # Update off x and ext cx in TextBox 11
        pattern = r'(<[a-zA-Z0-9:]*cNvPr[^>]*name="TextBox 11"[^>]*>.*?</[a-zA-Z0-9:]*nvSpPr>\s*<[a-zA-Z0-9:]*spPr>\s*<[a-zA-Z0-9:]*xfrm[^>]*>\s*<[a-zA-Z0-9:]*off x=")\d+(" y="\d+"/>\s*<[a-zA-Z0-9:]*ext cx=")\d+(" cy="\d+"/>)'
        replacement = rf'\g<1>{new_x}\g<2>{new_w}\g<3>'
        s_xml = re.sub(pattern, replacement, s_xml, flags=re.DOTALL)
        
    if len(st_name) > 20:
        new_sz = 4400 if len(st_name) > 24 else 4700
        s_xml = s_xml.replace('sz="5489"', f'sz="{new_sz}"')
        
    out_files[f"ppt/slides/slide{idx}.xml"] = s_xml.encode("utf-8")
    out_files[f"ppt/slides/_rels/slide{idx}.xml.rels"] = slide1_rels_bytes
    sld_id_lst_elements.append(f'<p:sldId id="{sld_id}" r:id="{r_id}"/>')

# Update presentation.xml
joined_sld_ids = "".join(sld_id_lst_elements)
new_sld_id_str = f"<p:sldIdLst>{joined_sld_ids}</p:sldIdLst>"
new_pres_xml = re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>", new_sld_id_str, pres_xml_str)
out_files["ppt/presentation.xml"] = new_pres_xml.encode("utf-8")
out_files["ppt/_rels/presentation.xml.rels"] = ET.tostring(root_rels, encoding="utf-8", xml_declaration=True)
out_files["[Content_Types].xml"] = ET.tostring(ct_root, encoding="utf-8", xml_declaration=True)

# Write output PPTX
out_path = "sample-data/test_output.pptx"
with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
    for name, content in out_files.items():
        zout.writestr(name, content)

print("SUCCESS: test_output.pptx created! Size:", os.path.getsize(out_path))
