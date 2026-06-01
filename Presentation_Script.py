import zipfile
import os
import json
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element, SubElement
from collections import defaultdict

# ----- XML namespaces -----
NSMAP = {
    'a':  'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r':  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p':  'http://schemas.openxmlformats.org/presentationml/2006/main',
    'rel':'http://schemas.openxmlformats.org/package/2006/relationships',
}
for prefix, uri in NSMAP.items():
    ET.register_namespace(prefix, uri)

NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'

DEFAULT_SIZE_EMU = 914400

# ---------- Media helpers ----------
def _next_image_num(zip_ref, prefix='ppt/media/image', ext='.svg'):
    existing = [n for n in zip_ref.namelist() if n.startswith(prefix) and n.endswith(ext)]
    max_num = 0
    for name in existing:
        try:
            num = int(name[len(prefix):-len(ext)])
            if num > max_num:
                max_num = num
        except:
            pass
    return max_num + 1

def _ensure_svg_content_type(root):
    for default in root.findall(f'{{{NS_CT}}}Default'):
        if default.get('Extension') == 'svg':
            return False
    default = Element(f'{{{NS_CT}}}Default', Extension='svg', ContentType='image/svg+xml')
    root.append(default)
    return True

def _next_rel_id(rels_element):
    max_id = 0
    for rel in rels_element.findall(f'{{{NS_REL}}}Relationship'):
        rid = rel.get('Id')
        if rid and rid.startswith('rId'):
            try:
                num = int(rid[3:])
                if num > max_id:
                    max_id = num
            except:
                pass
    return f'rId{max_id + 1}'

def _make_pic_element(left, top, width, height, rel_id, shape_name):
    """Create a picture element with a custom name that will persist in PowerPoint."""
    pic = Element(f'{{{NS_P}}}pic')
    nvPicPr = SubElement(pic, f'{{{NS_P}}}nvPicPr')
    SubElement(nvPicPr, f'{{{NS_P}}}cNvPr', id='0', name=shape_name)
    cNvPicPr = SubElement(nvPicPr, f'{{{NS_P}}}cNvPicPr')
    SubElement(cNvPicPr, f'{{{NS_A}}}picLocks', noChangeAspect='1')
    SubElement(nvPicPr, f'{{{NS_P}}}nvPr')
    blipFill = SubElement(pic, f'{{{NS_P}}}blipFill')
    SubElement(blipFill, f'{{{NS_A}}}blip', {f'{{{NS_R}}}embed': rel_id})
    stretch = SubElement(blipFill, f'{{{NS_A}}}stretch')
    SubElement(stretch, f'{{{NS_A}}}fillRect')
    spPr = SubElement(pic, f'{{{NS_P}}}spPr')
    xfrm = SubElement(spPr, f'{{{NS_A}}}xfrm')
    SubElement(xfrm, f'{{{NS_A}}}off', x=str(left), y=str(top))
    SubElement(xfrm, f'{{{NS_A}}}ext', cx=str(width), cy=str(height))
    prstGeom = SubElement(spPr, f'{{{NS_A}}}prstGeom', prst='rect')
    SubElement(prstGeom, f'{{{NS_A}}}avLst')
    return pic

def _make_textbox_element(left, top, width, height, text):
    """Create a <p:sp> text box element with the given text and geometry."""
    p = NS_P
    a = NS_A
    
    sp = Element(f'{{{p}}}sp')
    
    # nvSpPr
    nvSpPr = SubElement(sp, f'{{{p}}}nvSpPr')
    SubElement(nvSpPr, f'{{{p}}}cNvPr', id='0', name=text)
    SubElement(nvSpPr, f'{{{p}}}cNvSpPr', txBox='1')
    SubElement(nvSpPr, f'{{{p}}}nvPr')
    
    # spPr
    spPr = SubElement(sp, f'{{{p}}}spPr')
    xfrm = SubElement(spPr, f'{{{a}}}xfrm')
    SubElement(xfrm, f'{{{a}}}off', x=str(left), y=str(top))
    SubElement(xfrm, f'{{{a}}}ext', cx=str(width), cy=str(height))
    prstGeom = SubElement(spPr, f'{{{a}}}prstGeom', prst='rect')
    SubElement(prstGeom, f'{{{a}}}avLst')
    
    # txBody
    txBody = SubElement(sp, f'{{{p}}}txBody')
    bodyPr = SubElement(txBody, f'{{{a}}}bodyPr', wrap='square', rtlCol='0')
    SubElement(bodyPr, f'{{{a}}}spAutoFit')
    p_elem = SubElement(txBody, f'{{{a}}}p')
    r_elem = SubElement(p_elem, f'{{{a}}}r')
    rPr = SubElement(r_elem, f'{{{a}}}rPr', lang='en-US', sz='1800', b='0', i='0')
    SubElement(rPr, f'{{{a}}}solidFill')
    SubElement(r_elem, f'{{{a}}}t').text = text
    endParaRPr = SubElement(p_elem, f'{{{a}}}endParaRPr', lang='en-US')
    
    return sp

# ---------- Geometry resolution (for finding placeholder positions) ----------
def _get_placeholder_idx(shape):
    ph = shape.find(f'{{{NS_P}}}nvSpPr/{{{NS_P}}}nvPr/{{{NS_P}}}ph')
    return ph.get('idx') if ph is not None else None

def _get_shape_name(shape):
    cNvPr = shape.find(f'{{{NS_P}}}nvSpPr/{{{NS_P}}}cNvPr')
    return cNvPr.get('name', '') if cNvPr is not None else ''

def _resolve_geometry_from_layout(zip_ref, slide_rels_xml, placeholder_idx):
    rels_root = ET.fromstring(slide_rels_xml)
    layout_target = None
    for rel in rels_root.findall(f'{{{NS_REL}}}Relationship'):
        if rel.get('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout':
            layout_target = rel.get('Target')
            break
    if not layout_target:
        raise ValueError("Slide layout relationship not found")
    layout_filename = layout_target.split('/')[-1]
    layout_path = f'ppt/slideLayouts/{layout_filename}'
    layout_xml = zip_ref.read(layout_path)
    layout_root = ET.fromstring(layout_xml)
    for elem in layout_root.iter():
        if elem.tag == f'{{{NS_P}}}sp':
            if _get_placeholder_idx(elem) == placeholder_idx:
                xfrm = elem.find(f'{{{NS_P}}}spPr/{{{NS_A}}}xfrm')
                if xfrm is not None:
                    off = xfrm.find(f'{{{NS_A}}}off')
                    ext = xfrm.find(f'{{{NS_A}}}ext')
                    if off is not None and ext is not None:
                        return (int(off.get('x', 0)), int(off.get('y', 0)),
                                int(ext.get('cx', 0)), int(ext.get('cy', 0)))
    raise ValueError(f"Placeholder idx={placeholder_idx} not found in layout")

def _get_group_xfrm(grp_sp_element):
    xfrm = grp_sp_element.find(f'{{{NS_P}}}grpSpPr/{{{NS_A}}}xfrm')
    if xfrm is None:
        return 0, 0, None, None
    off = xfrm.find(f'{{{NS_A}}}off')
    ext = xfrm.find(f'{{{NS_A}}}ext')
    left = int(off.get('x', '0')) if off is not None else 0
    top = int(off.get('y', '0')) if off is not None else 0
    width = int(ext.get('cx', '0')) if ext is not None else None
    height = int(ext.get('cy', '0')) if ext is not None else None
    return left, top, width, height

def _resolve_shape_geometry(shape, zip_ref, slide_rels_xml, parent_offset, group_chain):
    cum_left, cum_top = parent_offset
    own_xfrm = shape.find(f'{{{NS_P}}}spPr/{{{NS_A}}}xfrm')
    if own_xfrm is not None:
        off = own_xfrm.find(f'{{{NS_A}}}off')
        ext = own_xfrm.find(f'{{{NS_A}}}ext')
        if off is not None and ext is not None:
            left = int(off.get('x', '0'))
            top = int(off.get('y', '0'))
            width = int(ext.get('cx', '0'))
            height = int(ext.get('cy', '0'))
            return (cum_left + left, cum_top + top, width, height)
    ph_idx = _get_placeholder_idx(shape)
    if ph_idx is not None:
        try:
            l, t, w, h = _resolve_geometry_from_layout(zip_ref, slide_rels_xml, ph_idx)
            return (cum_left + l, cum_top + t, w, h)
        except ValueError:
            pass
    final_w, final_h = None, None
    for _, grp_w, grp_h in reversed(group_chain):
        if final_w is None and grp_w is not None:
            final_w = grp_w
        if final_h is None and grp_h is not None:
            final_h = grp_h
        if final_w is not None and final_h is not None:
            break
    return (cum_left, cum_top, final_w or DEFAULT_SIZE_EMU, final_h or DEFAULT_SIZE_EMU)

def _find_text_shapes(spTree, zip_ref, slide_rels_xml, slide_num=None, search_text=None,
                      parent_offset=(0,0), group_chain=None):
    if group_chain is None:
        group_chain = []
    for child in list(spTree):
        tag = child.tag.split('}')[-1]
        if tag == 'grpSp':
            g_left, g_top, g_width, g_height = _get_group_xfrm(child)
            new_offset = (parent_offset[0] + g_left, parent_offset[1] + g_top)
            new_chain = group_chain + [(child, g_width, g_height)]
            yield from _find_text_shapes(child, zip_ref, slide_rels_xml,
                                         slide_num, search_text, new_offset, new_chain)
        elif tag == 'sp':
            txBody = child.find(f'{{{NS_P}}}txBody')
            if txBody is None:
                continue
            texts = []
            for t in txBody.findall(f'.//{{{NS_A}}}t'):
                if t.text:
                    texts.append(t.text)
            full_text = ''.join(texts).strip()
            if search_text is not None and search_text not in full_text:
                continue
            if not full_text:
                continue
            abs_l, abs_t, abs_w, abs_h = _resolve_shape_geometry(
                child, zip_ref, slide_rels_xml, parent_offset, group_chain)
            yield (child, spTree, abs_l, abs_t, abs_w, abs_h, full_text, None)

# ---------- Tracker ----------
class ReplacementTracker:
    def __init__(self):
        self.replacements = []
    
    def add_replacement(self, slide_num, placeholder_text, picture_name,
                        template_position, template_size):
        self.replacements.append({
            'slide': slide_num,
            'placeholder_text': placeholder_text,
            'picture_name': picture_name,
            'template_geometry': {
                'left': template_position[0], 'top': template_position[1],
                'width': template_size[0], 'height': template_size[1]
            },
            'output_geometry': None,
        })
    
    def set_output_geometry(self, picture_name, slide_num, left, top, width, height):
        for r in self.replacements:
            if r['picture_name'] == picture_name and r['slide'] == slide_num:
                r['output_geometry'] = {'left': left, 'top': top, 'width': width, 'height': height}
                return True
        return False
    
    def get_adjusted(self):
        return [r for r in self.replacements if r['output_geometry'] is not None]
    
    def print_report(self):
        print("\n" + "="*60)
        print("REPLACEMENT TRACKER")
        print("="*60)
        for i, r in enumerate(self.replacements, 1):
            status = "adjusted" if r['output_geometry'] else "unchanged"
            print(f"  {i}. Slide {r['slide']}: '{r['picture_name']}' – {status}")
            print(f"     Template: ({r['template_geometry']['left']}, {r['template_geometry']['top']}) "
                  f"{r['template_geometry']['width']}×{r['template_geometry']['height']}")
            if r['output_geometry']:
                print(f"     Output:   ({r['output_geometry']['left']}, {r['output_geometry']['top']}) "
                      f"{r['output_geometry']['width']}×{r['output_geometry']['height']}")
        print("="*60)

# ---------- Step 1: Replace placeholders ----------
def replace_with_svg(pptx_path, svg_path, search_text,
                     output_path=None, tracker=None):
    if output_path is None:
        output_path = pptx_path
    if tracker is None:
        tracker = ReplacementTracker()
    
    svg_filename = os.path.basename(svg_path)
    with open(svg_path, 'rb') as f:
        svg_data = f.read()
    
    with zipfile.ZipFile(pptx_path, 'r') as zin:
        with zin.open('[Content_Types].xml') as f:
            ct_tree = ET.parse(f)
        _ensure_svg_content_type(ct_tree.getroot())
        
        slide_files = sorted(n for n in zin.namelist()
                             if n.startswith('ppt/slides/slide') and n.endswith('.xml')
                             and '_rels' not in n)
        modified = {}
        next_img = _next_image_num(zin)
        
        for slide_file in slide_files:
            slide_num = int(slide_file[len('ppt/slides/slide'):-4])
            rels_name = f'ppt/slides/_rels/slide{slide_num}.xml.rels'
            slide_xml = zin.read(slide_file)
            slide_rels_xml = zin.read(rels_name) if rels_name in zin.namelist() else None
            slide_root = ET.fromstring(slide_xml)
            spTree = slide_root.find(f'{{{NS_P}}}cSld/{{{NS_P}}}spTree')
            if spTree is None:
                continue
            matches = list(_find_text_shapes(spTree, zin, slide_rels_xml, slide_num, search_text))
            if not matches:
                continue
            rels_root = ET.fromstring(slide_rels_xml) if slide_rels_xml else Element(f'{{{NS_REL}}}Relationships')
            for shape, parent, left, top, width, height, text, _ in matches:
                parent.remove(shape)
                img_num = next_img
                next_img += 1
                # Unique picture name: placeholder text + image number
                pic_name = f"{text}_{img_num}"
                # relationship
                rel_id = _next_rel_id(rels_root)
                rel = Element(f'{{{NS_REL}}}Relationship',
                              Id=rel_id,
                              Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                              Target=f'../media/image{img_num}.svg')
                rels_root.append(rel)
                pic = _make_pic_element(left, top, width, height, rel_id, pic_name)
                spTree.append(pic)
                tracker.add_replacement(slide_num, text, pic_name, (left, top), (width, height))
            modified[slide_file] = ET.tostring(slide_root, encoding='UTF-8', xml_declaration=True)
            modified[rels_name] = ET.tostring(rels_root, encoding='UTF-8', xml_declaration=True)
        
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename in modified:
                    zout.writestr(item, modified[item.filename])
                elif item.filename == '[Content_Types].xml':
                    zout.writestr(item, ET.tostring(ct_tree.getroot(), encoding='UTF-8', xml_declaration=True))
                else:
                    zout.writestr(item, zin.read(item.filename))
            # Write SVG for each image number used
            first_img = next_img - len(tracker.replacements)
            for img_num in range(first_img, next_img):
                zout.writestr(f'ppt/media/image{img_num}.svg', svg_data)
    
    return output_path, tracker

def replace_with_svg_document(pptx_template_path, replacement_map, output_path=None, tracker=None):
    """
    Replace multiple different placeholders with different SVGs in one pass.
    
    Args:
        pptx_template_path: Path to the template PPTX file
        replacement_map: Dictionary mapping placeholder text to SVG file path
            Example: {
                '{CompanyLogo}': 'path/to/logo.svg',
                '{Chart}': 'path/to/chart.svg',
                '{Photo}': 'path/to/photo.svg'
            }
        output_path: Where to save the output PPTX (default: template_replaced.pptx)
        tracker: Optional ReplacementTracker instance
    
    Returns:
        (output_path, tracker) tuple
    """
    if output_path is None:
        base, ext = os.path.splitext(pptx_template_path)
        output_path = f"{base}_replaced{ext}"
    if tracker is None:
        tracker = ReplacementTracker()
    
    # Load all SVGs into memory
    svg_data_map = {}
    for search_text, svg_path in replacement_map.items():
        with open(svg_path, 'rb') as f:
            svg_data_map[search_text] = f.read()
    
    with zipfile.ZipFile(pptx_template_path, 'r') as zin:
        with zin.open('[Content_Types].xml') as f:
            ct_tree = ET.parse(f)
        _ensure_svg_content_type(ct_tree.getroot())
        
        slide_files = sorted(n for n in zin.namelist()
                             if n.startswith('ppt/slides/slide') and n.endswith('.xml')
                             and '_rels' not in n)
        modified = {}
        next_img = _next_image_num(zin)
        
        for slide_file in slide_files:
            slide_num = int(slide_file[len('ppt/slides/slide'):-4])
            rels_name = f'ppt/slides/_rels/slide{slide_num}.xml.rels'
            slide_xml = zin.read(slide_file)
            slide_rels_xml = zin.read(rels_name) if rels_name in zin.namelist() else None
            slide_root = ET.fromstring(slide_xml)
            spTree = slide_root.find(f'{{{NS_P}}}cSld/{{{NS_P}}}spTree')
            if spTree is None:
                continue
            
            rels_root = ET.fromstring(slide_rels_xml) if slide_rels_xml else Element(f'{{{NS_REL}}}Relationships')
            
            # Process each placeholder type
            for search_text, svg_data in svg_data_map.items():
                svg_filename = os.path.basename(replacement_map[search_text])
                matches = list(_find_text_shapes(spTree, zin, slide_rels_xml, slide_num, search_text))
                
                for shape, parent, left, top, width, height, text, _ in matches:
                    parent.remove(shape)
                    img_num = next_img
                    next_img += 1
                    pic_name = f"{text}_{img_num}"
                    
                    # Relationship
                    rel_id = _next_rel_id(rels_root)
                    rel = Element(f'{{{NS_REL}}}Relationship',
                                  Id=rel_id,
                                  Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                                  Target=f'../media/image{img_num}.svg')
                    rels_root.append(rel)
                    
                    pic = _make_pic_element(left, top, width, height, rel_id, pic_name)
                    spTree.append(pic)
                    
                    tracker.add_replacement(slide_num, text, pic_name, (left, top), (width, height))
            
            modified[slide_file] = ET.tostring(slide_root, encoding='UTF-8', xml_declaration=True)
            modified[rels_name] = ET.tostring(rels_root, encoding='UTF-8', xml_declaration=True)
        
        # Write output
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename in modified:
                    zout.writestr(item, modified[item.filename])
                elif item.filename == '[Content_Types].xml':
                    zout.writestr(item, ET.tostring(ct_tree.getroot(), encoding='UTF-8', xml_declaration=True))
                else:
                    zout.writestr(item, zin.read(item.filename))
            
            # Write SVGs - need to track which image number got which SVG
            img_assignments = {}
            for r in tracker.replacements:
                img_num = int(r['picture_name'].split('_')[-1])
                search_text = r['placeholder_text']
                img_assignments[img_num] = svg_data_map[search_text]
            
            for img_num, svg_data in img_assignments.items():
                zout.writestr(f'ppt/media/image{img_num}.svg', svg_data)
    
    return output_path, tracker

# ---------- Step 2: Apply adjusted geometry ----------
def apply_geometry_to_template(template_path, adjusted_pptx_path, output_path=None, tracker=None):
    """Find pictures in adjusted file by their unique name, read new geometry,
       and create TEXT BOXES with the placeholder code at the adjusted positions in the fresh template."""
    if output_path is None:
        base, ext = os.path.splitext(template_path)
        output_path = f"{base}_adjusted{ext}"
    if tracker is None:
        raise ValueError("Tracker is required.")
    
    # 1) Extract current geometry of all pictures in the adjusted file, keyed by picture name
    adjusted_geoms = {}
    with zipfile.ZipFile(adjusted_pptx_path, 'r') as zadj:
        slide_files = [n for n in zadj.namelist()
                       if n.startswith('ppt/slides/slide') and n.endswith('.xml')
                       and '_rels' not in n]
        for slide_file in slide_files:
            slide_num = int(slide_file[len('ppt/slides/slide'):-4])
            slide_xml = zadj.read(slide_file)
            slide_root = ET.fromstring(slide_xml)
            for pic in slide_root.iter(f'{{{NS_P}}}pic'):
                cNvPr = pic.find(f'{{{NS_P}}}nvPicPr/{{{NS_P}}}cNvPr')
                if cNvPr is None:
                    continue
                pic_name = cNvPr.get('name')
                if not pic_name:
                    continue
                xfrm = pic.find(f'{{{NS_P}}}spPr/{{{NS_A}}}xfrm')
                if xfrm is None:
                    continue
                off = xfrm.find(f'{{{NS_A}}}off')
                ext = xfrm.find(f'{{{NS_A}}}ext')
                if off is None or ext is None:
                    continue
                left = int(off.get('x', '0'))
                top = int(off.get('y', '0'))
                width = int(ext.get('cx', '0'))
                height = int(ext.get('cy', '0'))
                adjusted_geoms[(slide_num, pic_name)] = (left, top, width, height)
                print(f"  Found picture '{pic_name}' on slide {slide_num} at ({left},{top}) {width}×{height}")
    
    # 2) Update tracker with found geometries
    matches = 0
    for r in tracker.replacements:
        key = (r['slide'], r['picture_name'])
        if key in adjusted_geoms:
            l, t, w, h = adjusted_geoms[key]
            r['output_geometry'] = {'left': l, 'top': t, 'width': w, 'height': h}
            matches += 1
            print(f"  ✓ Matched '{r['picture_name']}' on slide {r['slide']}")
        else:
            print(f"  ✗ Picture '{r['picture_name']}' not found in adjusted file")
    
    if matches == 0:
        print("WARNING: No adjusted pictures matched. Output will be a copy of the template.")
        import shutil
        shutil.copy2(template_path, output_path)
        return output_path, tracker
    
    # 3) Replace placeholders in fresh template with TEXT BOXES at adjusted geometry
    with zipfile.ZipFile(template_path, 'r') as zin:
        with zin.open('[Content_Types].xml') as f:
            ct_tree = ET.parse(f)
        # No need to add SVG content type since we're not adding images
        
        modified_slides = {}
        modified_rels = {}
        
        for r in tracker.replacements:
            if r['output_geometry'] is None:
                continue
            slide_num = r['slide']
            slide_file = f'ppt/slides/slide{slide_num}.xml'
            rels_file = f'ppt/slides/_rels/slide{slide_num}.xml.rels'
            
            # Load slide if not already loaded
            if slide_num not in modified_slides:
                slide_xml = zin.read(slide_file)
                slide_rels_xml = zin.read(rels_file) if rels_file in zin.namelist() else None
                slide_root = ET.fromstring(slide_xml)
                spTree = slide_root.find(f'{{{NS_P}}}cSld/{{{NS_P}}}spTree')
                rels_root = ET.fromstring(slide_rels_xml) if slide_rels_xml else Element(f'{{{NS_REL}}}Relationships')
                modified_slides[slide_num] = (slide_root, spTree)
                modified_rels[slide_num] = rels_root
            
            slide_root, spTree = modified_slides[slide_num]
            
            # Find and remove the ORIGINAL placeholder text box
            found = False
            for shape, parent, *_, text, _ in _find_text_shapes(spTree, zin, slide_rels_xml, slide_num, r['placeholder_text']):
                parent.remove(shape)
                found = True
                break
            
            if not found:
                print(f"  ⚠ Placeholder '{r['placeholder_text']}' not found on slide {slide_num}")
                continue
            
            # Create a new text box at the adjusted position
            out = r['output_geometry']
            new_textbox = _make_textbox_element(
                out['left'], out['top'], out['width'], out['height'],
                r['placeholder_text']
            )
            spTree.append(new_textbox)
            print(f"  ✓ Created textbox '{r['placeholder_text']}' at ({out['left']},{out['top']}) {out['width']}×{out['height']}")
        
        # Write final output
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == '[Content_Types].xml':
                    zout.writestr(item, ET.tostring(ct_tree.getroot(), encoding='UTF-8', xml_declaration=True))
                elif item.filename in [f'ppt/slides/slide{num}.xml' for num in modified_slides]:
                    slide_num = int(item.filename[len('ppt/slides/slide'):-4])
                    slide_root, _ = modified_slides[slide_num]
                    zout.writestr(item, ET.tostring(slide_root, encoding='UTF-8', xml_declaration=True))
                elif item.filename in [f'ppt/slides/_rels/slide{num}.xml.rels' for num in modified_rels]:
                    slide_num = int(item.filename[len('ppt/slides/_rels/slide'):-len('.xml.rels')])
                    rels_root = modified_rels[slide_num]
                    zout.writestr(item, ET.tostring(rels_root, encoding='UTF-8', xml_declaration=True))
                else:
                    zout.writestr(item, zin.read(item.filename))
    
    return output_path, tracker

# ---------- Example usage ----------
if __name__ == '__main__':
    tracker = ReplacementTracker()
    
    print("STEP 1: Replace placeholders with SVG")
    out1, tracker = replace_with_svg(
        r'C:/Development/Powerpoint/Test.pptx',
        r'C:/Development/Powerpoint/test.svg',
        '{CompanyLogo}',
        output_path=r'C:/Development/Powerpoint/Test3.pptx',
        tracker=tracker
    )
    tracker.print_report()
    
    input("\nNow adjust images in Test3.pptx (move/resize), save, then press Enter...")
    
    print("\nSTEP 2: Apply adjusted geometry to fresh template")
    out2, tracker = apply_geometry_to_template(
        template_path=r'C:/Development/Powerpoint/Test.pptx',
        adjusted_pptx_path=r'C:/Development/Powerpoint/Test3.pptx',
        output_path=r'C:/Development/Powerpoint/Test4_final.pptx',
        tracker=tracker
    )
    tracker.print_report()
    print(f"\nDone! Final output: {out2}")