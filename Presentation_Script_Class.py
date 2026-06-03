import zipfile
import os
import io
import json
import shutil
import tempfile
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element, SubElement
from collections import defaultdict
from pathlib import Path
import time
import re
import copy


class PowerPointTemplate:
    """A class for managing PowerPoint template placeholder replacement and geometry adjustment."""

    # XML namespaces
    NSMAP = {
        'a':  'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r':  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'p':  'http://schemas.openxmlformats.org/presentationml/2006/main',
        'rel':'http://schemas.openxmlformats.org/package/2006/relationships',
    }

    NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
    NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'

    DEFAULT_SIZE_EMU = 914400

    IMAGE_FORMATS = {
        '.svg':  'image/svg+xml',
        '.png':  'image/png',
        '.jpg':  'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.gif':  'image/gif',
        '.bmp':  'image/bmp',
        '.tiff': 'image/tiff',
        '.tif':  'image/tiff',
        '.webp': 'image/webp',
    }

    def __init__(self, template_path):
        self.template_path = template_path
        self.tracker = ReplacementTracker()
        self._tmp_dir = tempfile.TemporaryDirectory()
        self.tmp_path = os.path.join(self._tmp_dir.name, f"{hash(time.time())}_working.pptx")
        shutil.copy2(self.template_path, self.tmp_path)
        self.tracker.template_path = template_path
        for prefix, uri in self.NSMAP.items():
            ET.register_namespace(prefix, uri)

    def __del__(self):
        self.cleanup()

    # ---------- Helpers (unchanged) ----------
    @classmethod
    def _get_image_format(cls, file_path_or_bytes, filename=None):
        if isinstance(file_path_or_bytes, (str, Path)) and os.path.exists(str(file_path_or_bytes)):
            ext = os.path.splitext(str(file_path_or_bytes))[1].lower()
            if ext in cls.IMAGE_FORMATS:
                return ext, cls.IMAGE_FORMATS[ext]
        if filename:
            ext = os.path.splitext(filename)[1].lower()
            if ext in cls.IMAGE_FORMATS:
                return ext, cls.IMAGE_FORMATS[ext]
        if isinstance(file_path_or_bytes, bytes):
            if file_path_or_bytes.startswith(b'\x89PNG'):
                return '.png', 'image/png'
            elif file_path_or_bytes.startswith(b'\xff\xd8\xff'):
                return '.jpg', 'image/jpeg'
            elif file_path_or_bytes.startswith(b'GIF8'):
                return '.gif', 'image/gif'
            elif file_path_or_bytes.startswith(b'<?xml') or file_path_or_bytes.startswith(b'<svg'):
                return '.svg', 'image/svg+xml'
            elif file_path_or_bytes.startswith(b'RIFF') and b'WEBP' in file_path_or_bytes[:12]:
                return '.webp', 'image/webp'
        return '.png', 'image/png'

    @classmethod
    def _ensure_content_type(cls, root, extension, mime_type):
        for default in root.findall(f'{{{cls.NS_CT}}}Default'):
            if default.get('Extension') == extension.lstrip('.'):
                if default.get('ContentType') != mime_type:
                    default.set('ContentType', mime_type)
                return
        default = Element(f'{{{cls.NS_CT}}}Default',
                          Extension=extension.lstrip('.'),
                          ContentType=mime_type)
        root.append(default)

    @staticmethod
    def _fig_to_image(fig, format='svg', **kwargs):
        buf = io.BytesIO()
        save_kwargs = {
            'format': format,
            'dpi': kwargs.get('dpi', fig.dpi),
            'bbox_inches': kwargs.get('bbox_inches', 'tight'),
            'pad_inches': kwargs.get('pad_inches', 0.1),
        }
        if format == 'svg':
            save_kwargs['transparent'] = kwargs.get('transparent', True)
        elif format == 'png':
            save_kwargs['transparent'] = kwargs.get('transparent', False)
            save_kwargs.pop('pad_inches', None)
        fig.savefig(buf, **save_kwargs)
        buf.seek(0)
        return buf.read(), f'.{format}'

    @staticmethod
    def _next_image_num(zip_ref, prefix='ppt/media/image'):
        existing = [n for n in zip_ref.namelist() if n.startswith(prefix)]
        max_num = 0
        for name in existing:
            try:
                base = os.path.splitext(os.path.basename(name))[0]
                if base.startswith('image'):
                    num = int(base[5:])
                    if num > max_num:
                        max_num = num
            except:
                pass
        return max_num + 1

    @staticmethod
    def _next_rel_id(rels_element):
        max_id = 0
        for rel in rels_element.findall(f'{{{PowerPointTemplate.NS_REL}}}Relationship'):
            rid = rel.get('Id')
            if rid and rid.startswith('rId'):
                try:
                    num = int(rid[3:])
                    if num > max_id:
                        max_id = num
                except:
                    pass
        return f'rId{max_id + 1}'
    
    @staticmethod
    def find_textbox_xml(pptx_path, search_text):
        """Print the full XML of the first shape containing `search_text`."""
        NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
        NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    
        with zipfile.ZipFile(pptx_path, 'r') as z:
            slide_files = [n for n in z.namelist()
                           if n.startswith('ppt/slides/slide') and n.endswith('.xml') and '_rels' not in n]
            for slide_file in sorted(slide_files):
                slide_xml = z.read(slide_file)
                root = ET.fromstring(slide_xml)
                spTree = root.find(f'.//{{{NS_P}}}spTree')
                if spTree is None:
                    continue
                for shape in spTree.iter():
                    tag = shape.tag.split('}')[-1]
                    if tag not in ('sp', 'graphicFrame'):
                        continue
                    txBody = shape.find(f'{{{NS_P}}}txBody')
                    if txBody is None:
                        continue
                    # collect all text
                    texts = [t.text or '' for t in txBody.findall(f'.//{{{NS_A}}}t')]
                    full_text = ''.join(texts)
                    if search_text in full_text:
                        # found the shape — pretty‑print its XML
                        print(f"Found '{search_text}' in {slide_file}")
                        print(ET.tostring(shape, encoding='unicode'))
                        return  # stop after first match
            print(f"Text '{search_text}' not found in {pptx_path}")

    @staticmethod
    def _make_pic_element(left, top, width, height, rel_id, shape_name):
        pic = Element(f'{{{PowerPointTemplate.NS_P}}}pic')
        nvPicPr = SubElement(pic, f'{{{PowerPointTemplate.NS_P}}}nvPicPr')
        SubElement(nvPicPr, f'{{{PowerPointTemplate.NS_P}}}cNvPr', id='0', name=shape_name)
        cNvPicPr = SubElement(nvPicPr, f'{{{PowerPointTemplate.NS_P}}}cNvPicPr')
        SubElement(cNvPicPr, f'{{{PowerPointTemplate.NS_A}}}picLocks', noChangeAspect='1')
        SubElement(nvPicPr, f'{{{PowerPointTemplate.NS_P}}}nvPr')
        blipFill = SubElement(pic, f'{{{PowerPointTemplate.NS_P}}}blipFill')
        SubElement(blipFill, f'{{{PowerPointTemplate.NS_A}}}blip',
                   {f'{{{PowerPointTemplate.NS_R}}}embed': rel_id})
        stretch = SubElement(blipFill, f'{{{PowerPointTemplate.NS_A}}}stretch')
        SubElement(stretch, f'{{{PowerPointTemplate.NS_A}}}fillRect')
        spPr = SubElement(pic, f'{{{PowerPointTemplate.NS_P}}}spPr')
        xfrm = SubElement(spPr, f'{{{PowerPointTemplate.NS_A}}}xfrm')
        SubElement(xfrm, f'{{{PowerPointTemplate.NS_A}}}off', x=str(left), y=str(top))
        SubElement(xfrm, f'{{{PowerPointTemplate.NS_A}}}ext', cx=str(width), cy=str(height))
        prstGeom = SubElement(spPr, f'{{{PowerPointTemplate.NS_A}}}prstGeom', prst='rect')
        SubElement(prstGeom, f'{{{PowerPointTemplate.NS_A}}}avLst')
        return pic

    @staticmethod
    def _make_textbox_element(left, top, width, height, text):
        p = PowerPointTemplate.NS_P
        a = PowerPointTemplate.NS_A
        sp = Element(f'{{{p}}}sp')
        nvSpPr = SubElement(sp, f'{{{p}}}nvSpPr')
        SubElement(nvSpPr, f'{{{p}}}cNvPr', id='0', name=text)
        SubElement(nvSpPr, f'{{{p}}}cNvSpPr', txBox='1')
        SubElement(nvSpPr, f'{{{p}}}nvPr')
        spPr = SubElement(sp, f'{{{p}}}spPr')
        xfrm = SubElement(spPr, f'{{{a}}}xfrm')
        SubElement(xfrm, f'{{{a}}}off', x=str(left), y=str(top))
        SubElement(xfrm, f'{{{a}}}ext', cx=str(width), cy=str(height))
        prstGeom = SubElement(spPr, f'{{{a}}}prstGeom', prst='rect')
        SubElement(prstGeom, f'{{{a}}}avLst')
        txBody = SubElement(sp, f'{{{p}}}txBody')
        bodyPr = SubElement(txBody, f'{{{a}}}bodyPr', wrap='square', rtlCol='0')
        SubElement(bodyPr, f'{{{a}}}spAutoFit')
        p_elem = SubElement(txBody, f'{{{a}}}p')
        r_elem = SubElement(p_elem, f'{{{a}}}r')
        SubElement(r_elem, f'{{{a}}}rPr', lang='en-US', sz='1800', b='0', i='0')
        SubElement(r_elem, f'{{{a}}}t').text = text
        SubElement(p_elem, f'{{{a}}}endParaRPr', lang='en-US')
        return sp

    # ---------- Geometry resolution (unchanged) ----------
    @staticmethod
    def _get_placeholder_idx(shape):
        ph = shape.find(f'{{{PowerPointTemplate.NS_P}}}nvSpPr/'
                        f'{{{PowerPointTemplate.NS_P}}}nvPr/'
                        f'{{{PowerPointTemplate.NS_P}}}ph')
        return ph.get('idx') if ph is not None else None

    @staticmethod
    def _get_group_xfrm(grp_sp_element):
        NS_P = PowerPointTemplate.NS_P
        NS_A = PowerPointTemplate.NS_A
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

    @classmethod
    def _resolve_geometry_from_layout(cls, zip_ref, slide_rels_xml, placeholder_idx):
        NS_REL = cls.NS_REL
        NS_P = cls.NS_P
        NS_A = cls.NS_A
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
                if cls._get_placeholder_idx(elem) == placeholder_idx:
                    xfrm = elem.find(f'{{{NS_P}}}spPr/{{{NS_A}}}xfrm')
                    if xfrm is not None:
                        off = xfrm.find(f'{{{NS_A}}}off')
                        ext = xfrm.find(f'{{{NS_A}}}ext')
                        if off is not None and ext is not None:
                            return (int(off.get('x', 0)), int(off.get('y', 0)),
                                    int(ext.get('cx', 0)), int(ext.get('cy', 0)))
        raise ValueError(f"Placeholder idx={placeholder_idx} not found in layout")

    @classmethod
    def _resolve_shape_geometry(cls, shape, zip_ref, slide_rels_xml,
                                parent_offset, group_chain):
        cum_left, cum_top = parent_offset
        own_xfrm = shape.find(f'{{{cls.NS_P}}}spPr/{{{cls.NS_A}}}xfrm')
        if own_xfrm is not None:
            off = own_xfrm.find(f'{{{cls.NS_A}}}off')
            ext = own_xfrm.find(f'{{{cls.NS_A}}}ext')
            if off is not None and ext is not None:
                left = int(off.get('x', '0'))
                top = int(off.get('y', '0'))
                width = int(ext.get('cx', '0'))
                height = int(ext.get('cy', '0'))
                return (cum_left + left, cum_top + top, width, height)
        ph_idx = cls._get_placeholder_idx(shape)
        if ph_idx is not None:
            try:
                l, t, w, h = cls._resolve_geometry_from_layout(zip_ref, slide_rels_xml, ph_idx)
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
        return (cum_left, cum_top,
                final_w or cls.DEFAULT_SIZE_EMU,
                final_h or cls.DEFAULT_SIZE_EMU)

    @classmethod
    def _find_text_shapes(cls, spTree, zip_ref, slide_rels_xml,
                          slide_num=None, search_text=None,
                          parent_offset=(0, 0), group_chain=None):
        if group_chain is None:
            group_chain = []
        for child in list(spTree):
            tag = child.tag.split('}')[-1]
            if tag == 'grpSp':
                g_left, g_top, g_width, g_height = cls._get_group_xfrm(child)
                new_offset = (parent_offset[0] + g_left, parent_offset[1] + g_top)
                new_chain = group_chain + [(child, g_width, g_height)]
                yield from cls._find_text_shapes(child, zip_ref, slide_rels_xml,
                                                 slide_num, search_text,
                                                 new_offset, new_chain)
            elif tag == 'sp':
                txBody = child.find(f'{{{cls.NS_P}}}txBody')
                if txBody is None:
                    continue
                texts = []
                for t in txBody.findall(f'.//{{{cls.NS_A}}}t'):
                    if t.text:
                        texts.append(t.text)
                full_text = ''.join(texts).strip()
                if search_text is not None and search_text not in full_text:
                    continue
                if not full_text:
                    continue
                abs_l, abs_t, abs_w, abs_h = cls._resolve_shape_geometry(
                    child, zip_ref, slide_rels_xml, parent_offset, group_chain)
                yield (child, spTree, abs_l, abs_t, abs_w, abs_h, full_text, None)

    # ---------- Image replacement (unchanged) ----------
    def _replace_with_images(self, image_map):
        with zipfile.ZipFile(self.tmp_path, 'r') as zin:
            with zin.open('[Content_Types].xml') as f:
                ct_tree = ET.parse(f)
            seen_extensions = set()
            for search_text, (data, ext, mime) in image_map.items():
                if ext not in seen_extensions:
                    self._ensure_content_type(ct_tree.getroot(), ext, mime)
                    seen_extensions.add(ext)
            slide_files = sorted(n for n in zin.namelist()
                                 if n.startswith('ppt/slides/slide')
                                 and n.endswith('.xml')
                                 and '_rels' not in n)
            modified = {}
            next_img = self._next_image_num(zin)
            for slide_file in slide_files:
                slide_num = int(slide_file[len('ppt/slides/slide'):-4])
                rels_name = f'ppt/slides/_rels/slide{slide_num}.xml.rels'
                slide_xml = zin.read(slide_file)
                slide_rels_xml = zin.read(rels_name) if rels_name in zin.namelist() else None
                slide_root = ET.fromstring(slide_xml)
                spTree = slide_root.find(f'{{{self.NS_P}}}cSld/{{{self.NS_P}}}spTree')
                if spTree is None:
                    continue
                rels_root = ET.fromstring(slide_rels_xml) if slide_rels_xml else Element(f'{{{self.NS_REL}}}Relationships')
                for search_text, (img_data, ext, mime) in image_map.items():
                    matches = list(self._find_text_shapes(spTree, zin, slide_rels_xml,
                                                          slide_num, search_text))
                    for shape, parent, left, top, width, height, text, _ in matches:
                        parent.remove(shape)
                        img_num = next_img
                        next_img += 1
                        pic_name = f"{text}_{img_num}"
                        rel_id = self._next_rel_id(rels_root)
                        rel = Element(f'{{{self.NS_REL}}}Relationship',
                                      Id=rel_id,
                                      Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                                      Target=f'../media/image{img_num}{ext}')
                        rels_root.append(rel)
                        pic = self._make_pic_element(left, top, width, height, rel_id, pic_name)
                        spTree.append(pic)
                        self.tracker.add_replacement(slide_num, text, pic_name,
                                                     (left, top), (width, height),
                                                     replacement_type='image')
                modified[slide_file] = ET.tostring(slide_root, encoding='UTF-8', xml_declaration=True)
                modified[rels_name] = ET.tostring(rels_root, encoding='UTF-8', xml_declaration=True)
            new_zip_path = self.tmp_path + ".new"
            with zipfile.ZipFile(new_zip_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename in modified:
                        zout.writestr(item, modified[item.filename])
                    elif item.filename == '[Content_Types].xml':
                        zout.writestr(item, ET.tostring(ct_tree.getroot(), encoding='UTF-8', xml_declaration=True))
                    else:
                        zout.writestr(item, zin.read(item.filename))
                img_assignments = {}
                for r in self.tracker.replacements:
                    if r['replacement_type'] != 'image':
                        continue
                    img_num = int(r['picture_name'].split('_')[-1])
                    search_text = r['placeholder_text']
                    img_data, ext, _ = image_map[search_text]
                    img_assignments[img_num] = (img_data, ext)
                for img_num, (img_data, ext) in img_assignments.items():
                    zout.writestr(f'ppt/media/image{img_num}{ext}', img_data)
        shutil.move(new_zip_path, self.tmp_path)

    def replace_image(self, placeholder_map: dict, **kwargs):
        image_map = {}
        for search_text, value in placeholder_map.items():
            if isinstance(value, tuple) and len(value) == 2:
                data, filename = value
                ext, mime = self._get_image_format(data, filename)
                image_map[search_text] = (data, ext, mime)
            elif hasattr(value, 'savefig'):
                fmt = kwargs.get('format', 'svg')
                data, ext = self._fig_to_image(value, format=fmt, **kwargs)
                mime = self.IMAGE_FORMATS.get(ext, 'image/png')
                image_map[search_text] = (data, ext, mime)
            elif isinstance(value, bytes):
                ext, mime = self._get_image_format(value)
                image_map[search_text] = (value, ext, mime)
            elif isinstance(value, (str, Path)):
                path = Path(value)
                ext = path.suffix.lower()
                mime = self.IMAGE_FORMATS.get(ext, 'image/png')
                with open(path, 'rb') as f:
                    data = f.read()
                image_map[search_text] = (data, ext, mime)
            else:
                raise TypeError(f"Unsupported type for '{search_text}': {type(value)}")
        return self._replace_with_images(image_map)

    # ---------- Text replacement (NEW – formatting safe) ----------
    def replace_text(self, placeholder_map: dict):
        import re   # safe inside the method
    
        replacer = {f'{{{k}}}': str(v) for k, v in placeholder_map.items()}
    
        with zipfile.ZipFile(self.tmp_path, 'r') as zin:
            slide_files = sorted(n for n in zin.namelist()
                                 if n.startswith('ppt/slides/slide')
                                 and n.endswith('.xml')
                                 and '_rels' not in n)
            modified = {}
    
            for slide_file in slide_files:
                slide_num = int(re.search(r'slide(\d+)', slide_file).group(1))
                slide_xml = zin.read(slide_file)
                slide_root = ET.fromstring(slide_xml)
                spTree = slide_root.find(f'{{{self.NS_P}}}cSld/{{{self.NS_P}}}spTree')
                if spTree is None:
                    continue
    
                slide_changed = False
                for sp in spTree.findall(f'{{{self.NS_P}}}sp') + spTree.findall(f'{{{self.NS_P}}}graphicFrame'):
                    txBody = sp.find(f'{{{self.NS_P}}}txBody')
                    if txBody is None:
                        continue
    
                    for p_elem in txBody.findall(f'{{{self.NS_A}}}p'):
                        runs = p_elem.findall(f'{{{self.NS_A}}}r')
                        if not runs:
                            continue
    
                        # Build full paragraph text
                        texts = []
                        for r in runs:
                            t = r.find(f'{{{self.NS_A}}}t')
                            texts.append(t.text if t is not None and t.text else '')
                        full_text = ''.join(texts)
    
                        # Check for placeholders
                        matched = [k for k in replacer if k in full_text]
                        if not matched:
                            continue
    
                        old_key = matched[0]   # handle the first one found
                        new_text = replacer[old_key]
                        start = full_text.find(old_key)
                        end = start + len(old_key)
    
                        # Identify runs overlapping the placeholder
                        involved = []
                        pos = 0
                        for i, (r, txt) in enumerate(zip(runs, texts)):
                            r_start = pos
                            r_end = pos + len(txt)
                            if r_end > start and r_start < end:
                                involved.append((i, r, txt, r_start, r_end))
                            pos = r_end
    
                        if not involved:
                            continue
    
                        # Take formatting from the first involved run
                        first_run = involved[0][1]
                        first_rPr = first_run.find(f'{{{self.NS_A}}}rPr')
                        # Remove all involved runs from the paragraph
                        for _, r, _, _, _ in involved:
                            p_elem.remove(r)
    
                        # Build the replacement text
                        prefix = texts[involved[0][0]][:max(0, start - involved[0][3])]
                        last_txt = texts[involved[-1][0]]
                        suffix = last_txt[end - involved[-1][3]:] if end < involved[-1][4] else ''
                        new_run_text = prefix + new_text + suffix
    
                        # Create a new run with the same formatting
                        new_run = Element(f'{{{self.NS_A}}}r')
                        if first_rPr is not None:
                            new_run.append(copy.deepcopy(first_rPr))
                        new_t = SubElement(new_run, f'{{{self.NS_A}}}t')
                        new_t.text = new_run_text
    
                        # Insert the new run at the correct position
                        insert_index = involved[0][0]
                        p_elem.insert(insert_index, new_run)
    
                        # Record the replacement for tracking
                        cNvPr = sp.find(f'{{{self.NS_P}}}nvSpPr/{{{self.NS_P}}}cNvPr')
                        if cNvPr is not None:
                            shape_id = cNvPr.get('id')
                            tracking_name = f"txt_{slide_num}_{shape_id}"
                            cNvPr.set('name', tracking_name)
    
                            xfrm = sp.find(f'{{{self.NS_P}}}spPr/{{{self.NS_A}}}xfrm')
                            if xfrm is not None:
                                off = xfrm.find(f'{{{self.NS_A}}}off')
                                ext = xfrm.find(f'{{{self.NS_A}}}ext')
                                left = int(off.get('x', '0'))
                                top = int(off.get('y', '0'))
                                width = int(ext.get('cx', '0'))
                                height = int(ext.get('cy', '0'))
                                self.tracker.add_replacement(
                                    slide_num=slide_num,
                                    placeholder_text=old_key,
                                    picture_name=tracking_name,
                                    template_position=(left, top),
                                    template_size=(width, height),
                                    replacement_type='text'
                                )
                        slide_changed = True
    
                if slide_changed:
                    modified[slide_file] = ET.tostring(slide_root, encoding='UTF-8')
    
            if modified:
                new_zip_path = self.tmp_path + ".new"
                with zipfile.ZipFile(new_zip_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        if item.filename in modified:
                            zout.writestr(item, modified[item.filename])
                        else:
                            zout.writestr(item, zin.read(item.filename))
                shutil.move(new_zip_path, self.tmp_path)

    # ---------- Geometry extraction for tracker.update ----------
    @staticmethod
    def _extract_all_geoms(pptx_path):
        geoms = {}
        with zipfile.ZipFile(pptx_path, 'r') as z:
            slide_files = [n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml', n)]
            for sf in slide_files:
                slide_num = int(re.search(r'slide(\d+)', sf).group(1))
                root = ET.fromstring(z.read(sf))
                spTree = root.find(f'.//{{{PowerPointTemplate.NS_P}}}spTree')
                if spTree is None:
                    continue
                for elem in spTree.iter():
                    tag = elem.tag.split('}')[-1]
                    if tag in ('sp', 'pic', 'graphicFrame'):
                        cNvPr = elem.find(f'.//{{{PowerPointTemplate.NS_P}}}cNvPr')
                        name = cNvPr.get('name') if cNvPr is not None else None
                        if not name:
                            continue
                        xfrm = elem.find(f'.//{{{PowerPointTemplate.NS_P}}}spPr/{{{PowerPointTemplate.NS_A}}}xfrm')
                        if xfrm is None:
                            continue
                        off = xfrm.find(f'{{{PowerPointTemplate.NS_A}}}off')
                        ext = xfrm.find(f'{{{PowerPointTemplate.NS_A}}}ext')
                        if off is None or ext is None:
                            continue
                        left = int(off.get('x', '0'))
                        top = int(off.get('y', '0'))
                        width = int(ext.get('cx', '0'))
                        height = int(ext.get('cy', '0'))
                        geoms[(slide_num, name)] = (left, top, width, height)
        return geoms

    # ---------- Update template with position AND formatting ----------
    def update_template(self, output_template_path=None):
        if output_template_path is None:
            base, ext = os.path.splitext(self.template_path)
            output_template_path = f"{base}_adjusted{ext}"

        if not any(r.get('output_geometry') for r in self.tracker.replacements):
            print("WARNING: No output geometries found. Call tracker.update() first.")
            shutil.copy2(self.template_path, output_template_path)
            return output_template_path

        # Pre‑load output slides
        output_slides = {}
        with zipfile.ZipFile(self.tracker.output_path, 'r') as zout:
            for r in self.tracker.replacements:
                slide_file = f'ppt/slides/slide{r["slide"]}.xml'
                if slide_file not in output_slides:
                    output_slides[slide_file] = ET.fromstring(zout.read(slide_file))

        with zipfile.ZipFile(self.template_path, 'r') as zin:
            modified_slides = {}
            for r in self.tracker.replacements:
                out_geom = r.get('output_geometry')
                if out_geom is None:
                    continue
                slide_num = r['slide']
                slide_file = f'ppt/slides/slide{slide_num}.xml'

                if slide_num not in modified_slides:
                    slide_xml = zin.read(slide_file)
                    slide_root = ET.fromstring(slide_xml)
                    modified_slides[slide_num] = slide_root

                spTree = modified_slides[slide_num].find(f'{{{self.NS_P}}}cSld/{{{self.NS_P}}}spTree')
                if spTree is None:
                    continue

                placeholder = r['placeholder_text']
                found = False
                for shape in spTree.iter():
                    tag = shape.tag.split('}')[-1]
                    # print(tag)
                    if tag == 'sp': #in ('sp', 'pic', 'graphicFrame'):
                        texts = shape.text
                        if texts is None:
                            txBody = shape.find(f'{{{self.NS_P}}}txBody')
                            if txBody is None:
                                continue
                            texts = [t.text or '' for t in txBody.findall(f'.//{{{self.NS_A}}}t')]
                            full_text = ''.join(texts)
                        if placeholder in full_text:
                            # 1. Move / resize
                            xfrm = shape.find(f'{{{self.NS_P}}}spPr/{{{self.NS_A}}}xfrm')
                            if xfrm is not None:
                                off = xfrm.find(f'{{{self.NS_A}}}off')
                                ext = xfrm.find(f'{{{self.NS_A}}}ext')
                                off.set('x', str(out_geom['left']))
                                off.set('y', str(out_geom['top']))
                                ext.set('cx', str(out_geom['width']))
                                ext.set('cy', str(out_geom['height']))

                            # 2. Inherit formatting from output shape
                            if slide_file in output_slides:
                                out_root = output_slides[slide_file]
                                out_spTree = out_root.find(f'{{{self.NS_P}}}cSld/{{{self.NS_P}}}spTree')
                                if out_spTree is not None:
                                    for out_shape in out_spTree.iter():
                                        out_cNvPr = out_shape.find(f'{{{self.NS_P}}}cNvPr')
                                        if out_cNvPr is not None and out_cNvPr.get('name') == r['picture_name']:
                                            out_txBody = out_shape.find(f'{{{self.NS_P}}}txBody')
                                            if txBody is not None and out_txBody is not None:
                                                # Copy bodyPr attributes
                                                tmpl_bodyPr = txBody.find(f'{{{self.NS_A}}}bodyPr')
                                                out_bodyPr = out_txBody.find(f'{{{self.NS_A}}}bodyPr')
                                                if tmpl_bodyPr is not None and out_bodyPr is not None:
                                                    for attr, val in out_bodyPr.attrib.items():
                                                        tmpl_bodyPr.set(attr, val)

                                                # Copy paragraph & run formatting
                                                tmpl_paras = txBody.findall(f'{{{self.NS_A}}}p')
                                                out_paras = out_txBody.findall(f'{{{self.NS_A}}}p')
                                                for tp, op in zip(tmpl_paras, out_paras):
                                                    tp_pPr = tp.find(f'{{{self.NS_A}}}pPr')
                                                    op_pPr = op.find(f'{{{self.NS_A}}}pPr')
                                                    if tp_pPr is not None and op_pPr is not None:
                                                        tp.remove(tp_pPr)
                                                        tp.insert(0, copy.deepcopy(op_pPr))
                                                    tp_end = tp.find(f'{{{self.NS_A}}}endParaRPr')
                                                    op_end = op.find(f'{{{self.NS_A}}}endParaRPr')
                                                    if tp_end is not None and op_end is not None:
                                                        for attr, val in op_end.attrib.items():
                                                            tp_end.set(attr, val)
                                                    tmpl_runs = tp.findall(f'{{{self.NS_A}}}r')
                                                    out_runs = op.findall(f'{{{self.NS_A}}}r')
                                                    for tr, or_ in zip(tmpl_runs, out_runs):
                                                        tr_rPr = tr.find(f'{{{self.NS_A}}}rPr')
                                                        or_rPr = or_.find(f'{{{self.NS_A}}}rPr')
                                                        if tr_rPr is not None and or_rPr is not None:
                                                            tr.remove(tr_rPr)
                                                            tr.insert(0, copy.deepcopy(or_rPr))
                                            break
                            found = True
                            break
                if not found:
                    print(f"WARNING: Placeholder '{placeholder}' not found on slide {slide_num} in template.")

            with zipfile.ZipFile(output_template_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename in [f'ppt/slides/slide{num}.xml' for num in modified_slides]:
                        slide_num = int(item.filename[len('ppt/slides/slide'):-4])
                        zout.writestr(item, ET.tostring(modified_slides[slide_num], encoding='UTF-8', xml_declaration=True))
                    else:
                        zout.writestr(item, zin.read(item.filename))

        return output_template_path
    
    def update_template(self, output_template_path=None):
        
        #Update Tracker
        self.tracker.update()
        
        if output_template_path is None:
            base, ext = os.path.splitext(self.template_path)
            output_template_path = f"{base}_adjusted{ext}"
    
        if not any(r.get('output_geometry') for r in self.tracker.replacements):
            print("WARNING: No output geometries found. Call tracker.update() first.")
            shutil.copy2(self.template_path, output_template_path)
            return output_template_path
    
        # Pre‑load output slides
        output_slides = {}
        with zipfile.ZipFile(self.tracker.output_path, 'r') as zout:
            for r in self.tracker.replacements:
                slide_file = f'ppt/slides/slide{r["slide"]}.xml'
                if slide_file not in output_slides:
                    output_slides[slide_file] = ET.fromstring(zout.read(slide_file))
    
        with zipfile.ZipFile(self.template_path, 'r') as zin:
            modified_slides = {}
            for r in self.tracker.replacements:
                out_geom = r.get('output_geometry')
                if out_geom is None:
                    continue
                slide_num = r['slide']
                slide_file = f'ppt/slides/slide{slide_num}.xml'
    
                if slide_num not in modified_slides:
                    slide_xml = zin.read(slide_file)
                    slide_root = ET.fromstring(slide_xml)
                    modified_slides[slide_num] = slide_root
    
                spTree = modified_slides[slide_num].find(f'{{{self.NS_P}}}cSld/{{{self.NS_P}}}spTree')
                if spTree is None:
                    continue
    
                placeholder = r['placeholder_text']
                found = False
                for shape in spTree.iter():
                    tag = shape.tag.split('}')[-1]
                    if tag not in ('sp', 'graphicFrame'):
                        continue
                    txBody = shape.find(f'{{{self.NS_P}}}txBody')
                    if txBody is None:
                        continue
    
                    texts = [t.text or '' for t in txBody.findall(f'.//{{{self.NS_A}}}t')]
                    full_text = ''.join(texts)
                    if placeholder not in full_text:
                        continue
    
                    # ---- 1. Move / resize ----
                    xfrm = shape.find(f'{{{self.NS_P}}}spPr/{{{self.NS_A}}}xfrm')
                    if xfrm is not None:
                        off = xfrm.find(f'{{{self.NS_A}}}off')
                        ext = xfrm.find(f'{{{self.NS_A}}}ext')
                        off.set('x', str(out_geom['left']))
                        off.set('y', str(out_geom['top']))
                        ext.set('cx', str(out_geom['width']))
                        ext.set('cy', str(out_geom['height']))
    
                    # ---- 2. Inherit formatting from output shape ----
                    if slide_file in output_slides:
                        out_root = output_slides[slide_file]
                        out_spTree = out_root.find(f'{{{self.NS_P}}}cSld/{{{self.NS_P}}}spTree')
                        if out_spTree is not None:
                            for out_shape in out_spTree.iter():
                                # 🔧 FIXED: recursively find cNvPr (was `find`, now `findall` with `.//`)
                                out_cNvPr = out_shape.find(f'.//{{{self.NS_P}}}cNvPr')
                                if out_cNvPr is not None and out_cNvPr.get('name') == r['picture_name']:
                                    out_txBody = out_shape.find(f'.//{{{self.NS_P}}}txBody')
                                    if out_txBody is not None:
                                        # --- Body properties ---
                                        tmpl_bodyPr = txBody.find(f'{{{self.NS_A}}}bodyPr')
                                        out_bodyPr = out_txBody.find(f'{{{self.NS_A}}}bodyPr')
                                        if tmpl_bodyPr is not None and out_bodyPr is not None:
                                            tmpl_bodyPr.attrib.clear()
                                            tmpl_bodyPr.attrib.update(out_bodyPr.attrib)
                                            for child in list(tmpl_bodyPr):
                                                tmpl_bodyPr.remove(child)
                                            for child in out_bodyPr:
                                                tmpl_bodyPr.append(copy.deepcopy(child))
    
                                        # --- Paragraph and run formatting ---
                                        tmpl_paras = txBody.findall(f'{{{self.NS_A}}}p')
                                        out_paras = out_txBody.findall(f'{{{self.NS_A}}}p')
                                        for tp, op in zip(tmpl_paras, out_paras):
                                            # paragraph properties (pPr)
                                            tp_pPr = tp.find(f'{{{self.NS_A}}}pPr')
                                            op_pPr = op.find(f'{{{self.NS_A}}}pPr')
                                            if op_pPr is not None:
                                                if tp_pPr is None:
                                                    tp_pPr = Element(f'{{{self.NS_A}}}pPr')
                                                    tp.insert(0, tp_pPr)
                                                for child in list(tp_pPr):
                                                    tp_pPr.remove(child)
                                                for child in op_pPr:
                                                    tp_pPr.append(copy.deepcopy(child))
                                                tp_pPr.attrib.clear()
                                                tp_pPr.attrib.update(op_pPr.attrib)
                                            elif tp_pPr is not None:
                                                tp.remove(tp_pPr)
    
                                            # end paragraph run properties
                                            tp_end = tp.find(f'{{{self.NS_A}}}endParaRPr')
                                            op_end = op.find(f'{{{self.NS_A}}}endParaRPr')
                                            if op_end is not None:
                                                if tp_end is None:
                                                    tp_end = Element(f'{{{self.NS_A}}}endParaRPr')
                                                    tp.append(tp_end)
                                                for child in list(tp_end):
                                                    tp_end.remove(child)
                                                for child in op_end:
                                                    tp_end.append(copy.deepcopy(child))
                                                tp_end.attrib.clear()
                                                tp_end.attrib.update(op_end.attrib)
                                            elif tp_end is not None:
                                                tp.remove(tp_end)
    
                                            # runs
                                            tmpl_runs = tp.findall(f'{{{self.NS_A}}}r')
                                            out_runs = op.findall(f'{{{self.NS_A}}}r')
                                            for tr, or_ in zip(tmpl_runs, out_runs):
                                                tr_rPr = tr.find(f'{{{self.NS_A}}}rPr')
                                                or_rPr = or_.find(f'{{{self.NS_A}}}rPr')
                                                if or_rPr is not None:
                                                    if tr_rPr is None:
                                                        tr_rPr = Element(f'{{{self.NS_A}}}rPr')
                                                        tr.insert(0, tr_rPr)
                                                    for child in list(tr_rPr):
                                                        tr_rPr.remove(child)
                                                    for child in or_rPr:
                                                        tr_rPr.append(copy.deepcopy(child))
                                                    tr_rPr.attrib.clear()
                                                    tr_rPr.attrib.update(or_rPr.attrib)
                                                elif tr_rPr is not None:
                                                    tr.remove(tr_rPr)
                                    break   # output shape matched
                    found = True
                    break
    
                if not found:
                    print(f"WARNING: Placeholder '{placeholder}' not found on slide {slide_num} in template.")
    
            # ---- Write the new template ----
            with zipfile.ZipFile(output_template_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename in [f'ppt/slides/slide{num}.xml' for num in modified_slides]:
                        slide_num = int(item.filename[len('ppt/slides/slide'):-4])
                        zout.writestr(item, ET.tostring(modified_slides[slide_num], encoding='UTF-8', xml_declaration=True))
                    else:
                        zout.writestr(item, zin.read(item.filename))
    
        return output_template_path

    def save(self, output_path):
        shutil.copy2(self.tmp_path, output_path)
        self.tracker.output_path = output_path

    def cleanup(self):
        self._tmp_dir.cleanup()


class ReplacementTracker:
    def __init__(self):
        self.replacements = []
        self.template_path = None
        self.output_path = None

    def add_replacement(self, slide_num, placeholder_text, picture_name,
                        template_position, template_size, replacement_type='image'):
        if not placeholder_text.startswith('{'):
            placeholder_text = f'{{{placeholder_text}}}'
        self.replacements.append({
            'slide': slide_num,
            'placeholder_text': placeholder_text,
            'picture_name': picture_name,
            'replacement_type': replacement_type,
            'template_geometry': {
                'left': template_position[0],
                'top': template_position[1],
                'width': template_size[0],
                'height': template_size[1]
            },
            'output_geometry': None,
        })

    def update(self, output_path=None):
        path = output_path or self.output_path
        if path is None:
            raise ValueError("No output path provided.")
        geoms = PowerPointTemplate._extract_all_geoms(path)
        for r in self.replacements:
            key = (r['slide'], r['picture_name'])
            if key in geoms:
                l, t, w, h = geoms[key]
                r['output_geometry'] = {'left': l, 'top': t, 'width': w, 'height': h}

    def print_report(self):
        print("\n" + "=" * 60)
        print("REPLACEMENT TRACKER")
        print("=" * 60)
        for i, r in enumerate(self.replacements, 1):
            status = "adjusted" if r['output_geometry'] else "unchanged"
            print(f"  {i}. Slide {r['slide']}: '{r['picture_name']}' ({r['replacement_type']}) – {status}")
            print(f"     Template: ({r['template_geometry']['left']}, {r['template_geometry']['top']}) "
                  f"{r['template_geometry']['width']}×{r['template_geometry']['height']}")
            if r['output_geometry']:
                print(f"     Output:   ({r['output_geometry']['left']}, {r['output_geometry']['top']}) "
                      f"{r['output_geometry']['width']}×{r['output_geometry']['height']}")
        print("=" * 60)

    def save(self, path):
        payload = {
            "template_path": self.template_path,
            "output_path": self.output_path,
            "replacements": self.replacements,
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)

    @classmethod
    def load(cls, path):
        with open(path, "r", encoding="utf-8") as f:
            payload = json.load(f)
        obj = cls()
        obj.template_path = payload.get("template_path")
        obj.output_path = payload.get("output_path")
        obj.replacements = payload.get("replacements", [])
        return obj

# ---------- Self-contained example ----------
if __name__ == '__main__':
    # Create a simple matplotlib plot
    import matplotlib.pyplot as plt
    # from RBNZ_Toolbox import aplot
    
    import matplotlib.pyplot as plt
    import numpy as np
    
    # Create first figure (line plot)
    x1 = np.linspace(0, 10, 100)
    y1 = np.sin(x1)
    fig1, ax1 = plt.subplots(figsize=(6, 4))
    ax1.plot(x1, y1, color='blue', linewidth=2)
    ax1.set_title('Sine Wave')
    ax1.set_xlabel('x')
    ax1.set_ylabel('sin(x)')
    ax1.grid(True, alpha=0.3)
    
    # Create second figure (bar chart)
    categories = ['A', 'B', 'C', 'D']
    values = [23, 45, 56, 78]
    fig2, ax2 = plt.subplots(figsize=(6, 4))
    ax2.bar(categories, values, color=['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728'])
    ax2.set_title('Sample Bar Chart')
    ax2.set_xlabel('Category')
    ax2.set_ylabel('Value')
    ax2.grid(axis='y', alpha=0.3)

    # Load template and perform replacements
    ppt = PowerPointTemplate(r'C:/Development/Powerpoint_Automation/Test.pptx')

    # Replace image placeholder {Chart 1} with the matplotlib figure
    ppt.replace_image({
        '{Chart 1}': fig1,
        '{Chart 2}': fig2
    })

    # Replace text placeholder {Title} with some text
    ppt.replace_text({
        'Title1': 'My Adjusted Title'
    })

    # Save the output
    ppt.save(r'C:/Development/Powerpoint_Automation/Test3.pptx')
    print("Saved output:")

    # Show the tracker report – you'll see both the image and text replacements recorded
    ppt.tracker.print_report()

    # Now simulate that the output was manually edited (positions changed).
    # In a real workflow, you'd open output.pptx, move/resize the shapes,
    # then call update_template with that edited output.
    # For demonstration, we'll just use the same output (no actual changes).
    # The tracker will read the output geometries and prepare to adjust the template.
    # ppt.tracker.update()  # This will now do nothing because the output is the same as saved.
    ppt.tracker.print_report()
    # input("")
    ppt.update_template(r'C:/Development/Powerpoint_Automation/New_Template.pptx')
    # print("Updated template saved:", new_template_path)