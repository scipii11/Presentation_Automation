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

    # Supported image formats and their MIME types
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
        """Initialize with a template PPTX file."""
        self.template_path = template_path
        self.tracker = ReplacementTracker()
        
        # Create temp working directory for this instance
        self._tmp_dir = tempfile.TemporaryDirectory()
        self.tmp_path = os.path.join(self._tmp_dir.name, f"{hash(time.time())}_working.pptx")
        
        # Initialize tmp file as a copy of template
        shutil.copy2(self.template_path, self.tmp_path)
        
        # Set tracker template path
        self.tracker.template_path = template_path
        
        # Register namespaces
        for prefix, uri in self.NSMAP.items():
            ET.register_namespace(prefix, uri)
            
    def __del__(self):
        self.cleanup()


    # ========== Image Format Helpers ==========
    @classmethod
    def _get_image_format(cls, file_path_or_bytes, filename=None):
        """
        Detect the image format from a file path, bytes, or filename.
        Returns (extension, mime_type) tuple.
        """
        # If it's a file path
        if isinstance(file_path_or_bytes, (str, Path)) and os.path.exists(str(file_path_or_bytes)):
            ext = os.path.splitext(str(file_path_or_bytes))[1].lower()
            if ext in cls.IMAGE_FORMATS:
                return ext, cls.IMAGE_FORMATS[ext]
        
        # If a filename is provided
        if filename:
            ext = os.path.splitext(filename)[1].lower()
            if ext in cls.IMAGE_FORMATS:
                return ext, cls.IMAGE_FORMATS[ext]
        
        # Try to detect from bytes
        if isinstance(file_path_or_bytes, bytes):
            # Check magic bytes for common formats
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
        
        # Default to PNG
        return '.png', 'image/png'

    @classmethod
    def _convert_to_png_if_needed(cls, image_data, target_format='.svg'):
        """
        Convert non-SVG images to PNG (since PowerPoint handles PNG better than SVG for photos).
        SVG images are kept as-is.
        """
        ext, mime = cls._get_image_format(image_data)
        
        # SVGs are already perfect for PowerPoint
        if ext == '.svg':
            return image_data, ext
        
        # For raster formats, we could convert using PIL if available
        # For now, return as-is since PowerPoint supports PNG/JPEG natively
        return image_data, ext

    @classmethod
    def _ensure_content_type(cls, root, extension, mime_type):
        """Ensure the content type for a given extension is declared."""
        for default in root.findall(f'{{{cls.NS_CT}}}Default'):
            if default.get('Extension') == extension.lstrip('.'):
                # Update if mime type differs
                if default.get('ContentType') != mime_type:
                    default.set('ContentType', mime_type)
                return
        default = Element(f'{{{cls.NS_CT}}}Default',
                          Extension=extension.lstrip('.'),
                          ContentType=mime_type)
        root.append(default)

    # ========== Matplotlib Helper ==========
    @staticmethod
    def _fig_to_image(fig, format='svg', **kwargs):
        """
        Convert a matplotlib figure to image bytes.
        
        Args:
            fig: matplotlib figure object
            format: 'svg' or 'png'
            **kwargs: dpi, transparent, bbox_inches, etc.
        
        Returns:
            (image_data, extension) tuple
        """
        buf = io.BytesIO()
        
        # Set default kwargs
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

    # ========== XML Helpers ==========
    @staticmethod
    def _next_image_num(zip_ref, prefix='ppt/media/image'):
        """Find the next available image number across all formats."""
        existing = [n for n in zip_ref.namelist() if n.startswith(prefix)]
        max_num = 0
        for name in existing:
            try:
                # Extract number from 'ppt/media/image123.svg' or 'ppt/media/image123.png'
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

    # ========== Geometry Resolution ==========
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

    def _replace_with_images(self, image_map):
        """Core replacement logic that handles multiple image formats."""
        with zipfile.ZipFile(self.tmp_path, 'r') as zin:
            with zin.open('[Content_Types].xml') as f:
                ct_tree = ET.parse(f)
            
            # Ensure all necessary content types are declared
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
                        
                        # Image filename in the zip
                        img_filename = f'ppt/media/image{img_num}{ext}'

                        rel_id = self._next_rel_id(rels_root)
                        rel = Element(f'{{{self.NS_REL}}}Relationship',
                                      Id=rel_id,
                                      Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                                      Target=f'../media/image{img_num}{ext}')
                        rels_root.append(rel)

                        pic = self._make_pic_element(left, top, width, height, rel_id, pic_name)
                        spTree.append(pic)

                        self.tracker.add_replacement(slide_num, text, pic_name,
                                                     (left, top), (width, height))

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
            
                # Write images
                img_assignments = {}
                for r in self.tracker.replacements:
                    img_num = int(r['picture_name'].split('_')[-1])
                    search_text = r['placeholder_text']
                    img_data, ext, _ = image_map[search_text]
                    img_assignments[img_num] = (img_data, ext)
            
                for img_num, (img_data, ext) in img_assignments.items():
                    zout.writestr(f'ppt/media/image{img_num}{ext}', img_data)

        # Replace original safely
        shutil.move(new_zip_path, self.tmp_path)


        return
    
    # ========== Main Operations ==========
    def replace(self, placeholder_map):
        pass
    
    def replace_image(self, placeholder_map:dict, **kwargs):
        """
        Replace placeholders with images. Automatically detects and handles:
        - SVG files (.svg) → inserted as vector graphics
        - PNG files (.png) → inserted as raster images
        - JPEG files (.jpg, .jpeg) → inserted as raster images
        - GIF files (.gif) → inserted as raster images
        - BMP files (.bmp) → inserted as raster images
        - TIFF files (.tiff, .tif) → inserted as raster images
        - WebP files (.webp) → inserted as raster images
        - Matplotlib figures → converted to SVG by default
        - Raw bytes → format auto-detected
        
        Args:
            placeholder_map: Dict mapping placeholder text to:
                - str/Path: path to an image file
                - matplotlib.figure.Figure: a matplotlib figure
                - bytes: raw image data
                - tuple: (bytes, filename) for format detection
            output_path: Output PPTX path
            **kwargs: For matplotlib figures: format='svg'|'png', dpi, transparent
        
        Returns:
            Path to output file
        
        Example:
            ppt = PowerPointTemplate('template.pptx')
            ppt.replace({
                '{Logo}': 'logo.svg',              # SVG file
                '{Photo}': 'photo.png',            # PNG file
                '{Chart}': matplotlib_figure,       # matplotlib figure
                '{Icon}': b'<svg>...</svg>',       # raw SVG bytes
                '{Graph}': png_bytes, # bytes with filename hint
            })
        """
        output_path = self.tmp_path
        
        image_map = {}  # {search_text: (image_data, extension, mime_type)}
        
        for search_text, value in placeholder_map.items():
            # Handle tuples: (bytes, filename)
            if isinstance(value, tuple) and len(value) == 2:
                data, filename = value
                ext, mime = self._get_image_format(data, filename)
                image_map[search_text] = (data, ext, mime)
            
            # Handle matplotlib figures
            elif hasattr(value, 'savefig'):
                fmt = kwargs.get('format', 'svg')
                data, ext = self._fig_to_image(value, format=fmt, **kwargs)
                mime = self.IMAGE_FORMATS.get(ext, 'image/png')
                image_map[search_text] = (data, ext, mime)
            
            # Handle raw bytes
            elif isinstance(value, bytes):
                ext, mime = self._get_image_format(value)
                image_map[search_text] = (value, ext, mime)
            
            # Handle file paths (str or Path)
            elif isinstance(value, (str, Path)):
                path = Path(value)
                ext = path.suffix.lower()
                mime = self.IMAGE_FORMATS.get(ext, 'image/png')
                with open(path, 'rb') as f:
                    data = f.read()
                image_map[search_text] = (data, ext, mime)
            
            else:
                raise TypeError(
                    f"Unsupported type for '{search_text}': {type(value)}. "
                    f"Expected str (file path), bytes, tuple, or matplotlib Figure."
                )
        
        return self._replace_with_images(image_map)
    
    
    def replace_text(self, placeholder_map: dict):
        """
        Replace text placeholders in a PowerPoint presentation.
    
        Args:
            placeholder_map: dict mapping keys -> replacement text
                  e.g. {"Title": "New Title"} replaces {Title}
            save_location: where to save output
            save: whether to save file
    
        Returns:
            Presentation object
        """
        from pptx import Presentation as PPTXPresentation
    
        # Load if a path is provided
        prs = PPTXPresentation(self.tmp_path)
    
        # Build placeholder mapping: {Key} -> value
        replacer = {f'{{{k}}}': str(v) for k, v in placeholder_map.items()}
    
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text:
                    # Only process shapes that actually contain placeholders
                    if "{" in shape.text and "}" in shape.text:
                        for old, new in replacer.items():
                            if old in shape.text:
                                shape.text = shape.text.replace(old, new)
    
        # Handle saving
        prs.save(self.tmp_path)
        
        return


    def update_template(self, adjusted_pptx_path, output_path=None):
        """Create a new template with text boxes at the adjusted positions."""
        if output_path is None:
            base, ext = os.path.splitext(self.template_path)
            output_path = f"{base}_adjusted{ext}"

        adjusted_geoms = {}
        with zipfile.ZipFile(adjusted_pptx_path, 'r') as zadj:
            slide_files = [n for n in zadj.namelist()
                           if n.startswith('ppt/slides/slide')
                           and n.endswith('.xml')
                           and '_rels' not in n]
            for slide_file in slide_files:
                slide_num = int(slide_file[len('ppt/slides/slide'):-4])
                slide_xml = zadj.read(slide_file)
                slide_root = ET.fromstring(slide_xml)
                for pic in slide_root.iter(f'{{{self.NS_P}}}pic'):
                    cNvPr = pic.find(f'{{{self.NS_P}}}nvPicPr/{{{self.NS_P}}}cNvPr')
                    if cNvPr is None:
                        continue
                    pic_name = cNvPr.get('name')
                    if not pic_name:
                        continue
                    xfrm = pic.find(f'{{{self.NS_P}}}spPr/{{{self.NS_A}}}xfrm')
                    if xfrm is None:
                        continue
                    off = xfrm.find(f'{{{self.NS_A}}}off')
                    ext = xfrm.find(f'{{{self.NS_A}}}ext')
                    if off is None or ext is None:
                        continue
                    left = int(off.get('x', '0'))
                    top = int(off.get('y', '0'))
                    width = int(ext.get('cx', '0'))
                    height = int(ext.get('cy', '0'))
                    adjusted_geoms[(slide_num, pic_name)] = (left, top, width, height)

        matches = 0
        for r in self.tracker.replacements:
            key = (r['slide'], r['picture_name'])
            if key in adjusted_geoms:
                l, t, w, h = adjusted_geoms[key]
                r['output_geometry'] = {'left': l, 'top': t, 'width': w, 'height': h}
                matches += 1

        if matches == 0:
            print("WARNING: No matches. Copying template unchanged.")
            shutil.copy2(self.template_path, output_path)
            return output_path

        with zipfile.ZipFile(self.template_path, 'r') as zin:
            with zin.open('[Content_Types].xml') as f:
                ct_tree = ET.parse(f)

            modified_slides = {}
            modified_rels = {}

            for r in self.tracker.replacements:
                if r['output_geometry'] is None:
                    continue
                slide_num = r['slide']
                slide_file = f'ppt/slides/slide{slide_num}.xml'
                rels_file = f'ppt/slides/_rels/slide{slide_num}.xml.rels'

                if slide_num not in modified_slides:
                    slide_xml = zin.read(slide_file)
                    slide_rels_xml = zin.read(rels_file) if rels_file in zin.namelist() else None
                    slide_root = ET.fromstring(slide_xml)
                    spTree = slide_root.find(f'{{{self.NS_P}}}cSld/{{{self.NS_P}}}spTree')
                    rels_root = ET.fromstring(slide_rels_xml) if slide_rels_xml else Element(f'{{{self.NS_REL}}}Relationships')
                    modified_slides[slide_num] = (slide_root, spTree)
                    modified_rels[slide_num] = rels_root

                slide_root, spTree = modified_slides[slide_num]

                found = False
                for shape, parent, *_, text, _ in self._find_text_shapes(
                    spTree, zin, slide_rels_xml, slide_num, r['placeholder_text']
                ):
                    parent.remove(shape)
                    found = True
                    break

                if not found:
                    continue

                out = r['output_geometry']
                new_textbox = self._make_textbox_element(
                    out['left'], out['top'], out['width'], out['height'],
                    r['placeholder_text']
                )
                spTree.append(new_textbox)

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

        return output_path
    
    def cleanup(self):
        self._tmp_dir.cleanup()
    
    def save(self, output_path):
        shutil.copy2(self.tmp_path, output_path)
        self.tracker.output_path = output_path


class ReplacementTracker:
    """Tracks which placeholders were replaced and their geometry."""
    

    def __init__(self):
        self.replacements = []
        self.template_path = None
        self.output_path = None
        
    @staticmethod
    def _extract_asset_geoms(pptx_path):
        import zipfile
        import xml.etree.ElementTree as ET
    
        NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
        NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    
        geoms = {}
    
        def get_off_ext(xfrm):
            off = xfrm.find(f"{{{NS_A}}}off")
            ext = xfrm.find(f"{{{NS_A}}}ext")
    
            if off is None or ext is None:
                return None
    
            return (
                int(off.get("x", 0)),
                int(off.get("y", 0)),
                int(ext.get("cx", 0)),
                int(ext.get("cy", 0)),
            )
    
        def walk(node, acc_x=0, acc_y=0):
    
            tag = node.tag.split("}")[-1]
    
            # -------------------------
            # GROUP (CRITICAL FIX)
            # -------------------------
            if tag == "grpSp":
    
                xfrm = node.find(f".//{{{NS_P}}}grpSpPr/{{{NS_A}}}xfrm")
                if xfrm is not None:
                    off = xfrm.find(f"{{{NS_A}}}off")
                    if off is not None:
                        acc_x += int(off.get("x", 0))
                        acc_y += int(off.get("y", 0))
    
                for child in node:
                    walk(child, acc_x, acc_y)
    
            # -------------------------
            # PICTURE
            # -------------------------
            elif tag == "pic":
    
                cNvPr = node.find(f".//{{{NS_P}}}cNvPr")
                xfrm = node.find(f".//{{{NS_P}}}spPr/{{{NS_A}}}xfrm")
    
                if cNvPr is None or xfrm is None:
                    return
    
                name = cNvPr.get("name")
    
                geom = get_off_ext(xfrm)
                if geom is None:
                    return
    
                x, y, w, h = geom
    
                # FINAL ACCUMULATED POSITION
                geoms[name] = (acc_x + x, acc_y + y, w, h)
    
            # -------------------------
            # RECURSE
            # -------------------------
            for child in node:
                walk(child, acc_x, acc_y)
    
        # -------------------------
        # READ PPTX
        # -------------------------
        with zipfile.ZipFile(pptx_path, "r") as z:
    
            slides = [
                n for n in z.namelist()
                if n.startswith("ppt/slides/slide")
                and n.endswith(".xml")
                and "_rels" not in n
            ]
    
            for slide_file in slides:
                slide_num = int(slide_file[len("ppt/slides/slide"):-4])
                root = ET.fromstring(z.read(slide_file))
    
                spTree = root.find(f".//{{{NS_P}}}spTree")
                if spTree is None:
                    continue
    
                walk(spTree)
    
                # attach slide number
                for k, v in list(geoms.items()):
                    geoms[(slide_num, k)] = v
                    del geoms[k]
    
        return geoms
    
    @staticmethod
    def _extract_text_geoms(pptx_path):
        import zipfile
        import xml.etree.ElementTree as ET
    
        NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
        NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    
        results = {}
    
        def get_xywh(node):
            xfrm = node.find(f".//{{{NS_P}}}xfrm")
            if xfrm is None:
                return None
    
            off = xfrm.find(f"{{{NS_A}}}off")
            ext = xfrm.find(f"{{{NS_A}}}ext")
    
            if off is None or ext is None:
                return None
    
            return (
                int(off.get("x", 0)),
                int(off.get("y", 0)),
                int(ext.get("cx", 0)),
                int(ext.get("cy", 0)),
            )
    
        def extract_text(node):
            return "".join(
                t.text or ""
                for t in node.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}t")
            ).strip()
    
        def walk(node, slide_num, acc_x=0, acc_y=0):
    
            tag = node.tag.split("}")[-1]
    
            # -------------------------
            # GROUPS
            # -------------------------
            if tag == "grpSp":
                xfrm = node.find(f".//{{{NS_P}}}grpSpPr/{{{NS_A}}}xfrm")
    
                if xfrm is not None:
                    off = xfrm.find(f"{{{NS_A}}}off")
                    if off is not None:
                        acc_x += int(off.get("x", 0))
                        acc_y += int(off.get("y", 0))
    
                for child in node:
                    walk(child, slide_num, acc_x, acc_y)
    
                return
    
            # -------------------------
            # SHAPES (normal textboxes)
            # -------------------------
            if tag == "sp":
                txBody = node.find(f".//{{{NS_P}}}txBody")
                if txBody is not None:
                    text = extract_text(txBody)
    
                    if text:
                        geom = get_xywh(node)
                        if geom:
                            x, y, w, h = geom
    
                            cNvPr = node.find(f".//{{{NS_P}}}cNvPr")
                            shape_id = cNvPr.get("id") if cNvPr is not None else None
                            name = cNvPr.get("name") if cNvPr is not None else None
    
                            key = (slide_num, shape_id or name)
    
                            results[key] = {
                                "text": text,
                                "name": name,
                                "left": acc_x + x,
                                "top": acc_y + y,
                                "width": w,
                                "height": h,
                            }
    
            # -------------------------
            # TABLES (CRITICAL FIX)
            # -------------------------
            if tag == "graphicFrame":
                text_nodes = node.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}t")
                if text_nodes:
                    text = "".join(t.text or "" for t in text_nodes).strip()
    
                    if text:
                        geom = get_xywh(node)
                        if geom:
                            x, y, w, h = geom
    
                            cNvPr = node.find(f".//{{{NS_P}}}cNvPr")
                            shape_id = cNvPr.get("id") if cNvPr is not None else None
                            name = cNvPr.get("name") if cNvPr is not None else None
    
                            key = (slide_num, shape_id or name or f"table_{len(results)}")
    
                            results[key] = {
                                "text": text,
                                "name": name,
                                "left": acc_x + x,
                                "top": acc_y + y,
                                "width": w,
                                "height": h,
                            }
    
            # -------------------------
            # CONTINUE WALK
            # -------------------------
            for child in node:
                walk(child, slide_num, acc_x, acc_y)
    
        # -------------------------
        # READ PPTX
        # -------------------------
        with zipfile.ZipFile(pptx_path, "r") as z:
            slides = [
                n for n in z.namelist()
                if n.startswith("ppt/slides/slide")
                and n.endswith(".xml")
                and "_rels" not in n
            ]
    
            for slide_file in slides:
                import re
                m = re.search(r"slide(\d+)\.xml$", slide_file)
                if not m:
                    continue
    
                slide_num = int(m.group(1))
    
                root = ET.fromstring(z.read(slide_file))
                spTree = root.find(f".//{{{NS_P}}}spTree")
    
                if spTree is None:
                    continue
    
                walk(spTree, slide_num)
    
        return results

    def add_replacement(self, slide_num, placeholder_text, picture_name,
                        template_position, template_size):
        self.replacements.append({
            'slide': slide_num,
            'placeholder_text': placeholder_text,
            'picture_name': picture_name,
            'template_geometry': {
                'left': template_position[0],
                'top': template_position[1],
                'width': template_size[0],
                'height': template_size[1]
            },
            'output_geometry': None,
        })

    def print_report(self):
        print("\n" + "=" * 60)
        print("REPLACEMENT TRACKER")
        print("=" * 60)
        for i, r in enumerate(self.replacements, 1):
            status = "adjusted" if r['output_geometry'] else "unchanged"
            print(f"  {i}. Slide {r['slide']}: '{r['picture_name']}' – {status}")
            print(f"     Template: ({r['template_geometry']['left']}, {r['template_geometry']['top']}) "
                  f"{r['template_geometry']['width']}×{r['template_geometry']['height']}")
            if r['output_geometry']:
                print(f"     Output:   ({r['output_geometry']['left']}, {r['output_geometry']['top']}) "
                      f"{r['output_geometry']['width']}×{r['output_geometry']['height']}")
        print("=" * 60)
        
    def update(self):
        """
        Reconcile replacements with actual output PPTX geometry.
        """
    
        if self.output_path is None:
            raise Exception("Save an output first to use as the example for the new template")
    
        import zipfile
        import xml.etree.ElementTree as ET
    
        NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
        NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    
        # ----------------------------
        # 1. Extract both PPTX states
        # ----------------------------
        template_geoms = self._extract_asset_geoms(self.template_path)
        output_geoms = self._extract_asset_geoms(self.output_path)
    
        # ----------------------------
        # 2. Update tracker
        # ----------------------------
        updated = 0
    
        for r in self.replacements:
    
            slide = r["slide"]
    
            # ----------------------------
            # PRIMARY MATCH (BEST)
            # ----------------------------
            key = (slide, r["picture_name"])
    
            if key in output_geoms:
                l, t, w, h = output_geoms[key]
    
                r["output_geometry"] = {
                    "left": l,
                    "top": t,
                    "width": w,
                    "height": h,
                }
    
                updated += 1
                continue
    
        return updated
    
    def save(self, path):
        """
        Save tracker state to disk (JSON).
        """
    
        import json
    
        payload = {
            "template_path": self.template_path,
            "output_path": self.output_path,
            "replacements": self.replacements,
        }
    
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
    
        return 
    
    @classmethod
    def load(cls, path):
        """
        Load tracker state from disk (JSON).
        Returns a ReplacementTracker instance.
        """
    
        import json
    
        with open(path, "r", encoding="utf-8") as f:
            payload = json.load(f)
    
        obj = cls()
        obj.template_path = payload.get("template_path")
        obj.output_path = payload.get("output_path")
        obj.replacements = payload.get("replacements", [])
    
        return obj


# ---------- Example usage ----------
if __name__ == '__main__':
    import matplotlib.pyplot as plt
    from RBNZ_Toolbox import aplot
    from RBNZ_Data import ffs
    
    data, ax = aplot(ffs.get('LVRN.MMB1.AA', table=True))
    
    ppt = PowerPointTemplate(r'C:/Development/Presentation_Automation/Test.pptx')
    
    out = ppt.replace_image({
        '{Chart 1}': ax['fig'],
        '{Chart 2}': ax['fig']
    }
        )
    
    out = ppt.replace_text({'Title1':'Title will be here'})
    
    ppt.save(r'C:/Development/Presentation_Automation/Test_Output.pptx')
    
    ppt.tracker.update()
    
    ppt.update_template(r'C:/Development/Presentation_Automation/Test_Output.pptx', r'C:\Development\Presentation_Automation/New_template1.pptx')
