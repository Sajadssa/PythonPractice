import xml.etree.ElementTree as ET
from xml.dom import minidom
import uuid
import time

class ImprovedQCWorkflowGenerator:
    """
    نسخه بدون duplicate ID - با UUID یکتا
    """
    
    def __init__(self):
        # استفاده از timestamp برای یکتایی بیشتر
        self.timestamp = int(time.time() * 1000)
        self.cell_counter = 0
        self.cells = []
        
        # پالت رنگی
        self.colors = {
            'primary_blue': '#E3F2FD',
            'border_blue': '#1976D2',
            'light_gray': '#F5F5F5',
            'border_gray': '#424242',
            'green_fill': '#E8F5E9',
            'border_green': '#388E3C',
            'orange_fill': '#FFF3E0',
            'border_orange': '#F57C00',
            'purple_fill': '#EDE7F6',
            'border_purple': '#7B1FA2',
            'white': '#FFFFFF',
            'text_dark': '#212121',
            'arrow_solid': '#424242',
            'arrow_dashed': '#757575',
        }
        
    def get_unique_id(self, prefix='cell'):
        """
        تولید ID کاملاً یکتا
        ترکیب: prefix + timestamp + counter + random
        """
        self.cell_counter += 1
        # استفاده از UUID برای اطمینان کامل از عدم تکرار
        unique_part = str(uuid.uuid4())[:8]
        return f"{prefix}_{self.timestamp}_{self.cell_counter}_{unique_part}"
    
    def create_box(self, x, y, width, height, text, fill_color, border_color,
                   rounded=True, shadow=False, font_size=11, bold=False):
        """ایجاد باکس با ID یکتا"""
        cell_id = self.get_unique_id('box')
        
        style_parts = [
            "whiteSpace=wrap",
            "html=1",
            f"fillColor={fill_color}",
            f"strokeColor={border_color}",
            "strokeWidth=2",
            f"fontSize={font_size}",
            f"fontColor={self.colors['text_dark']}",
            "align=center",
            "verticalAlign=middle",
        ]
        
        if rounded:
            style_parts.insert(0, "rounded=1")
        
        if shadow:
            style_parts.append("shadow=1")
        
        if bold:
            style_parts.append("fontStyle=1")
        
        style = ";".join(style_parts) + ";"
        
        cell = {
            'id': cell_id,
            'value': text.replace('\n', '<br>'),
            'style': style,
            'vertex': '1',
            'parent': '1',
            'geometry': {'x': x, 'y': y, 'width': width, 'height': height}
        }
        self.cells.append(cell)
        return cell_id
    
    def create_list_box(self, x, y, width, height, title, items, fill_color, border_color):
        """ایجاد لیست با ID یکتا"""
        cell_id = self.get_unique_id('list')
        
        html = f'<div style="padding:8px;">'
        if title:
            html += f'<div style="font-weight:bold;margin-bottom:8px;font-size:12px;">{title}</div>'
        
        for i, item in enumerate(items, 1):
            html += f'<div style="margin:3px 0;font-size:10px;">{i}. {item}</div>'
        
        html += '</div>'
        
        style = (
            f"rounded=1;whiteSpace=wrap;html=1;"
            f"fillColor={fill_color};"
            f"strokeColor={border_color};"
            f"strokeWidth=2;"
            f"align=left;verticalAlign=top;"
        )
        
        cell = {
            'id': cell_id,
            'value': html,
            'style': style,
            'vertex': '1',
            'parent': '1',
            'geometry': {'x': x, 'y': y, 'width': width, 'height': height}
        }
        self.cells.append(cell)
        return cell_id
    
    def create_cylinder(self, x, y, width, height, text, fill_color, border_color):
        """دیتابیس - شکل سیلندر با ID یکتا"""
        cell_id = self.get_unique_id('cylinder')
        
        style = (
            f"shape=cylinder3;whiteSpace=wrap;html=1;"
            f"boundedLbl=1;backgroundOutline=1;size=15;"
            f"fillColor={fill_color};"
            f"strokeColor={border_color};"
            f"strokeWidth=2;"
            f"fontSize=13;fontStyle=1;"
            f"fontColor={self.colors['text_dark']};"
        )
        
        cell = {
            'id': cell_id,
            'value': text,
            'style': style,
            'vertex': '1',
            'parent': '1',
            'geometry': {'x': x, 'y': y, 'width': width, 'height': height}
        }
        self.cells.append(cell)
        return cell_id
    
    def create_arrow(self, source_id, target_id, label="", dashed=False):
        """ایجاد فلش با ID یکتا"""
        edge_id = self.get_unique_id('arrow')
        
        color = self.colors['arrow_dashed'] if dashed else self.colors['arrow_solid']
        
        style_parts = [
            "edgeStyle=orthogonalEdgeStyle",
            "rounded=1",
            "orthogonalLoop=1",
            "jettySize=auto",
            "html=1",
            "endArrow=classic",
            "endFill=1",
            f"strokeColor={color}",
            "strokeWidth=2",
        ]
        
        if dashed:
            style_parts.extend(["dashed=1", "dashPattern=5 5"])
        
        edge_style = ";".join(style_parts) + ";"
        
        cell = {
            'id': edge_id,
            'value': label,
            'style': edge_style,
            'edge': '1',
            'parent': '1',
            'source': source_id,
            'target': target_id,
            'geometry': {'relative': '1', 'as': 'geometry'}
        }
        self.cells.append(cell)
        return edge_id
    
    def create_title(self, x, y, width, height):
        """عنوان اصلی با ID یکتا"""
        cell_id = self.get_unique_id('title')
        
        html = (
            '<div style="text-align:center;">'
            '<div style="font-size:22px;font-weight:bold;margin-bottom:5px;">'
            'Quality Control Process Workflow'
            '</div>'
            '<div style="font-size:11px;font-style:italic;color:#666;">'
            'Process Flow Diagram'
            '</div>'
            '</div>'
        )
        
        style = (
            f"rounded=0;whiteSpace=wrap;html=1;"
            f"fillColor=none;"
            f"strokeColor=none;"
            f"align=center;verticalAlign=middle;"
        )
        
        cell = {
            'id': cell_id,
            'value': html,
            'style': style,
            'vertex': '1',
            'parent': '1',
            'geometry': {'x': x, 'y': y, 'width': width, 'height': height}
        }
        self.cells.append(cell)
        return cell_id
    
    def validate_unique_ids(self):
        """بررسی یکتا بودن IDها"""
        all_ids = [cell['id'] for cell in self.cells]
        duplicates = [id for id in all_ids if all_ids.count(id) > 1]
        
        if duplicates:
            print(f"⚠️  هشدار! ID های تکراری پیدا شد: {set(duplicates)}")
            return False
        else:
            print(f"✅ همه {len(all_ids)} ID یکتا هستند")
            return True
    
    def generate_workflow(self):
        """تولید دیاگرام کامل"""
        
        print("🎨 شروع ساخت دیاگرام...")
        
        # ===== عنوان =====
        title = self.create_title(250, 10, 500, 60)
        
        # ===== ستون چپ =====
        start = self.create_box(60, 80, 90, 40, "Start", 
                               self.colors['green_fill'], self.colors['border_green'],
                               True, False, 12, True)
        
        qc_plan = self.create_box(45, 135, 120, 40, "QC Plan",
                                 self.colors['primary_blue'], self.colors['border_blue'],
                                 True, False, 12, True)
        
        inspection = self.create_box(100, 180, 80, 25, "Inspection",
                                    self.colors['primary_blue'], self.colors['border_blue'],
                                    True, False, 10, True)
        
        qc_items_list = [
            "THK Measurement",
            "Painting Check", 
            "CIP Check",
            "QC & C Check",
            "Cathode Protection",
            "Lab. Result",
            "Calibration",
            "Pipeline Patrolling",
            "Internal/External Tank",
            "Internal/External Vessel",
            "PIG Driving",
            "Material Request"
        ]
        qc_items = self.create_list_box(35, 215, 140, 265, "", qc_items_list,
                                       self.colors['white'], self.colors['border_gray'])
        
        construction_list = [
            "New Construction",
            "Punch Clearance",
            "Overhaul",
            "Civil Work"
        ]
        construction = self.create_list_box(35, 495, 140, 100, "Construction:", 
                                           construction_list,
                                           self.colors['orange_fill'], self.colors['border_orange'])
        
        subsurface_list = [
            "Well Test",
            "Well Service",
            "Equipment"
        ]
        subsurface = self.create_list_box(35, 610, 140, 85, "Subsurface:",
                                         subsurface_list,
                                         self.colors['purple_fill'], self.colors['border_purple'])
        
        procedures = self.create_box(75, 710, 110, 40, "Procedures",
                                    self.colors['white'], self.colors['border_gray'],
                                    True, False, 11, True)
        
        # ===== ستون وسط چپ =====
        rfi = self.create_box(195, 295, 45, 25, "RFI",
                             self.colors['primary_blue'], self.colors['border_blue'],
                             True, False, 10, True)
        
        tpi = self.create_box(220, 325, 130, 65, "TPI\nInspection",
                             self.colors['primary_blue'], self.colors['border_blue'],
                             True, False, 13, True)
        
        qc_report = self.create_box(270, 240, 90, 30, "QC Report",
                                   self.colors['white'], self.colors['border_gray'],
                                   True, False, 11, False)
        
        # ===== ستون مرکز =====
        idms_list = [
            "Process",
            "KPIs",
            "Management Status",
            "QC Dashboard"
        ]
        idms = self.create_list_box(380, 165, 130, 115, "IDMS:",
                                   idms_list,
                                   self.colors['primary_blue'], self.colors['border_blue'])
        
        rbi = self.create_cylinder(390, 360, 110, 75, "RBI",
                                   self.colors['green_fill'], self.colors['border_green'])
        
        engineering = self.create_box(260, 490, 140, 85, "Engineering",
                                     self.colors['primary_blue'], self.colors['border_blue'],
                                     True, False, 13, True)
        
        tech_list = ["Technical Office", "I.TP"]
        technical = self.create_list_box(435, 505, 115, 60, "",
                                        tech_list,
                                        self.colors['white'], self.colors['border_gray'])
        
        moc_list = ["MOC", "Modification Of Design"]
        moc = self.create_list_box(275, 620, 135, 60, "",
                                  moc_list,
                                  self.colors['white'], self.colors['border_gray'])
        
        # ===== ستون راست =====
        wrfm = self.create_box(615, 145, 105, 85, "WRFM",
                              self.colors['primary_blue'], self.colors['border_blue'],
                              True, False, 14, True)
        
        min_list = [
            "Cost Control",
            "Preventive Corrosion",
            "Preventive Of Material",
            "Anti-Corrosion"
        ]
        min_inspection = self.create_list_box(795, 105, 140, 120, "Min Inspection:",
                                             min_list,
                                             self.colors['primary_blue'], self.colors['border_blue'])
        
        iow = self.create_box(640, 285, 70, 55, "IOW",
                             self.colors['white'], self.colors['border_gray'],
                             True, False, 12, True)
        
        rbi_module = self.create_box(860, 305, 90, 25, "RBI Module",
                                    self.colors['white'], self.colors['border_gray'],
                                    True, False, 10, False)
        
        material_list = ["IN", "I/R", "IRN"]
        material = self.create_list_box(560, 510, 125, 85, "Material Check:",
                                       material_list,
                                       self.colors['white'], self.colors['border_gray'])
        
        warehouse = self.create_box(715, 510, 105, 65, "Warehouse",
                                   self.colors['orange_fill'], self.colors['border_orange'],
                                   True, False, 13, True)
        
        cmms = self.create_cylinder(655, 390, 100, 65, "CMMS",
                                   self.colors['green_fill'], self.colors['border_green'])
        
        rca_list = ["PM change", "Min CM", "Failure Analysis"]
        rca = self.create_list_box(805, 365, 125, 90, "RCA:",
                                  rca_list,
                                  self.colors['white'], self.colors['border_gray'])
        
        scope = self.create_box(805, 610, 125, 65, "Scope Of Work",
                               self.colors['purple_fill'], self.colors['border_purple'],
                               True, False, 12, True)
        
        # ===== اتصالات =====
        print("🔗 ایجاد اتصالات...")
        
        self.create_arrow(start, qc_plan)
        self.create_arrow(qc_items, tpi)
        self.create_arrow(construction, tpi)
        self.create_arrow(tpi, idms)
        self.create_arrow(qc_report, idms)
        self.create_arrow(idms, rbi)
        self.create_arrow(idms, wrfm)
        self.create_arrow(wrfm, min_inspection)
        self.create_arrow(wrfm, iow, dashed=True)
        self.create_arrow(iow, wrfm, dashed=True)
        self.create_arrow(rbi, engineering)
        self.create_arrow(rbi, cmms, dashed=True)
        self.create_arrow(cmms, rbi, dashed=True)
        self.create_arrow(rbi, iow, dashed=True)
        self.create_arrow(engineering, technical)
        self.create_arrow(engineering, moc)
        self.create_arrow(engineering, procedures)
        self.create_arrow(technical, material)
        self.create_arrow(material, warehouse)
        self.create_arrow(warehouse, cmms)
        self.create_arrow(cmms, rca)
        self.create_arrow(moc, scope, dashed=True)
        self.create_arrow(warehouse, scope, dashed=True)
        self.create_arrow(subsurface, procedures)
        self.create_arrow(procedures, engineering)
        
        print(f"✅ {len(self.cells)} المان ایجاد شد")
        
        # بررسی یکتا بودن IDها
        self.validate_unique_ids()
    
    def generate_xml(self):
        """تولید XML با ID های یکتا"""
        mxfile = ET.Element('mxfile', {
            'host': 'app.diagrams.net',
            'modified': '2025-11-27T00:00:00.000Z',
            'agent': 'QC Workflow Generator v2.0',
            'version': '24.0.0',
            'type': 'device'
        })
        
        diagram = ET.SubElement(mxfile, 'diagram', {
            'name': 'QC Workflow',
            'id': f'qc_workflow_{self.timestamp}'
        })
        
        mxGraphModel = ET.SubElement(diagram, 'mxGraphModel', {
            'dx': '1400',
            'dy': '900',
            'grid': '1',
            'gridSize': '10',
            'guides': '1',
            'tooltips': '1',
            'connect': '1',
            'arrows': '1',
            'fold': '1',
            'page': '1',
            'pageScale': '1',
            'pageWidth': '1000',
            'pageHeight': '800',
            'background': '#FFFFFF',
            'math': '0',
            'shadow': '0'
        })
        
        root = ET.SubElement(mxGraphModel, 'root')
        
        # پارنت اصلی با ID یکتا
        root_id = f'root_{self.timestamp}_0'
        parent_id = f'parent_{self.timestamp}_1'
        
        ET.SubElement(root, 'mxCell', {'id': root_id})
        ET.SubElement(root, 'mxCell', {'id': parent_id, 'parent': root_id})
        
        # تغییر parent همه سلول‌ها به parent_id جدید
        for cell in self.cells:
            cell['parent'] = parent_id
            
            cell_attrs = {
                'id': cell['id'],
                'value': cell['value'],
                'style': cell['style'],
                'parent': cell['parent']
            }
            
            if 'vertex' in cell:
                cell_attrs['vertex'] = cell['vertex']
            if 'edge' in cell:
                cell_attrs['edge'] = cell['edge']
            if 'source' in cell:
                cell_attrs['source'] = cell['source']
            if 'target' in cell:
                cell_attrs['target'] = cell['target']
            
            cell_elem = ET.SubElement(root, 'mxCell', cell_attrs)
            
            geom = cell['geometry']
            geom_attrs = {'as': 'geometry'}
            
            if 'x' in geom:
                geom_attrs.update({
                    'x': str(geom['x']),
                    'y': str(geom['y']),
                    'width': str(geom['width']),
                    'height': str(geom['height'])
                })
            
            if 'relative' in geom:
                geom_attrs['relative'] = geom['relative']
            
            ET.SubElement(cell_elem, 'mxGeometry', geom_attrs)
        
        return mxfile
    
    def save_to_file(self, filename='qc_workflow_fixed.drawio'):
        """ذخیره فایل با بررسی ID"""
        print("\n" + "="*70)
        print("🎨 QC WORKFLOW GENERATOR - نسخه بدون Duplicate ID")
        print("="*70 + "\n")
        
        self.generate_workflow()
        
        # بررسی نهایی
        if not self.validate_unique_ids():
            print("❌ خطا در تولید IDها!")
            return None
        
        xml_tree = self.generate_xml()
        
        xml_str = ET.tostring(xml_tree, encoding='unicode')
        dom = minidom.parseString(xml_str)
        pretty_xml = dom.toprettyxml(indent='  ')
        pretty_xml = '\n'.join([line for line in pretty_xml.split('\n') if line.strip()])
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(pretty_xml)
        
        print(f"\n✅ فایل {filename} با موفقیت ذخیره شد!")
        print(f"📊 تعداد المان‌ها: {len(self.cells)}")
        print(f"💾 حجم فایل: {len(pretty_xml):,} بایت")
        print(f"🔐 Timestamp: {self.timestamp}\n")
        
        self.print_usage()
        
        return filename
    
    def print_usage(self):
        """راهنما"""
        print("="*70)
        print("📖 نحوه استفاده:\n")
        print("1️⃣  به https://app.diagrams.net بروید")
        print("2️⃣  File → Open from → Device")
        print("3️⃣  فایل qc_workflow_fixed.drawio را باز کنید")
        print("\n🔧 مشکل ID Duplicate حل شد:")
        print("   ✓ استفاده از UUID یکتا")
        print("   ✓ Timestamp برای جلوگیری از تداخل")
        print("   ✓ Counter جداگانه برای هر المان")
        print("   ✓ بررسی خودکار یکتا بودن")
        print("\n💡 قابلیت‌ها:")
        print("   • هر بار اجرا IDهای کاملاً متفاوت")
        print("   • بدون تداخل با دیاگرام‌های قبلی")
        print("   • سازگار 100% با Draw.io")
        print("="*70 + "\n")

# اجرا
if __name__ == "__main__":
    generator = ImprovedQCWorkflowGenerator()
    generator.save_to_file('qc_workflow_fixed.drawio')
