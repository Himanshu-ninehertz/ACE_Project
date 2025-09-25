import fitz  # PyMuPDF
import ezdxf
import matplotlib.pyplot as plt
from matplotlib.patches import Circle
import re
from PIL import Image
import io
import os
import math
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image as RLImage, Spacer


POINTS_TO_MM = 25.4 / 72.0
BLUE_RGB = (0, 0, 255)
RED_RGB = (255, 0, 0)


def _rgb_from_fitz(color_tuple):
    if not color_tuple:
        return (0, 0, 0)
    return tuple(int(max(0, min(1, c)) * 255) for c in color_tuple[:3])


def _is_color_close(rgb, target, tol=80):
    return all(abs(rgb[i] - target[i]) <= tol for i in range(3))


def _polyline_length_points(points):
    total = 0.0
    for i in range(len(points)-1):
        dx = points[i+1][0] - points[i][0]
        dy = points[i+1][1] - points[i][1]
        total += math.hypot(dx, dy)
    return total


def detect_scale_from_text(text_dict):
    all_text = " ".join(
        s.get("text", "")
        for b in text_dict.get("blocks", [])
        for l in b.get("lines", [])
        for s in l.get("spans", [])
    )
    m = re.search(r'1\s*[:xX]\s*(\d{1,4})', all_text)
    if m:
        return float(m.group(1))
    return None


def _find_nearest_text_label(points, text_dict, max_dist_pts=60.0):
    if not points:
        return None
    cx = sum(p[0] for p in points) / len(points)
    cy = sum(p[1] for p in points) / len(points)
    nearest_text, nearest_d = None, float("inf")
    for b in text_dict.get("blocks", []):
        for l in b.get("lines", []):
            for s in l.get("spans", []):
                bbox = s.get("bbox", None)
                if not bbox:
                    continue
                tx = (bbox[0] + bbox[2]) / 2.0
                ty = (bbox[1] + bbox[3]) / 2.0
                d = math.hypot(tx - cx, ty - cy)
                if d < nearest_d and d <= max_dist_pts:
                    nearest_d = d
                    nearest_text = s.get("text", "").strip()
    return nearest_text


def _parse_size_from_text(txt):
    if not txt:
        return None
    m = re.search(r'(\d{2,4})\s*[mM][mM]', txt)
    if m:
        return f"{m.group(1)}mm"
    m2 = re.search(r'(\d{2,4})\s*ø', txt)
    if m2:
        return f"{m2.group(1)}mm"
    m3 = re.search(r'(\d{2,4})\s*[xX]\s*(\d{2,4})', txt)
    if m3:
        return f"{m3.group(1)}x{m3.group(2)}"
    return None


def crop_region_to_image(pdf_path, page_no, bbox, out_path, zoom=2.0):
    doc = fitz.open(pdf_path)
    page = doc[page_no]
    clip = fitz.Rect(*bbox)
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, clip=clip)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    img.save(out_path)
    doc.close()
    return out_path


def measure_ducts(pdf_path):
    doc = fitz.open(pdf_path)
    page = doc[0]
    drawings = page.get_drawings()
    text_dict = page.get_text("dict")
    doc.close()

    scale_den = detect_scale_from_text(text_dict) or 1.0

    results = {
        "scale_denominator": scale_den,
        "supply": {"lengths_mm": {}, "counts": {}, "images": {}},
        "extract": {"lengths_mm": {}, "counts": {}, "images": {}},
    }

    for drawing in drawings:
        rgb = _rgb_from_fitz(drawing.get("color"))
        role = "extract" if _is_color_close(rgb, BLUE_RGB) else "supply" if _is_color_close(rgb, RED_RGB) else None
        if not role:
            continue
        for itm in drawing.get("items", []):
            tag = itm[0]
            points, bbox = [], None
            if tag == "l":
                points = [itm[1], itm[2]]
                bbox = [min(p[0] for p in points), min(p[1] for p in points),
                        max(p[0] for p in points), max(p[1] for p in points)]
            elif tag == "re":
                rect = itm[1]
                points = [(rect.x0, rect.y0), (rect.x1, rect.y0), (rect.x1, rect.y1), (rect.x0, rect.y1)]
                bbox = [rect.x0, rect.y0, rect.x1, rect.y1]
            elif tag == "c":
                points = itm[1:]
                if points:
                    bbox = [min(p[0] for p in points), min(p[1] for p in points),
                            max(p[0] for p in points), max(p[1] for p in points)]
            if not points:
                continue

            length_pts = _polyline_length_points(points)
            length_mm = length_pts * POINTS_TO_MM * scale_den
            label = _find_nearest_text_label(points, text_dict)
            size = _parse_size_from_text(label) or "unknown"

            results[role]["lengths_mm"][size] = results[role]["lengths_mm"].get(size, 0.0) + length_mm
            if bbox:
                results[role]["images"].setdefault(size, []).append(bbox)

    return results


def export_results_to_excel(results, excel_path, pdf_path, out_img_dir="output_images"):
    os.makedirs(os.path.dirname(excel_path) or ".", exist_ok=True)
    os.makedirs(out_img_dir, exist_ok=True)

    writer = pd.ExcelWriter(excel_path, engine="xlsxwriter")
    workbook = writer.book

    for role in ("supply", "extract"):
        rows = []
        lengths = results[role]["lengths_mm"]
        for size, length in lengths.items():
            rows.append({"Type": "Ductwork", "Size": size, "Length_m": round(length/1000, 2)})
        for comp, count in results[role]["counts"].items():
            rows.append({"Type": comp, "Size": "", "Length_m": count})
        df = pd.DataFrame(rows)
        df.to_excel(writer, sheet_name=role.capitalize(), index=False)

        ws = writer.sheets[role.capitalize()]
        row = len(df) + 2
        for size, bboxes in results[role]["images"].items():
            ws.write(row, 0, f"Images for {size}")
            for i, bbox in enumerate(bboxes[:5]):
                out_path = os.path.join(out_img_dir, f"{role}_{size}_{i}.png")
                try:
                    crop_region_to_image(pdf_path, 0, bbox, out_path)
                    ws.insert_image(row, 1+i, out_path, {'x_scale': 0.4, 'y_scale': 0.4})
                except Exception as e:
                    ws.write(row, 1+i, f"(img error {e})")
            row += 15

    writer.close()
    return excel_path


def generate_pdf_report(results, output_path, pdf_path, out_img_dir="output_images"):
    """
    Generate a PDF report with measurement results and visualizations
    """
    doc = SimpleDocTemplate(output_path, pagesize=A4)
    story = []
    styles = getSampleStyleSheet()

    # Title
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Title'],
        fontSize=24,
        spaceAfter=30
    )
    title = Paragraph("Ductwork Measurement Report", title_style)
    story.append(title)

    # Summary Tables
    for section in ("supply", "extract"):
        # Section Header
        header_style = ParagraphStyle(
            'SectionHeader',
            parent=styles['Heading1'],
            fontSize=18,
            spaceAfter=10
        )
        story.append(Paragraph(f"{section.capitalize()} Measurements", header_style))

        # Ductwork Length Table
        lengths = results.get(section, {}).get("lengths_mm", {})
        if lengths:
            story.append(Paragraph("Ductwork Lengths", styles['Heading2']))
            data = [["Size", "Length (mm)", "Length (m)"]]
            for size, length_mm in lengths.items():
                data.append([size, f"{round(length_mm, 2)}", f"{round(length_mm/1000.0, 3)}"])
            
            table_style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 10),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ])
            
            t = Table(data, colWidths=[2*inch, 2*inch, 2*inch])
            t.setStyle(table_style)
            story.append(t)
            story.append(Spacer(1, 20))

        # Components Count Table
        counts = results.get(section, {}).get("counts", {})
        if counts:
            story.append(Paragraph("Component Counts", styles['Heading2']))
            data = [["Component", "Count"]]
            for name, count in counts.items():
                data.append([name, str(count)])
            
            t = Table(data, colWidths=[4*inch, 2*inch])
            t.setStyle(table_style)
            story.append(t)
            story.append(Spacer(1, 30))

    # Notes
    if "notes" in results:
        story.append(Paragraph("Notes:", styles['Heading2']))
        note_style = ParagraphStyle(
            'Note',
            parent=styles['Normal'],
            fontSize=10,
            spaceAfter=6
        )
        for key, value in results["notes"].items():
            story.append(Paragraph(f"• {key}: {value}", note_style))

    # Add visualizations for each section
    for section in ("supply", "extract"):
        images = results.get(section, {}).get("images", {})
        if images:
            story.append(Paragraph(f"{section.capitalize()} Visualizations", styles['Heading2']))
            for size, bboxes in images.items():
                story.append(Paragraph(f"Size: {size}", styles['Heading3']))
                for i, bbox in enumerate(bboxes[:3]):  # Limit to 3 images per size
                    img_path = os.path.join(out_img_dir, f"{section}_{size}_{i}.png")
                    try:
                        crop_region_to_image(pdf_path, 0, bbox, img_path)
                        img = RLImage(img_path)
                        # Calculate scaling to fit width while maintaining aspect ratio
                        max_width = 5 * inch  # Leave some margin
                        max_height = 7 * inch  # Leave room for other content
                        img_ratio = float(img.imageHeight) / float(img.imageWidth)
                        
                        # Start with maximum width
                        img.drawWidth = max_width
                        img.drawHeight = img.drawWidth * img_ratio
                        
                        # If height is too large, scale down based on height
                        if img.drawHeight > max_height:
                            img.drawHeight = max_height
                            img.drawWidth = img.drawHeight / img_ratio
                        story.append(img)
                        story.append(Spacer(1, 10))
                    except Exception as e:
                        story.append(Paragraph(f"Error loading image: {str(e)}", styles['Normal']))

    # Build the PDF
    doc.build(story)
    return output_path

    for role in ("supply", "extract"):
        elems.append(Paragraph(f"<b>{role.capitalize()} Ventilation</b>", styles['Heading1']))

        # Table of lengths/counts
        data = [["Type", "Size", "Length (m)"]]
        lengths = results[role]["lengths_mm"]
        for size, length in lengths.items():
            data.append(["Ductwork", size, round(length/1000, 2)])
        for comp, count in results[role]["counts"].items():
            data.append([comp, "", count])

        table = Table(data, hAlign="LEFT")
        table.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.grey)]))
        elems.append(table)
        elems.append(Spacer(1, 12))

        # Images
        for size, bboxes in results[role]["images"].items():
            elems.append(Paragraph(f"Images for {size}", styles['Heading3']))
            for i, bbox in enumerate(bboxes[:3]):
                out_path = os.path.join(out_img_dir, f"{role}_{size}_{i}.png")
                try:
                    crop_region_to_image(source_pdf, 0, bbox, out_path)
                    elems.append(RLImage(out_path, width=200, height=150))
                except:
                    elems.append(Paragraph(f"(Image error for {size})", styles['Normal']))
            elems.append(Spacer(1, 12))

    doc.build(elems)
    return out_pdf_path
