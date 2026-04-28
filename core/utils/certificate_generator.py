"""
certificate_generator.py
Generates a filled Certificate of Proof Load Test .docx
by cloning the SAMPLECERTE.docx template and replacing placeholder XML.

Place this file at:
    <your_django_app>/utils/certificate_generator.py

The SAMPLECERTE.docx template must be at:
    <your_django_app>/static/templates/SAMPLECERTE.docx
  OR set CERT_TEMPLATE_PATH in Django settings.
"""

import io
import copy
import zipfile
import re
from datetime import date
from lxml import etree


# ── Namespace map used throughout ────────────────────────────────────────────
NS = {
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
}


def _w(tag):
    return '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}' + tag


def _fmt_date(d):
    """Format a date object or string as 'Month DD, YYYY'."""
    if d is None:
        return ''
    if isinstance(d, str):
        return d
    try:
        return d.strftime('%B %d, %Y')
    except Exception:
        return str(d)


def _make_run(text, bold=False, italic=False, underline=False,
              color=None, size=20, font='Times New Roman'):
    """Build a minimal <w:r> element with formatting."""
    nsw = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    r = etree.Element(f'{{{nsw}}}r')
    rPr = etree.SubElement(r, f'{{{nsw}}}rPr')

    fonts = etree.SubElement(rPr, f'{{{nsw}}}rFonts')
    fonts.set(f'{{{nsw}}}ascii', font)
    fonts.set(f'{{{nsw}}}hAnsi', font)
    fonts.set(f'{{{nsw}}}cs', font)

    if bold:
        etree.SubElement(rPr, f'{{{nsw}}}b')
        etree.SubElement(rPr, f'{{{nsw}}}bCs')
    if italic:
        etree.SubElement(rPr, f'{{{nsw}}}i')
        etree.SubElement(rPr, f'{{{nsw}}}iCs')
    if underline:
        u = etree.SubElement(rPr, f'{{{nsw}}}u')
        u.set(f'{{{nsw}}}val', 'single')
    if color:
        c = etree.SubElement(rPr, f'{{{nsw}}}color')
        c.set(f'{{{nsw}}}val', color)

    sz = etree.SubElement(rPr, f'{{{nsw}}}sz')
    sz.set(f'{{{nsw}}}val', str(size))
    szCs = etree.SubElement(rPr, f'{{{nsw}}}szCs')
    szCs.set(f'{{{nsw}}}val', str(size))

    t = etree.SubElement(r, f'{{{nsw}}}t')
    t.text = text
    if text and (text[0] == ' ' or text[-1] == ' '):
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return r


def _replace_paragraph_text(para_el, new_text, bold=False, italic=False,
                              underline=False, color=None, size=20,
                              font='Times New Roman', keep_prefix_runs=False):
    """
    Replace all <w:r> children of a paragraph with a single new run,
    preserving paragraph properties (<w:pPr>).
    If keep_prefix_runs=True, keep any runs that contain drawings/images.
    """
    nsw = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    pPr = para_el.find(f'{{{nsw}}}pPr')

    # Collect runs to keep (those with drawings)
    kept = []
    if keep_prefix_runs:
        for child in list(para_el):
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'r':
                has_drawing = child.find(
                    './/{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}anchor'
                ) is not None or child.find(
                    './/{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline'
                ) is not None
                if has_drawing:
                    kept.append(copy.deepcopy(child))

    # Clear all children except pPr
    for child in list(para_el):
        para_el.remove(child)

    if pPr is not None:
        para_el.append(pPr)

    for k in kept:
        para_el.append(k)

    new_run = _make_run(new_text, bold=bold, italic=italic, underline=underline,
                        color=color, size=size, font=font)
    para_el.append(new_run)


def _get_all_paragraphs(body):
    """Return all <w:p> elements including those inside tables."""
    nsw = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    paras = []
    for el in body.iter():
        if el.tag == f'{{{nsw}}}p':
            paras.append(el)
    return paras


def _para_text(para_el):
    """Extract all text from a paragraph element."""
    nsw = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    parts = []
    for t in para_el.iter(f'{{{nsw}}}t'):
        if t.text:
            parts.append(t.text)
    return ''.join(parts)


def _find_para_containing(body, text_fragment):
    """Find the first paragraph whose text contains text_fragment."""
    for p in _get_all_paragraphs(body):
        if text_fragment in _para_text(p):
            return p
    return None


def _append_run_to_para(para_el, text, **kwargs):
    """Append a new <w:r> run to an existing paragraph."""
    run = _make_run(text, **kwargs)
    para_el.append(run)


def _get_table_data_rows(table_el):
    """Return all <w:tr> elements from a table excluding the header row(s)."""
    nsw = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    rows = table_el.findall(f'{{{nsw}}}tr')
    return rows


def _clear_table_row_cells(row_el):
    """Return list of <w:tc> elements from a table row."""
    nsw = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    return row_el.findall(f'{{{nsw}}}tc')


def _set_cell_text(cell_el, text, size=20, bold=False, font='Times New Roman'):
    """Set the text content of a table cell, preserving cell properties."""
    nsw = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    tcPr = cell_el.find(f'{{{nsw}}}tcPr')

    # Find existing paragraph or create one
    paras = cell_el.findall(f'{{{nsw}}}p')
    if not paras:
        p = etree.SubElement(cell_el, f'{{{nsw}}}p')
    else:
        p = paras[0]
        # Remove extra paragraphs
        for extra in paras[1:]:
            cell_el.remove(extra)

    _replace_paragraph_text(p, text, bold=bold, size=size, font=font)


def generate_certificate(cert, equipment_rows):
    """
    Generate a .docx certificate from the SAMPLECERTE.docx template.

    Parameters
    ----------
    cert : Certificate model instance
        Fields used:
          certificate_number, customer (FK with .name, .address),
          address (override), product_description, capacity,
          model_number, serial_number, working_load_limits,
          tested_load, pressure_relief (optional),
          date_of_inspection, due_next_inspection,
          technician_name, load_cell_cert_text
    equipment_rows : QuerySet of TestEquipmentRow
        Each row: description, capacity, model_no, serial_no,
                  part_no, expiry_date

    Returns
    -------
    io.BytesIO  – ready to stream as HTTP response
    """
    import os
    from django.conf import settings

    # ── Locate the template ───────────────────────────────────────────────
    template_path = getattr(settings, 'CERT_TEMPLATE_PATH', None)
    if not template_path:
        # Default: <app>/static/templates/SAMPLECERTE.docx
        base = settings.BASE_DIR
        template_path = os.path.join(base, 'static', 'templates', 'SAMPLECERTE.docx')

    if not os.path.exists(template_path):
        raise FileNotFoundError(
            f'Certificate template not found at {template_path}. '
            'Copy SAMPLECERTE.docx there or set CERT_TEMPLATE_PATH in settings.py'
        )

    # ── Load template as a zip / XML ─────────────────────────────────────
    with open(template_path, 'rb') as f:
        template_bytes = f.read()

    zin = zipfile.ZipFile(io.BytesIO(template_bytes), 'r')
    doc_xml = zin.read('word/document.xml')
    tree = etree.fromstring(doc_xml)
    body = tree.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body')
    nsw = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    # ── Collect data ──────────────────────────────────────────────────────
    customer_name    = cert.customer.name if cert.customer else ''
    customer_address = getattr(cert, 'address', None) or (cert.customer.address if cert.customer else '')
    cert_number      = cert.certificate_number or ''
    product_desc     = cert.product_description or ''
    capacity         = cert.capacity or ''
    model_number     = cert.model_number or ''
    serial_number    = cert.serial_number or ''
    wll              = getattr(cert, 'working_load_limits', '') or ''
    tested_load      = getattr(cert, 'tested_load', '') or ''
    pressure_relief  = getattr(cert, 'pressure_relief', '') or ''
    date_insp        = _fmt_date(cert.date_of_inspection)
    date_due         = _fmt_date(cert.due_next_inspection)
    technician       = cert.technician_name or ''
    lc_cert_text     = getattr(cert, 'load_cell_cert_text', '') or \
                       'LF 2026-0006 Philippines Geoanalytics Calibration & Measurement Laboratory'

    all_paragraphs = _get_all_paragraphs(body)

    # ── 1. CERTIFICATE NO. line ────────────────────────────────────────────
    # The template has "CERTIFICATE NO. " as a standalone paragraph
    cert_no_para = _find_para_containing(body, 'CERTIFICATE NO.')
    if cert_no_para is not None:
        # Clear existing runs but keep the label+value on same paragraph
        pPr = cert_no_para.find(f'{{{nsw}}}pPr')
        for child in list(cert_no_para):
            cert_no_para.remove(child)
        if pPr is not None:
            cert_no_para.append(pPr)
        # Label
        cert_no_para.append(_make_run('CERTIFICATE NO. ', size=20,
                                       font='Times New Roman'))
        # Value (red, bold)
        cert_no_para.append(_make_run(cert_number, bold=True, size=20,
                                       color='FF0000', font='Times New Roman'))

    # ── 2. CUSTOMER / ADDRESS ─────────────────────────────────────────────
    cust_para = _find_para_containing(body, 'CUSTOMER')
    if cust_para is not None:
        pPr = cust_para.find(f'{{{nsw}}}pPr')
        for child in list(cust_para):
            cust_para.remove(child)
        if pPr is not None:
            cust_para.append(pPr)
        # Bold label + colon + value, then ADDRESS on next line via <w:br/>
        r = _make_run('', bold=True, size=18, font='Times New Roman', color='000000')
        # Inline text with line break
        t_parts = [
            ('CUSTOMER ', True),
            ('\t\t: ', True),
            (customer_name, True),
        ]
        for txt, bld in t_parts:
            cust_para.append(_make_run(txt, bold=bld, size=18,
                                        font='Times New Roman', color='000000'))
        # Line break
        br_r = etree.SubElement(cust_para, f'{{{nsw}}}r')
        etree.SubElement(br_r, f'{{{nsw}}}br')
        # ADDRESS line
        cust_para.append(_make_run('ADDRESS \t\t: ', bold=True, size=18,
                                    font='Times New Roman', color='000000'))
        cust_para.append(_make_run(customer_address, bold=True, size=18,
                                    font='Times New Roman', color='000000'))

    # ── 3. DATE OF INSPECTION ─────────────────────────────────────────────
    doi_para = _find_para_containing(body, 'DATE OF INSPECTION')
    if doi_para is not None:
        pPr = doi_para.find(f'{{{nsw}}}pPr')
        for child in list(doi_para):
            doi_para.remove(child)
        if pPr is not None:
            doi_para.append(pPr)
        doi_para.append(_make_run('DATE OF INSPECTION \t: ', bold=True, size=18,
                                   font='Times New Roman', color='000000'))
        doi_para.append(_make_run(date_insp, bold=True, size=18,
                                   font='Times New Roman', color='000000'))

    # ── 4. DUE NEXT INSPECTION ────────────────────────────────────────────
    dne_para = _find_para_containing(body, 'DUE NEXT INSPECTION')
    if dne_para is not None:
        pPr = dne_para.find(f'{{{nsw}}}pPr')
        for child in list(dne_para):
            dne_para.remove(child)
        if pPr is not None:
            dne_para.append(pPr)
        dne_para.append(_make_run('DUE NEXT INSPECTION  : ', bold=True, size=18,
                                   font='Times New Roman', color='000000'))
        dne_para.append(_make_run(date_due, bold=True, size=18,
                                   font='Times New Roman', color='000000'))

    # ── 5. PRODUCT DESCRIPTION ────────────────────────────────────────────
    prod_para = _find_para_containing(body, 'PRODUCT DESCRIPTION')
    if prod_para is not None:
        pPr = prod_para.find(f'{{{nsw}}}pPr')
        for child in list(prod_para):
            prod_para.remove(child)
        if pPr is not None:
            prod_para.append(pPr)
        prod_para.append(_make_run('PRODUCT DESCRIPTION: ', bold=True, size=18,
                                    font='Times New Roman', color='000000'))
        prod_para.append(_make_run(product_desc, bold=True, size=18,
                                    font='Times New Roman', color='000000'))

    # ── 6. CAPACITY ───────────────────────────────────────────────────────
    cap_para = _find_para_containing(body, 'CAPACITY \t\t:')
    if cap_para is None:
        cap_para = _find_para_containing(body, 'CAPACITY')
        # Make sure we don't match the table header row
        if cap_para is not None and 'CAPACITY' == _para_text(cap_para).strip():
            cap_para = None  # That's the table header, skip it
    if cap_para is not None:
        pPr = cap_para.find(f'{{{nsw}}}pPr')
        for child in list(cap_para):
            cap_para.remove(child)
        if pPr is not None:
            cap_para.append(pPr)
        cap_para.append(_make_run('CAPACITY \t\t: ', bold=True, size=18,
                                   font='Times New Roman', color='000000'))
        cap_para.append(_make_run(capacity, bold=True, size=18,
                                   font='Times New Roman', color='000000'))

    # ── 7. MODEL NUMBER ───────────────────────────────────────────────────
    mod_para = _find_para_containing(body, 'MODEL NUMBER')
    if mod_para is not None:
        pPr = mod_para.find(f'{{{nsw}}}pPr')
        for child in list(mod_para):
            mod_para.remove(child)
        if pPr is not None:
            mod_para.append(pPr)
        mod_para.append(_make_run('MODEL NUMBER \t: ', bold=True, size=18,
                                   font='Times New Roman', color='000000'))
        mod_para.append(_make_run(model_number, bold=True, size=18,
                                   font='Times New Roman', color='000000'))

    # ── 8. SERIAL NUMBER ──────────────────────────────────────────────────
    ser_para = _find_para_containing(body, 'SERIAL NUMBER')
    if ser_para is not None:
        pPr = ser_para.find(f'{{{nsw}}}pPr')
        for child in list(ser_para):
            ser_para.remove(child)
        if pPr is not None:
            ser_para.append(pPr)
        ser_para.append(_make_run('SERIAL NUMBER \t: ', bold=True, size=18,
                                   font='Times New Roman', color='000000'))
        ser_para.append(_make_run(serial_number, bold=True, size=18,
                                   font='Times New Roman', color='000000'))

    # ── 9. TEST RESULTS table (WLL / Tested Load / Pressure Relief) ───────
    # Find the table that has "TEST RESULTS" header
    tables = body.findall(f'{{{nsw}}}tbl')
    test_results_table = None
    test_equipment_table = None
    load_cell_table = None

    for tbl in tables:
        tbl_text = ''.join(
            t.text or '' for t in tbl.iter(f'{{{nsw}}}t')
        )
        if 'TEST RESULTS' in tbl_text and 'WORKING LOAD LIMITS' in tbl_text:
            test_results_table = tbl
        elif 'TEST EQUIPMENT' in tbl_text and 'DESCRIPTION' in tbl_text:
            test_equipment_table = tbl
        elif 'Load Cell Certificate' in tbl_text:
            load_cell_table = tbl

    if test_results_table is not None:
        rows = test_results_table.findall(f'{{{nsw}}}tr')
        # Row 0: "TEST RESULTS" header (span)
        # Row 1: column headers: WORKING LOAD LIMITS | TESTED LOAD | PRESSURE RELIEF
        # Row 2: data row
        if len(rows) >= 3:
            data_row = rows[2]
            cells = data_row.findall(f'{{{nsw}}}tc')
            if len(cells) >= 1:
                _set_cell_text(cells[0], wll, size=20)
            if len(cells) >= 2:
                _set_cell_text(cells[1], tested_load, size=20)
            if len(cells) >= 3:
                _set_cell_text(cells[2], pressure_relief, size=20)

    # ── 10. TEST EQUIPMENT table ──────────────────────────────────────────
    if test_equipment_table is not None:
        rows = test_equipment_table.findall(f'{{{nsw}}}tr')
        # Row 0: "TEST EQUIPMENT" header
        # Row 1: column headers: DESCRIPTION | CAPACITY | MODEL NO. | SERIAL NO. | PART NO. | EXPIRY DATE
        # Row 2+: sample/data rows — we overwrite row 2 and add/remove as needed

        eq_list = list(equipment_rows)

        # Keep header rows (0, 1), replace/add data rows from index 2 onward
        header_rows = rows[:2]
        existing_data_rows = rows[2:]

        # Remove all existing data rows
        for old_row in existing_data_rows:
            test_equipment_table.remove(old_row)

        if not eq_list:
            # Add one blank row so the table isn't empty
            if existing_data_rows:
                blank = copy.deepcopy(existing_data_rows[0])
                # Clear all cell texts
                for tc in blank.findall(f'{{{nsw}}}tc'):
                    _set_cell_text(tc, '', size=20)
                test_equipment_table.append(blank)
        else:
            template_row = existing_data_rows[0] if existing_data_rows else None
            for eq in eq_list:
                if template_row is not None:
                    new_row = copy.deepcopy(template_row)
                else:
                    # Build a row from scratch using header row structure
                    new_row = copy.deepcopy(header_rows[1]) if len(header_rows) > 1 else etree.Element(f'{{{nsw}}}tr')

                cells = new_row.findall(f'{{{nsw}}}tc')
                values = [
                    str(getattr(eq, 'description', '') or ''),
                    str(getattr(eq, 'capacity', '') or ''),
                    str(getattr(eq, 'model_no', '') or getattr(eq, 'model_number', '') or ''),
                    str(getattr(eq, 'serial_no', '') or getattr(eq, 'serial_number', '') or ''),
                    str(getattr(eq, 'part_no', '') or getattr(eq, 'part_number', '') or ''),
                    _fmt_date(getattr(eq, 'expiry_date', '') or ''),
                ]
                for i, tc in enumerate(cells):
                    if i < len(values):
                        _set_cell_text(tc, values[i], size=20)
                test_equipment_table.append(new_row)

    # ── 11. Load Cell Certificate table ───────────────────────────────────
    if load_cell_table is not None:
        rows = load_cell_table.findall(f'{{{nsw}}}tr')
        if rows:
            cells = rows[0].findall(f'{{{nsw}}}tc')
            if len(cells) >= 2:
                _set_cell_text(cells[1], lc_cert_text, size=20)

    # ── 12. Technician name (underlined) ──────────────────────────────────
    tech_para = _find_para_containing(body, 'Ariel Lacsina')
    if tech_para is None:
        # Search more broadly
        for p in _get_all_paragraphs(body):
            txt = _para_text(p)
            if technician and technician in txt:
                tech_para = p
                break
    if tech_para is not None:
        pPr = tech_para.find(f'{{{nsw}}}pPr')
        for child in list(tech_para):
            tech_para.remove(child)
        if pPr is not None:
            tech_para.append(pPr)
        # Technician name underlined
        tech_para.append(_make_run(technician, underline=True, size=20,
                                    font='Times New Roman', color='000000'))
        # Spacer tabs
        spacer_r = etree.SubElement(tech_para, f'{{{nsw}}}r')
        for _ in range(9):
            etree.SubElement(spacer_r, f'{{{nsw}}}tab')

    # ── Serialize modified XML back ───────────────────────────────────────
    new_doc_xml = etree.tostring(tree, xml_declaration=True,
                                  encoding='UTF-8', standalone=True)

    # ── Build output zip (clone template, replace document.xml) ───────────
    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == 'word/document.xml':
                zout.writestr(item, new_doc_xml)
            else:
                zout.writestr(item, zin.read(item.filename))

    zin.close()
    out_buf.seek(0)
    return out_buf