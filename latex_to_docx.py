import re
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

DRIVE_DIR = Path(
    r"G:\Mi unidad\Borrar - Dualismo filosófico - Eneko\Espiritualismo\Estudios\Euskal Herriko Unibertsitatea\Compartido Eneko\4. Kurtsoa\4. Kurtsoa 1. Lauhilekoa\RKA\rka_6\2. Proiektua"
)
BASE_DIR = Path(__file__).resolve().parent

OUTPUT_DOCX_PARALLEL = DRIVE_DIR / "Praktika_Txostena_RKA_berrikusia.docx"
OUTPUT_DOCX_MAIN = DRIVE_DIR / "Praktika_Txostena_RKA.docx"
LOCAL_DOCX_PARALLEL = BASE_DIR / "Praktika_Txostena_RKA_berrikusia.docx"
LOCAL_DOCX_MAIN = BASE_DIR / "Praktika_Txostena_RKA.docx"

REFS = {
    "fig:ibilbideak": "1",
    "fig:traj-p1": "1a",
    "fig:traj-p2": "1b",
    "fig:traj-p3": "1c",
    "fig:kamera-segmentazioa": "2",
    "fig:sim-zirkuitua": "3",
    "fig:abiadurak-telemetria": "4",
    "tab:rwander-10min": "1",
    "tab:rwander-5min": "1",
    "tab:sim-emaitzak": "2",
    "tab:road-5saiakera": "3",
    "tab:rwander-extra": "4",
    "sec:oz-sektoreak": "2.1",
    "sec:oz-mugimendua": "2.2",
    "sec:oz-esplorazioa": "2.3",
    "sec:oz-emaitzak": "2.4",
    "sec:oz-arriskuak": "2.5",
    "sec:er-eremuak": "3.1",
    "sec:er-abiadurak": "3.2",
    "sec:er-simulazioa": "3.3",
    "sec:er-doikuntzak": "3.4",
    "sec:er-emaitzak": "3.5",
    "sec:eranskina-probak": "5.1",
    "sec:eranskina-materialak": "5.2",
}

CITES = {
    "ros2": "[1]",
    "lazkano2026": "[2]",
    "gazebo": "[3]",
}


def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))


def clean_latex(text: str) -> str:
    t = text
    # Replace references
    def replace_ref(m):
        key = m.group(1).strip()
        return REFS.get(key, "1")
    t = re.sub(r'\\ref\{([^}]+)\}', replace_ref, t)
    
    # Replace citations
    def replace_cite(m):
        key = m.group(1).strip()
        return CITES.get(key, "[1]")
    t = re.sub(r'\\cite\{([^}]+)\}', replace_cite, t)

    # Basic formatting
    t = re.sub(r'\\textbf\{([^}]+)\}', r'\1', t)
    t = re.sub(r'\\texttt\{([^}]+)\}', r'\1', t)
    t = re.sub(r'\\emph\{([^}]+)\}', r'\1', t)
    t = re.sub(r'\\mathbf\{([^}]+)\}', r'\1', t)
    t = re.sub(r'\\text\{([^}]+)\}', r'\1', t)
    t = re.sub(r'\\mathcal\{U\}', 'U', t)
    t = re.sub(r'\\label\{[^}]+\}', '', t)
    t = re.sub(r'\\setlength\{[^}]+\}\{[^}]+\}', '', t)
    t = re.sub(r'\\tabcolsep', '', t)
    t = re.sub(r'\\bar\{t\}', 't̄', t)
    
    # Fractions: \frac{a}{b} -> (a / b)
    t = re.sub(r'\\frac\{([^}]+)\}\{([^}]+)\}', r'(\1 / \2)', t)
    
    # Spacing and escaping
    t = t.replace('~', ' ')
    t = t.replace(r'\,', ' ')
    t = t.replace(r'\_', '_')
    t = t.replace(r'\%', '%')
    t = t.replace(r'\dots', '...')
    t = t.replace('{,}', ',')
    
    # Math symbols
    t = t.replace(r'\pm', '±')
    t = t.replace(r'\le', '≤')
    t = t.replace(r'\ge', '≥')
    t = t.replace(r'\approx', '≈')
    t = t.replace(r'\pi', 'π')
    t = t.replace(r'\theta', 'θ')
    t = t.replace(r'\omega', 'ω')
    t = t.replace(r'\cdot', '·')
    t = t.replace(r'\times', '×')
    t = t.replace(r'^\circ', '°')
    t = t.replace(r'\circ', '°')
    t = re.sub(r'\$([^$]+)\$', r'\1', t)
    
    # Clean up double spaces or residual backslashes
    t = re.sub(r' +', ' ', t)
    return t.strip()


def add_code_block(doc, code_text):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, "F4F5F7")
    cell.width = Inches(6.5)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    run = p.add_run(code_text.strip())
    run.font.name = "Consolas"
    run.font.size = Pt(9.5)
    run.font.color.rgb = RGBColor(30, 41, 59)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def parse_latex_table(doc, table_text: str):
    caption = ""
    m_cap = re.search(r'\\caption(?:\[[^\]]*\])?\{([^}]+)\}', table_text)
    if m_cap:
        caption = clean_latex(m_cap.group(1))

    # Extract rows between \midrule (or \toprule) and \bottomrule
    rows = []
    for line in table_text.splitlines():
        line = line.strip()
        if not line or any(cmd in line for cmd in ("\\begin", "\\end", "\\toprule", "\\midrule", "\\bottomrule", "\\caption", "\\centering", "\\label", "\\setlength")):
            continue
        if "\\multicolumn" in line:
            m_txt = re.search(r'\{@\{\}l\}\{([^}]+)\}', line)
            if m_txt:
                rows.append([clean_latex(m_txt.group(1))])
            continue
        if "\\\\" in line:
            line = line.replace("\\\\", "")
        parts = [clean_latex(c) for c in line.split("&")]
        if any(p.strip() for p in parts):
            rows.append(parts)

    if not rows:
        return

    num_cols = max(len(r) for r in rows)
    tbl = doc.add_table(rows=len(rows), cols=num_cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    for r_idx, r_data in enumerate(rows):
        row_obj = tbl.rows[r_idx]
        trPr = row_obj._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

        is_header = (r_idx == 0)
        is_colspan = (len(r_data) == 1 and num_cols > 1)
        
        if is_colspan:
            cell = tbl.cell(r_idx, 0)
            cell.text = r_data[0]
            for c_idx in range(1, num_cols):
                cell.merge(tbl.cell(r_idx, c_idx))
            cell.paragraphs[0].runs[0].bold = True
            set_cell_background(cell, "F8FAFC")
            continue

        for c_idx, val in enumerate(r_data):
            cell = tbl.cell(r_idx, c_idx)
            cell.text = val
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            if is_header:
                set_cell_background(cell, "E2E8F0")
                if p.runs:
                    p.runs[0].bold = True
            else:
                set_cell_background(cell, "FFFFFF" if r_idx % 2 == 1 else "F8FAFC")

    if caption:
        p_cap = doc.add_paragraph()
        r = p_cap.add_run(f"Taula: {caption}")
        r.font.size = Pt(9.5)
        r.italic = True
        r.font.color.rgb = RGBColor(100, 116, 139)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def parse_figure_block(doc, fig_text: str):
    subfigs = re.findall(
        r'\\begin\{subfigure\}(?:\[[^\]]*\])?\{[^}]*\}\s*\\includegraphics(?:\[[^\]]*\])?\{([^}]+)\}\s*\\caption\{([^}]+)\}',
        fig_text
    )
    if subfigs:
        tbl = doc.add_table(rows=1, cols=len(subfigs))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        for col_idx, (img_name, subcap) in enumerate(subfigs):
            cell = tbl.cell(0, col_idx)
            img_path = BASE_DIR / "irudiak" / img_name.strip()
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if img_path.exists():
                p.add_run().add_picture(str(img_path), width=Inches(1.95))
            p_cap = cell.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_c = p_cap.add_run(clean_latex(subcap))
            r_c.font.size = Pt(8.5)
            r_c.italic = True
        
        # Remove borders
        tblPr = tbl._tbl.tblPr
        tblBorders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>'
            f'<w:insideH w:val="none"/><w:insideV w:val="none"/>'
            f'</w:tblBorders>'
        )
        tblPr.append(tblBorders)
    else:
        m_img = re.search(r'\\includegraphics(?:\[[^\]]*\])?\{([^}]+)\}', fig_text)
        if m_img:
            img_name = m_img.group(1).strip()
            img_path = BASE_DIR / "irudiak" / img_name
            if img_path.exists():
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(6)
                p_img.paragraph_format.space_after = Pt(4)
                width = Inches(5.5) if "kamera" in img_name or "abiadurak" in img_name else Inches(4.5)
                p_img.add_run().add_picture(str(img_path), width=width)

    m_cap = re.search(r'\\caption(?:\[[^\]]*\])?\{([^}]+)\}', fig_text)
    if m_cap:
        cap_text = clean_latex(m_cap.group(1))
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(8)
        r = p_cap.add_run(f"Irudia: {cap_text}")
        r.font.size = Pt(9.5)
        r.italic = True
        r.font.color.rgb = RGBColor(71, 85, 105)


def parse_chapter_content(doc, content: str):
    lines = content.splitlines()
    in_code = False
    code_lines = []
    in_table = False
    table_lines = []
    in_figure = False
    figure_lines = []
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        
        # Comments
        if line.startswith("%"):
            i += 1
            continue

        # Code block
        if "\\begin{lstlisting}" in line:
            in_code = True
            code_lines = []
            i += 1
            continue
        elif "\\end{lstlisting}" in line:
            in_code = False
            add_code_block(doc, "\n".join(code_lines))
            i += 1
            continue
        elif in_code:
            code_lines.append(lines[i])
            i += 1
            continue

        # Table block
        if "\\begin{table}" in line:
            in_table = True
            table_lines = [line]
            i += 1
            while i < len(lines) and "\\end{table}" not in lines[i]:
                table_lines.append(lines[i])
                i += 1
            if i < len(lines):
                table_lines.append(lines[i])
                i += 1
            parse_latex_table(doc, "\n".join(table_lines))
            in_table = False
            continue
        elif "\\begin{tabular" in line:
            in_table = True
            table_lines = [line]
            i += 1
            while i < len(lines) and "\\end{tabular" not in lines[i]:
                table_lines.append(lines[i])
                i += 1
            if i < len(lines):
                table_lines.append(lines[i])
                i += 1
            parse_latex_table(doc, "\n".join(table_lines))
            in_table = False
            continue

        # Figure block
        if "\\begin{figure}" in line:
            in_figure = True
            figure_lines = [line]
            i += 1
            while i < len(lines) and "\\end{figure}" not in lines[i]:
                figure_lines.append(lines[i])
                i += 1
            if i < len(lines):
                figure_lines.append(lines[i])
                i += 1
            parse_figure_block(doc, "\n".join(figure_lines))
            in_figure = False
            continue

        # Headings
        m_sec = re.match(r'\\section\{([^}]+)\}', line)
        if m_sec:
            doc.add_heading(clean_latex(m_sec.group(1)), level=2)
            i += 1
            continue

        m_ssec = re.match(r'\\subsection\{([^}]+)\}', line)
        if m_ssec:
            doc.add_heading(clean_latex(m_ssec.group(1)), level=3)
            i += 1
            continue

        m_para = re.match(r'\\paragraph\{([^}]+)\}', line)
        if m_para:
            p = doc.add_paragraph()
            r = p.add_run(clean_latex(m_para.group(1)) + " ")
            r.bold = True
            rest = line[m_para.end():].strip()
            if rest:
                p.add_run(clean_latex(rest))
            i += 1
            continue

        # Enumerate / Itemize
        if line.startswith("\\item"):
            item_text = line[len("\\item"):].strip()
            m_bold = re.match(r'\\textbf\{([^}]+)\}', item_text)
            if m_bold:
                prefix = m_bold.group(1)
                body = item_text[m_bold.end():].strip()
                p = doc.add_paragraph(style='List Bullet')
                r = p.add_run(clean_latex(prefix) + " ")
                r.bold = True
                if body:
                    p.add_run(clean_latex(body))
            else:
                p = doc.add_paragraph(clean_latex(item_text), style='List Bullet')
            i += 1
            continue

        # Equations
        if line.startswith("\\begin{equation}"):
            eq_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("\\end{equation}"):
                eq_lines.append(lines[i].strip())
                i += 1
            i += 1
            eq_text = clean_latex(" ".join(eq_lines))
            p_eq = doc.add_paragraph()
            p_eq.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_eq.paragraph_format.space_before = Pt(6)
            p_eq.paragraph_format.space_after = Pt(6)
            r_eq = p_eq.add_run(eq_text)
            r_eq.italic = True
            r_eq.font.size = Pt(11)
            continue

        # Bold question prompts (from peer review / guide)
        m_q = re.match(r'\\textbf\{([^}]+)\}', line)
        if m_q and line.endswith("}"):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
            r = p.add_run(clean_latex(m_q.group(1)))
            r.bold = True
            r.font.color.rgb = RGBColor(15, 23, 42)
            i += 1
            continue

        if line in (
            "\\begin{itemize}", "\\end{itemize}",
            "\\begin{enumerate}", "\\end{enumerate}",
            "\\begin{description}", "\\end{description}",
            "\\begin{figure}", "\\end{figure}",
            "\\centering", "\\clearpage", "\\hfill"
        ):
            i += 1
            continue

        # Regular paragraph
        if line:
            p_lines = [line]
            i += 1
            while i < len(lines):
                next_l = lines[i].strip()
                if not next_l or next_l.startswith("\\"):
                    break
                p_lines.append(next_l)
                i += 1
            doc.add_paragraph(clean_latex(" ".join(p_lines)))
        else:
            i += 1


def build_document():
    doc = Document()
    
    # Configure margins
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(12)
    p_title.paragraph_format.space_after = Pt(4)
    r_title = p_title.add_run("Oztopoak ekiditen eta errepideari jarraitzen")
    r_title.bold = True
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(22)
    r_title.font.color.rgb = RGBColor(15, 23, 42)

    # Subtitle
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run("Sentsoreen funtzionamendua lantzen: ROSMASTER X3 robotarekin egindako praktika")
    r_sub.italic = True
    r_sub.font.size = Pt(12)
    r_sub.font.color.rgb = RGBColor(71, 85, 105)

    # Authors & Info box
    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_meta.paragraph_format.space_after = Pt(24)
    r_meta = p_meta.add_run(
        "Egileak: Eneko Ruiz, Eneko Pardavila, Omar Ouinas\n"
        "Informatika Ingeniaritzako Gradua · Robotika eta Kontrol Adimendunak (Irakaslea: Elena Lazkano)\n"
        "Euskal Herriko Unibertsitatea (UPV/EHU)"
    )
    r_meta.font.size = Pt(10)
    r_meta.font.color.rgb = RGBColor(100, 116, 139)

    doc.add_heading("Laburpena", level=1)
    laburpena_path = BASE_DIR / "kapituluak" / "laburpena.tex"
    if laburpena_path.exists():
        text = laburpena_path.read_text(encoding="utf-8")
        paras = [p.strip() for p in text.split("\n\n") if p.strip() and not p.strip().startswith("%")]
        for p in paras:
            doc.add_paragraph(clean_latex(p))

    chapters = [
        ("1. Sarrera", "01_sarrera.tex"),
        ("2. Oztopoak ekiditen (LiDAR)", "02_oztopoak.tex"),
        ("3. Errepideari jarraitzen (Kamera)", "03_errepidea.tex"),
        ("4. Ondorioak", "04_ondorioak.tex"),
        ("5. Eranskina: Proba osagarriak eta biltegia", "05_eranskina.tex"),
    ]

    for ch_title, filename in chapters:
        f_path = BASE_DIR / "kapituluak" / filename
        if not f_path.exists():
            continue
        
        doc.add_heading(ch_title, level=1)
        content = f_path.read_text(encoding="utf-8")
        parse_chapter_content(doc, content)

    # Save to both local destinations
    for loc_p in [LOCAL_DOCX_PARALLEL, LOCAL_DOCX_MAIN]:
        if loc_p.exists():
            try:
                loc_p.unlink()
            except Exception:
                pass
        doc.save(str(loc_p))
        print(f"[OK] Guardado localmente: {loc_p.name}")
    
    # Save to Google Drive
    if DRIVE_DIR.exists():
        for drive_p in [OUTPUT_DOCX_PARALLEL, OUTPUT_DOCX_MAIN]:
            try:
                if drive_p.exists():
                    drive_p.unlink()
            except Exception:
                pass
            doc.save(str(drive_p))
            print(f"[OK] Guardado en Google Drive: {drive_p.name}")
    else:
        print(f"[ALERTA] No se encontro la carpeta de Drive: {DRIVE_DIR}")


if __name__ == "__main__":
    build_document()
