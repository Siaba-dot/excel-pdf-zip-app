def excel_to_pdf_reportlab(xlsx_path: str, pdf_path: str):
    """
    Generuoja PDF iš pirmo Excel lapo, kiek įmanoma išlaikant išdėstymą:
    - stulpelių plotis / eilučių aukštis (matuojami ir suskalaujami prie A4 su paraštėmis),
    - rėmeliai, lygiuotės, bold/italic,
    - merged cells.
    Pastaba: tai nėra 1:1 Excel renderis, bet žymiai artimesnis už 'matplotlib' lentelę.
    """
    try:
        wb = load_workbook(xlsx_path, data_only=True)
        ws: Worksheet = wb.active

        # A4 kraštai ir darbinis plotas
        page_w, page_h = A4  # pts
        margin = 12 * mm
        content_w = page_w - 2 * margin
        content_h = page_h - 2 * margin

        # Naudojamas diapazonas (arba print_area)
        if ws.print_area:
            # print_area pvz.: 'A1:F40'
            area = str(ws.print_area)
        else:
            area = ws.calculate_dimension()  # pvz. 'A1:F40' pagal used range

        min_col, min_row, max_col, max_row = ws.calculate_dimension().split(':')[0], None, None, None
        # Patikslinam normaliai:
        min_col_idx, min_row_idx, max_col_idx, max_row_idx = ws.calculate_dimension().bounds

        # Sudarome stulpelių pločius (Excel vienetus -> pts, vėliau skaluosim)
        col_widths = []
        for c in range(min_col_idx, max_col_idx + 1):
            letter = get_column_letter(c)
            cw = ws.column_dimensions[letter].width
            if cw is None:
                cw = 8.43  # Excel default
            # Excel width (approx chars) -> pixels (~7 px/char) -> points (72/96)
            pts = cw * 7 * (72.0 / 96.0)
            col_widths.append(pts)

        row_heights = []
        for r in range(min_row_idx, max_row_idx + 1):
            rh = ws.row_dimensions[r].height
            if rh is None:
                rh = 15  # Excel default row height in points (approx)
            else:
                # openpyxl row height jau būna pts
                pass
            row_heights.append(float(rh))

        # Suskaičiuojam bendrą dydį ir skaluojam, kad tilptų į A4 (išlaikant proporcijas)
        total_w_pts = sum(col_widths)
        total_h_pts = sum(row_heights)

        scale_x = content_w / total_w_pts if total_w_pts > 0 else 1.0
        scale_y = content_h / total_h_pts if total_h_pts > 0 else 1.0
        scale = min(scale_x, scale_y, 1.0)  # nemažinam paraštėmis, bet neviršijam lapo

        # Koord. pradžia (kairė-apačia)
        origin_x = margin + (content_w - total_w_pts * scale) / 2.0
        origin_y = margin + (content_h - total_h_pts * scale) / 2.0

        # Paruošiam drobę
        os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
        c = canvas.Canvas(pdf_path, pagesize=A4)

        # Paruošiam merged ranges lookup
        merged_rects = {}  # (row, col) -> (rowspan, colspan)
        for m in ws.merged_cells.ranges:
            minr, minc, maxr, maxc = m.min_row, m.min_col, m.max_row, m.max_col
            merged_rects[(minr, minc)] = (maxr - minr + 1, maxc - minc + 1)

        # Kad nežymėti vidinių merged langelių (tik viršutinį-kairį piešiam)
        merged_members = set()
        for (sr, sc), (rs, cs) in merged_rects.items():
            for rr in range(sr, sr + rs):
                for cc in range(sc, sc + cs):
                    if not (rr == sr and cc == sc):
                        merged_members.add((rr, cc))

        # Pagalbinės sumos → koordinatėms
        col_acc = [0]
        for w in col_widths:
            col_acc.append(col_acc[-1] + w * scale)
        row_acc = [0]
        for h in row_heights:
            row_acc.append(row_acc[-1] + h * scale)

        # Piešiam langelius nuo viršaus į apačią (Reportlab koord. apačioj, tad verčiam)
        def cell_xywh(r, c, rowspan=1, colspan=1):
            # r,c yra 1-indeksuoti Excel koordinatės diapazone
            r0 = r - min_row_idx
            c0 = c - min_col_idx
            x = origin_x + col_acc[c0]
            y = origin_y + (total_h_pts * scale - row_acc[r0 + rowspan] + row_acc[r0])
            w = col_acc[c0 + colspan] - col_acc[c0]
            h = row_acc[r0 + rowspan] - row_acc[r0]
            return x, y, w, h

        # Pirmiau nupiešiam rėmelius ir užpildus, tada tekstą
        for r in range(min_row_idx, max_row_idx + 1):
            for c_idx in range(min_col_idx, max_col_idx + 1):
                if (r, c_idx) in merged_members:
                    continue
                rowspan, colspan = 1, 1
                if (r, c_idx) in merged_rects:
                    rowspan, colspan = merged_rects[(r, c_idx)]
                x, y, w, h = cell_xywh(r, c_idx, rowspan, colspan)

                cell = ws.cell(row=r, column=c_idx)
                # Užpildymas (jei yra)
                fill = cell.fill
                if fill and fill.start_color and getattr(fill.start_color, "rgb", None) and fill.start_color.rgb not in (None, "00000000", "000000"):
                    try:
                        rgb = fill.start_color.rgb  # 'FFRRGGBB'
                        if rgb and len(rgb) == 8:
                            rr = int(rgb[2:4], 16) / 255.0
                            gg = int(rgb[4:6], 16) / 255.0
                            bb = int(rgb[6:8], 16) / 255.0
                            c.setFillColor(colors.Color(rr, gg, bb))
                            c.rect(x, y, w, h, fill=1, stroke=0)
                    except Exception:
                        pass

                # Rėmeliai (supaprastintai: jei bet koks border -> plona linija)
                border = cell.border
                draw_border = any([
                    border.left and border.left.style,
                    border.right and border.right.style,
                    border.top and border.top.style,
                    border.bottom and border.bottom.style
                ])
                if draw_border:
                    c.setStrokeColor(colors.black)
                    c.setLineWidth(0.6)
                    c.rect(x, y, w, h, fill=0, stroke=1)
                else:
                    # plona tinklinė linija, jei norite: praleidžiam, kad nebūtų per daug „grid“
                    pass

        # Tekstas (su lygiuotėmis ir šriftais)
        for r in range(min_row_idx, max_row_idx + 1):
            for c_idx in range(min_col_idx, max_col_idx + 1):
                if (r, c_idx) in merged_members:
                    continue
                rowspan, colspan = 1, 1
                if (r, c_idx) in merged_rects:
                    rowspan, colspan = merged_rects[(r, c_idx)]
                x, y, w, h = cell_xywh(r, c_idx, rowspan, colspan)
                cell = ws.cell(row=r, column=c_idx)
                val = "" if cell.value is None else str(cell.value)

                # Stiliai
                font = cell.font
                bold = bool(font and font.bold)
                italic = bool(font and font.italic)
                font_name = "Helvetica-Bold" if bold else "Helvetica"
                if italic and bold:
                    font_name = "Helvetica-BoldOblique"
                elif italic and not bold:
                    font_name = "Helvetica-Oblique"

                font_size = 9  # default
                if font and font.sz:
                    try:
                        font_size = float(font.sz)
                    except Exception:
                        pass

                # Lygiuotės
                ha = "left"
                va = "middle"
                if cell.alignment:
                    if cell.alignment.horizontal in ("center", "centerContinuous", "distributed", "justify"):
                        ha = "center"
                    elif cell.alignment.horizontal in ("right",):
                        ha = "right"
                    if cell.alignment.vertical in ("top", "distributed", "justify"):
                        va = "top"
                    elif cell.alignment.vertical in ("bottom",):
                        va = "bottom"

                # Paraštėlės tekste
                pad_x = 2  # pt
                pad_y = 1  # pt

                # Teksto laukelio koordinatė (Reportlab origin bottom-left)
                tx = x + pad_x
                ty = y + pad_y

                c.setFont(font_name, font_size)
                c.setFillColor(colors.black)

                # Horizontalus pozicionavimas
                if ha == "left":
                    text_x = tx
                elif ha == "center":
                    text_x = x + w / 2.0
                else:  # right
                    text_x = x + w - pad_x

                # Vertikalus pozicionavimas
                # Naudojam baseline ~ ty + ...
                if va == "top":
                    text_y = y + h - pad_y - font_size
                elif va == "bottom":
                    text_y = y + pad_y
                else:
                    text_y = y + (h - font_size) / 2.0  # apytiksliai middle

                if ha == "center":
                    c.drawCentredString(text_x, text_y, val)
                elif ha == "right":
                    c.drawRightString(text_x, text_y, val)
                else:
                    c.drawString(text_x, text_y, val)

        c.showPage()
        c.save()
        wb.close()
        return True, None
    except Exception as e:
        return False, str(e)

