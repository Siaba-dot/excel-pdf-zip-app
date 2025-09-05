import io
import os
import re
import zipfile
import shutil
import tempfile
import calendar
from datetime import datetime
from pathlib import Path

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.pyplot as plt

# ------------------------------
# 📆 Einamo mėn. pabaiga + pavadinimas (LT)
# ------------------------------
def get_current_month_end_and_name():
    today = datetime.today()
    month_end_day = calendar.monthrange(today.year, today.month)[1]
    current_month_end = today.replace(day=month_end_day)
    month_names = [
        "sausio", "vasario", "kovo", "balandžio", "gegužės", "birželio",
        "liepos", "rugpjūčio", "rugsėjo", "spalio", "lapkričio", "gruodžio"
    ]
    current_month_name = month_names[today.month - 1]
    return current_month_end.strftime("%Y-%m-%d"), current_month_name

# ------------------------------
# 🧮 Pirmo lapo konvertavimas į paprastą PDF (be Excel formatavimo)
# ------------------------------
def excel_to_simple_pdf(xlsx_path: str, pdf_path: str):
    """
    Perskaito pirmą skydelį (sheet) į DataFrame ir sukuria paprastą PDF su lentele.
    Tai nėra Excel formatavimo kopija – tik duomenų peržiūrai/spausdinimui.
    """
    try:
        # Nuskaitome pirmą lapą
        df = pd.read_excel(xlsx_path, sheet_name=0, header=None)

        # Sukuriame PDF su vienu puslapiu
        pdf_dir = os.path.dirname(pdf_path)
        os.makedirs(pdf_dir, exist_ok=True)
        with PdfPages(pdf_path) as pdf:
            fig, ax = plt.subplots(figsize=(11.69, 8.27))  # A4 horizontal (apytiksliai)
            ax.axis('off')
            tbl = ax.table(cellText=df.values.astype(str),
                           colLabels=None,
                           loc='center')
            tbl.auto_set_font_size(False)
            tbl.set_fontsize(8)
            tbl.scale(1, 1.2)
            pdf.savefig(fig, bbox_inches='tight')
            plt.close(fig)
        return True, None
    except Exception as e:
        return False, str(e)

# ------------------------------
# 🛠️ Excel apdorojimas visame medyje
# ------------------------------
def process_excels_in_tree(base_dir: str, log_lines: list):
    current_month_end, current_month_name = get_current_month_end_and_name()
    months = [
        "sausio", "vasario", "kovo", "balandžio", "gegužės", "birželio",
        "liepos", "rugpjūčio", "rugsėjo", "spalio", "lapkričio", "gruodžio"
    ]

    all_ok = True

    for root, _, files in os.walk(base_dir):
        rel_root = os.path.relpath(root, base_dir)
        log_lines.append(f"📁 Aplankas: {rel_root if rel_root != '.' else '/'}")
        for filename in files:
            if not filename.lower().endswith(".xlsx"):
                continue

            file_path = os.path.join(root, filename)
            log_lines.append(f"  🔄 Failas: {os.path.join(rel_root, filename)}")

            try:
                # 1) Excel redagavimas (C5 data, A9 mėnesis)
                wb = load_workbook(file_path)
                sheet = wb.active

                # C5 data
                try:
                    if sheet["C5"].value is not None:
                        # įrašome kaip datetime.date, openpyxl pats suformatuos pagal cell numformat
                        dt = datetime.strptime(current_month_end, "%Y-%m-%d").date()
                        sheet["C5"].value = dt
                        log_lines.append(f"    ✅ C5 -> {current_month_end}")
                except Exception as e:
                    log_lines.append(f"    ⚠️ Nepavyko atnaujinti C5: {e}")

                # A9 mėnesio žodis
                try:
                    if sheet["A9"].value:
                        cell_value = str(sheet["A9"].value).strip()
                        # Pakeičiam bet kurį mėnesio žodį į einamą
                        replaced = False
                        for month in months:
                            # case-insensitive paieška
                            if re.search(month, cell_value, flags=re.IGNORECASE):
                                new_val = re.sub(month, current_month_name, cell_value, flags=re.IGNORECASE)
                                sheet["A9"].value = new_val
                                replaced = True
                                log_lines.append(f"    ✅ A9 -> {new_val}")
                                break
                        if not replaced:
                            log_lines.append("    ℹ️ A9: nerastas mėnesio pavadinimas – nepakeista.")
                except Exception as e:
                    log_lines.append(f"    ⚠️ Nepavyko atnaujinti A9: {e}")

                wb.save(file_path)
                wb.close()

                # 2) Pervadinam failą, jei vardas turi YYYY_MM
                try:
                    year = current_month_end[:4]
                    month_num = current_month_end[5:7]
                    new_filename = re.sub(r"(\d{4})_(\d{2})", f"{year}_{month_num}", filename)
                    if new_filename != filename:
                        new_path = os.path.join(root, new_filename)
                        # jeigu toks jau yra, pridėsim sufiksą
                        if os.path.exists(new_path):
                            base, ext = os.path.splitext(new_filename)
                            i = 1
                            while True:
                                candidate = os.path.join(root, f"{base}_v{i}{ext}")
                                if not os.path.exists(candidate):
                                    new_path = candidate
                                    break
                                i += 1
                        os.rename(file_path, new_path)
                        file_path = new_path
                        log_lines.append(f"    📁 Pervadinta -> {os.path.join(rel_root, os.path.basename(file_path))}")
                except Exception as e:
                    log_lines.append(f"    ⚠️ Nepavyko pervadinti: {e}")

                # 3) PDF generavimas (paprastas)
                try:
                    pdf_path = os.path.splitext(file_path)[0] + ".pdf"
                    # jei toks egzistuoja – versijuojam
                    if os.path.exists(pdf_path):
                        base, ext = os.path.splitext(pdf_path)
                        counter = 1
                        while True:
                            candidate = f"{base}_v{counter}{ext}"
                            if not os.path.exists(candidate):
                                pdf_path = candidate
                                break
                            counter += 1

                    ok, err = excel_to_simple_pdf(file_path, pdf_path)
                    if ok:
                        log_lines.append(f"    ✅ PDF -> {os.path.join(rel_root, os.path.basename(pdf_path))}")
                    else:
                        all_ok = False
                        log_lines.append(f"    ❌ PDF klaida: {err}")
                except Exception as e:
                    all_ok = False
                    log_lines.append(f"    ❌ PDF generavimo klaida: {e}")

            except Exception as e:
                all_ok = False
                log_lines.append(f"    ❌ Apdorojimo klaida: {e}")

    return all_ok

# ------------------------------
# 📦 ZIP -> dir ir dir -> ZIP
# ------------------------------
def unzip_to_temp(uploaded_zip_file) -> tuple[str, tempfile.TemporaryDirectory]:
    tmp_dir = tempfile.TemporaryDirectory()
    zip_bytes = uploaded_zip_file.read()
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        zf.extractall(tmp_dir.name)
    return tmp_dir.name, tmp_dir

def zip_tree_to_bytes(root_dir: str) -> bytes:
    mem_zip = io.BytesIO()
    with zipfile.ZipFile(mem_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for folder_name, subfolders, filenames in os.walk(root_dir):
            for fn in filenames:
                abs_path = os.path.join(folder_name, fn)
                arcname = os.path.relpath(abs_path, root_dir)
                zf.write(abs_path, arcname=arcname)
    mem_zip.seek(0)
    return mem_zip.read()

# ------------------------------
# 🖥️ Streamlit UI
# ------------------------------
st.set_page_config(page_title="Aktų apdorojimas (Excel → PDF) | ZIP įkėlimas", page_icon="📄", layout="centered")

st.title("📄 Aktų apdorojimas (Streamlit Cloud)")
st.write(
    "Įkelkite **viso aplanko ZIP** (su poaplankiais). Programa atnaujins Excel failus (C5 datą, A9 mėnesį), "
    "prireikus pervadins failus `YYYY_MM` formatu, sugeneruos paprastus PDF ir grąžins visą medį kaip ZIP."
)

uploaded = st.file_uploader("Įkelkite aplanką kaip .zip", type=["zip"])

if uploaded is not None:
    with st.status("Apdorojama…", expanded=True) as status:
        logs = []
        try:
            base_dir, tmp_handle = unzip_to_temp(uploaded)
            logs.append("📦 ZIP sėkmingai išarchyvuotas.")

            all_ok = process_excels_in_tree(base_dir, logs)

            # Supakuojame atgal į ZIP
            out_bytes = zip_tree_to_bytes(base_dir)
            logs.append("🧷 Paruoštas atsisiunčiamas ZIP su rezultatais.")

            status.update(label="Apdorojimas baigtas.", state="complete")

            # Rodyti log'ą
            st.text("\n".join(logs))

            st.download_button(
                label="⬇️ Parsisiųsti rezultatą (.zip)",
                data=out_bytes,
                file_name=f"apdorota_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip"
            )

            if all_ok:
                st.success("🎉 Visi failai apdoroti sėkmingai!")
            else:
                st.warning("⚠️ Kai kurių failų apdoroti nepavyko. Žr. žurnalą (log).")

        except zipfile.BadZipFile:
            status.update(label="Nepavyko išarchyvuoti ZIP.", state="error")
            st.error("❌ Netinkamas ZIP failas.")
        except Exception as e:
            status.update(label="Įvyko klaida.", state="error")
            st.error(f"❌ Klaida: {e}")
else:
    st.info("👉 Pirmiausia įkelkite **.zip** su savo Excel failais.")

