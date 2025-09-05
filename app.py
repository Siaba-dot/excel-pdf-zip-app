import io
import os
import re
import zipfile
import tempfile
import calendar
from datetime import datetime

import streamlit as st
from openpyxl import load_workbook


# ------------------------------
# Pagalbinės funkcijos
# ------------------------------
def get_current_month_end_and_name():
    """Einamo mėnesio paskutinė diena ir pavadinimas (lietuviškai)."""
    today = datetime.today()
    month_end_day = calendar.monthrange(today.year, today.month)[1]
    current_month_end = today.replace(day=month_end_day)
    month_names = [
        "sausio", "vasario", "kovo", "balandžio", "gegužės", "birželio",
        "liepos", "rugpjūčio", "rugsėjo", "spalio", "lapkričio", "gruodžio"
    ]
    return current_month_end.strftime("%Y-%m-%d"), month_names[today.month - 1]


def unzip_to_temp(uploaded_zip_file):
    """Išarchyvuoja ZIP į laikiną aplanką."""
    tmp_dir = tempfile.TemporaryDirectory()
    zip_bytes = uploaded_zip_file.read()
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        zf.extractall(tmp_dir.name)
    return tmp_dir.name, tmp_dir


def zip_tree_to_bytes(root_dir: str) -> bytes:
    """Supakuoja aplanką į ZIP baitus."""
    mem_zip = io.BytesIO()
    with zipfile.ZipFile(mem_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for folder_name, _, filenames in os.walk(root_dir):
            for fn in filenames:
                abs_path = os.path.join(folder_name, fn)
                arcname = os.path.relpath(abs_path, root_dir)
                zf.write(abs_path, arcname=arcname)
    mem_zip.seek(0)
    return mem_zip.read()


def process_excels_in_tree(base_dir: str, log_lines: list):
    """Atnaujina visus Excel failus medyje."""
    current_month_end, current_month_name = get_current_month_end_and_name()
    months = [
        "sausio", "vasario", "kovo", "balandžio", "gegužės", "birželio",
        "liepos", "rugpjūčio", "rugsėjo", "spalio", "lapkričio", "gruodžio"
    ]

    for root, _, files in os.walk(base_dir):
        rel_root = os.path.relpath(root, base_dir)
        log_lines.append(f"📁 Aplankas: {rel_root if rel_root != '.' else '/'}")
        for filename in files:
            if not filename.lower().endswith((".xlsx", ".xlsm")):
                continue

            file_path = os.path.join(root, filename)
            log_lines.append(f"  🔄 Failas: {os.path.join(rel_root, filename)}")

            try:
                wb = load_workbook(file_path)
                sheet = wb.active

                # C5 – mėn. pabaigos data
                if sheet["C5"].value is not None:
                    dt = datetime.strptime(current_month_end, "%Y-%m-%d").date()
                    sheet["C5"].value = dt
                    log_lines.append(f"    ✅ C5 -> {current_month_end}")

                # A9 – mėnesio pavadinimo keitimas
                if sheet["A9"].value is not None:
                    cell_value = str(sheet["A9"].value).strip()
                    replaced = False
                    for month in months:
                        if re.search(month, cell_value, flags=re.IGNORECASE):
                            new_val = re.sub(month, current_month_name, cell_value, flags=re.IGNORECASE)
                            sheet["A9"].value = new_val
                            replaced = True
                            log_lines.append(f"    ✅ A9 -> {new_val}")
                            break
                    if not replaced:
                        log_lines.append("    ℹ️ A9: nerastas mėnesio pavadinimas – nepakeista.")

                wb.save(file_path)
                wb.close()

                # Pervadinimas pagal YYYY_MM
                year = current_month_end[:4]
                month_num = current_month_end[5:7]
                new_filename = re.sub(r"(\d{4})_(\d{2})", f"{year}_{month_num}", filename)
                if new_filename != filename:
                    new_path = os.path.join(root, new_filename)
                    if os.path.exists(new_path):  # jei failas jau yra, pridėti versiją
                        base, ext = os.path.splitext(new_filename)
                        i = 1
                        while True:
                            candidate = os.path.join(root, f"{base}_v{i}{ext}")
                            if not os.path.exists(candidate):
                                new_path = candidate
                                break
                            i += 1
                    os.rename(file_path, new_path)
                    log_lines.append(f"    📁 Pervadinta -> {os.path.join(rel_root, os.path.basename(new_path))}")

            except Exception as e:
                log_lines.append(f"    ❌ Klaida: {e}")


# ------------------------------
# Streamlit UI
# ------------------------------
st.set_page_config(page_title="Excel aktų atnaujinimas", page_icon="📄", layout="centered")

st.title("📄 Excel aktų atnaujinimas")
st.write(
    "Įkelkite **viso aplanko ZIP** (su poaplankiais). Programa atnaujins Excel failus (C5 datą, A9 mėnesį), "
    "prireikus pervadins failus, ir grąžins viską atgal ZIP formatu."
)

uploaded = st.file_uploader("Įkelkite aplanką kaip .zip", type=["zip"])

if uploaded is not None:
    with st.status("Apdorojama…", expanded=True) as status:
        logs = []
        try:
            base_dir, tmp_handle = unzip_to_temp(uploaded)
            logs.append("📦 ZIP sėkmingai išarchyvuotas.")

            process_excels_in_tree(base_dir, logs)

            out_bytes = zip_tree_to_bytes(base_dir)
            logs.append("🧷 Paruoštas atsisiunčiamas ZIP su rezultatais.")

            status.update(label="Apdorojimas baigtas.", state="complete")

            # Peržiūros langas (log'ai)
            st.text("\n".join(logs))

            st.download_button(
                label="⬇️ Parsisiųsti atnaujintą ZIP",
                data=out_bytes,
                file_name=f"atnaujinta_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip"
            )

            st.success("🎉 Excel failai atnaujinti sėkmingai!")

        except zipfile.BadZipFile:
            status.update(label="Nepavyko išarchyvuoti ZIP.", state="error")
            st.error("❌ Netinkamas ZIP failas.")
        except Exception as e:
            status.update(label="Įvyko klaida.", state="error")
            st.error(f"❌ Klaida: {e}")
else:
    st.info("👉 Įkelkite .zip failą su savo Excel failais.")


      

   
    
    
        
    

      
