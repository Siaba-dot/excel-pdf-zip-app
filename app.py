import io
import os
import re
import zipfile
import tempfile
import calendar
from datetime import datetime

import streamlit as st
from openpyxl import load_workbook


# =========================
# Pagalbinės funkcijos
# =========================
def get_current_month_end_and_name():
    """Grąžina (YYYY-MM-DD, mėnesio_pavadinimas_LT_genityvas)."""
    today = datetime.today()
    month_end_day = calendar.monthrange(today.year, today.month)[1]
    current_month_end = today.replace(day=month_end_day)
    month_names = [
        "sausio", "vasario", "kovo", "balandžio", "gegužės", "birželio",
        "liepos", "rugpjūčio", "rugsėjo", "spalio", "lapkričio", "gruodžio"
    ]
    return current_month_end.strftime("%Y-%m-%d"), month_names[today.month - 1]


def unzip_to_temp(uploaded_zip_file):
    """Išarchyvuoja ZIP į laikiną aplanką ir grąžina (dir_path, tmp_handle)."""
    tmp_dir = tempfile.TemporaryDirectory()
    zip_bytes = uploaded_zip_file.read()
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        zf.extractall(tmp_dir.name)
    return tmp_dir.name, tmp_dir


def zip_only_excels_to_bytes(root_dir: str) -> bytes:
    """Supakuoja tik .xlsx ir .xlsm failus, išlaikant poaplankių struktūrą."""
    mem_zip = io.BytesIO()
    with zipfile.ZipFile(mem_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for folder_name, _, filenames in os.walk(root_dir):
            for fn in filenames:
                if not fn.lower().endswith((".xlsx", ".xlsm")):
                    continue
                abs_path = os.path.join(folder_name, fn)
                arcname = os.path.relpath(abs_path, root_dir)
                zf.write(abs_path, arcname=arcname)
    mem_zip.seek(0)
    return mem_zip.read()


# =========================
# Pagrindinis apdorojimas
# =========================
def process_excels_streaming(base_dir: str, progress_cb, line_cb, done_cb):
    """
    Apdoroja visus .xlsx/.xlsm medyje ir „stream'ina“ būsenas į UI per callback'us:
      - progress_cb(current, total)
      - line_cb(text)   # eilutė po eilutės
      - done_cb()       # kai baigta
    """
    current_month_end, current_month_name = get_current_month_end_and_name()
    months = [
        "sausio", "vasario", "kovo", "balandžio", "gegužės", "birželio",
        "liepos", "rugpjūčio", "rugsėjo", "spalio", "lapkričio", "gruodžio"
    ]

    # Surenkam visų apdorotinų failų sąrašą iš anksto (kad žinotume total)
    excel_files = []
    for root, _, files in os.walk(base_dir):
        for filename in files:
            if filename.lower().endswith((".xlsx", ".xlsm")):
                excel_files.append(os.path.join(root, filename))

    total = len(excel_files)
    processed = 0

    if total == 0:
        line_cb("ℹ️ Nerasta nė vieno Excel failo (.xlsx, .xlsm).")
        done_cb()
        return

    for file_path in excel_files:
        rel_path = os.path.relpath(file_path, base_dir)
        try:
            # 1) Atnaujinimai Excel faile
            wb = load_workbook(file_path)
            sh = wb.active

            # C5 – mėnesio pabaigos data (jei langelis egzistuoja ir ne None)
            c5_changed = False
            try:
                if sh["C5"].value is not None:
                    dt = datetime.strptime(current_month_end, "%Y-%m-%d").date()
                    sh["C5"].value = dt
                    c5_changed = True
            except Exception as e:
                line_cb(f"⚠️ {rel_path}: nepavyko atnaujinti C5 ({e}).")

            # A9 – mėnesio pavadinimas (pakeičiam bet kurį mėnesį į einamą)
            a9_changed = False
            try:
                if sh["A9"].value is not None:
                    cell_value = str(sh["A9"].value).strip()
                    replaced = False
                    for month in months:
                        if re.search(month, cell_value, flags=re.IGNORECASE):
                            new_val = re.sub(month, current_month_name, cell_value, flags=re.IGNORECASE)
                            sh["A9"].value = new_val
                            replaced = True
                            a9_changed = True
                            break
                    if not replaced:
                        line_cb(f"ℹ️ {rel_path}: A9 nerastas mėnesio pavadinimas – nepakeista.")
            except Exception as e:
                line_cb(f"⚠️ {rel_path}: nepavyko atnaujinti A9 ({e}).")

            wb.save(file_path)
            wb.close()

            # 2) Pervadinimas pagal YYYY_MM (jei tokią dalį randa pavadinime)
            year = current_month_end[:4]
            month_num = current_month_end[5:7]
            filename = os.path.basename(file_path)
            new_filename = re.sub(r"(\d{4})_(\d{2})", f"{year}_{month_num}", filename)

            if new_filename != filename:
                new_path = os.path.join(os.path.dirname(file_path), new_filename)
                if os.path.exists(new_path):
                    base, ext = os.path.splitext(new_filename)
                    i = 1
                    while True:
                        candidate = os.path.join(os.path.dirname(file_path), f"{base}_v{i}{ext}")
                        if not os.path.exists(candidate):
                            new_path = candidate
                            break
                        i += 1
                os.rename(file_path, new_path)
                rel_new = os.path.relpath(new_path, base_dir)
                line_cb(f"✅ {rel_new}: atnaujinta C5 ({'OK' if c5_changed else 'skip'}), A9 ({'OK' if a9_changed else 'skip'}), pervadinta.")
            else:
                line_cb(f"✅ {rel_path}: atnaujinta C5 ({'OK' if c5_changed else 'skip'}), A9 ({'OK' if a9_changed else 'skip'}).")

            processed += 1
            progress_cb(processed, total)

        except Exception as e:
            line_cb(f"❌ {rel_path}: apdorojimo klaida – {e}")
            processed += 1
            progress_cb(processed, total)

    done_cb()


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="Excel aktų atnaujinimas (gyvas progresas)", page_icon="📄", layout="centered")

st.title("📄 Excel aktų atnaujinimas")
st.write(
    "Įkelkite **viso aplanko ZIP** (su poaplankiais). Programa **gyvai** rodys kiekvieną apdorotą failą: "
    "atnaujins C5 datą (mėnesio pabaiga), A9 mėnesio pavadinimą, prireikus pervadins failą, o pabaigoje leis "
    "atsisiųsti **tik Excel** failus ZIP formatu."
)

uploaded = st.file_uploader("Įkelkite aplanką kaip .zip", type=["zip"])

if uploaded is not None:
    # Paruošiame UI vietas „streaminimui“
    status_box = st.status("Apdorojama…", expanded=True)
    progress_bar = st.progress(0)
    counter_placeholder = st.empty()
    lines_container = st.container()  # čia dėsime eilučių sąrašą
    results_placeholder = st.empty()  # čia atsiras download mygtukas pabaigoje

    logs = []

    def progress_cb(done, total):
        progress_bar.progress(done / total)
        counter_placeholder.write(f"Progresas: **{done}/{total}**")

    def line_cb(text):
        logs.append(text)
        # išvedame tik paskutines N eilučių, kad neperpildytume UI
        N = 200
        lines_container.text("\n".join(logs[-N:]))

    def done_cb():
        status_box.update(label="Apdorojimas baigtas.", state="complete")

    try:
        base_dir, tmp_handle = unzip_to_temp(uploaded)
        line_cb("📦 ZIP sėkmingai išarchyvuotas.")

        # Paleidžiam apdorojimą su gyvu atnaujinimu
        process_excels_streaming(base_dir, progress_cb, line_cb, done_cb)

        # Sukuriame ZIP tik iš Excel failų
        out_bytes = zip_only_excels_to_bytes(base_dir)
        line_cb("🧷 Paruoštas atsisiunčiamas ZIP su atnaujintais Excel failais.")

        # Parsisiuntimui
        results_placeholder.download_button(
            label="⬇️ Parsisiųsti atnaujintus Excel (.zip)",
            data=out_bytes,
            file_name=f"atnaujinta_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
            mime="application/zip"
        )

        st.success("🎉 Visi Excel failai apdoroti!")

    except zipfile.BadZipFile:
        status_box.update(label="Nepavyko išarchyvuoti ZIP.", state="error")
        st.error("❌ Netinkamas ZIP failas.")
    except Exception as e:
        status_box.update(label="Įvyko klaida.", state="error")
        st.error(f"❌ Klaida: {e}")

else:
    st.info("👉 Įkelkite **.zip** su savo Excel failais.")
