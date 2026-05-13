import os
import shutil
import tempfile
import threading
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from uuid import uuid4

from flask import Flask, Response, flash, redirect, render_template_string, request, send_file, url_for

import Consolidado_Emendas as base
import Consolidado_Emendas_V2 as v2


APP_DIR = Path(__file__).resolve().parent
REFERENCE_XLSX = APP_DIR / "Planilha_emendas_pl4_2025 atualizada até 09.04.2026.xlsx"
DOWNLOAD_DIR = APP_DIR / "_web_outputs"
DOWNLOAD_DIR.mkdir(exist_ok=True)
WORK_DIR = APP_DIR / "_web_work"
WORK_DIR.mkdir(exist_ok=True)

MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "900"))

JOB_LOCK = threading.Lock()
JOB_STATE = {
    "running": False,
    "kind": "",
    "message": "",
    "done": 0,
    "total": 0,
    "download": "",
    "error": "",
}

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "troque-esta-chave-em-producao")
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


PAGE = """
<!doctype html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  {% if job.running %}<meta http-equiv="refresh" content="5">{% endif %}
  <title>Consolidado Emendas PL 04/2025 - Reforma do Código Civil</title>
  <style>
    :root { color-scheme: light; font-family: Arial, sans-serif; }
    body { margin: 0; background: #f5f7fb; color: #172033; }
    main { max-width: 1120px; margin: 0 auto; padding: 32px 20px 48px; }
    h1 { margin: 0 0 6px; font-size: 30px; }
    p { line-height: 1.45; }
    .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 18px; margin-top: 22px; }
    .card { background: #fff; border: 1px solid #d8deea; border-radius: 8px; padding: 18px; box-shadow: 0 1px 2px rgba(16,24,40,.06); }
    .card h2 { margin: 0 0 10px; font-size: 20px; }
    label { display: block; font-weight: 700; margin: 14px 0 6px; }
    input[type=file] { width: 100%; box-sizing: border-box; border: 1px solid #c7cfdd; border-radius: 6px; padding: 10px; background: #fbfcff; }
    button { margin-top: 16px; border: 0; border-radius: 6px; padding: 11px 14px; background: #1652a3; color: white; font-weight: 700; cursor: pointer; }
    button:hover { background: #103f7d; }
    button:disabled { background: #9aa8ba; cursor: not-allowed; }
    .hint { font-size: 13px; color: #526173; margin-top: 8px; }
    .flash { background: #fff7e6; border: 1px solid #ffd58a; border-radius: 8px; padding: 12px 14px; margin: 16px 0; }
    .download { background: #eef8f1; border: 1px solid #a7ddb6; border-radius: 8px; padding: 12px 14px; margin: 16px 0; }
    .download a { font-weight: 700; color: #0f6b2c; }
    .status { background: #eef4ff; border: 1px solid #b8cdf5; border-radius: 8px; padding: 12px 14px; margin: 16px 0; }
    .bar { height: 10px; background: #d7e2f7; border-radius: 999px; overflow: hidden; margin-top: 8px; }
    .bar span { display: block; height: 100%; background: #1652a3; width: {{ job.percent }}%; }
  </style>
</head>
<body>
  <main>
    <h1>Consolidado Emendas PL 04/2025 - Reforma do Código Civil</h1>
    <p>Envie os arquivos e aguarde o processamento. O resultado será uma planilha Excel pronta para baixar.</p>

    {% with messages = get_flashed_messages() %}
      {% if messages %}
        {% for message in messages %}
          <div class="flash">{{ message }}</div>
        {% endfor %}
      {% endif %}
    {% endwith %}

    {% if job.error %}
      <div class="flash">{{ job.error }}</div>
    {% endif %}

    {% if download_url %}
      <div class="download">Arquivo gerado: <a href="{{ download_url }}">baixar planilha</a></div>
    {% endif %}

    {% if job.running %}
      <div class="status">
        <strong>{{ job.kind }}</strong>: {{ job.message }}
        {% if job.total %}
          <div>{{ job.done }} de {{ job.total }} PDFs processados</div>
          <div class="bar"><span></span></div>
        {% endif %}
      </div>
    {% endif %}

    <section class="grid">
      <form class="card" method="post" action="{{ url_for('run_v1') }}" enctype="multipart/form-data">
        <h2>Gerar Consolidado completo</h2>
        <p>Use quando quiser processar uma pasta zipada contendo todos os PDFs das emendas.</p>
        <label for="pdf_zip_v1">ZIP com emendas em PDF</label>
        <input id="pdf_zip_v1" name="pdf_zip" type="file" accept=".zip" required>
        <div class="hint">Os PDFs devem manter o padrão de nome: 001 - Autor.pdf.</div>
        <button type="submit" {% if job.running %}disabled{% endif %}>Gerar Consolidado_Emendas.xlsx</button>
      </form>

      <form class="card" method="post" action="{{ url_for('run_v2') }}" enctype="multipart/form-data">
        <h2>Gerar Consolidado Incremental</h2>
        <p>Use quando já houver uma planilha Emendas.xlsx revisada manualmente e novas emendas em PDF a serem incrementadas na Planilha.</p>
        <label for="emendas_xlsx">Planilha Emendas.xlsx</label>
        <input id="emendas_xlsx" name="emendas_xlsx" type="file" accept=".xlsx" required>
        <label for="pdf_zip_v2">ZIP com novas emendas em PDF</label>
        <input id="pdf_zip_v2" name="pdf_zip" type="file" accept=".zip">
        <div class="hint">O Consolidado Incremental copia a aba Emendas da planilha enviada e acrescenta PDFs ainda não listados. Se nenhum ZIP for enviado, ele gera as abas derivadas apenas a partir da planilha.</div>
        <button type="submit" {% if job.running %}disabled{% endif %}>Gerar Consolidado Incremental.xlsx</button>
      </form>
    </section>
  </main>
</body>
</html>
"""


def job_snapshot() -> dict:
    with JOB_LOCK:
        job = dict(JOB_STATE)
    total = job.get("total") or 0
    done = job.get("done") or 0
    job["percent"] = 0 if not total else min(100, int(done * 100 / total))
    return job


def start_job(kind: str, message: str) -> bool:
    with JOB_LOCK:
        if JOB_STATE["running"]:
            return False
        JOB_STATE.update({
            "running": True,
            "kind": kind,
            "message": message,
            "done": 0,
            "total": 0,
            "download": "",
            "error": "",
        })
        return True


def update_job(**changes) -> None:
    with JOB_LOCK:
        JOB_STATE.update(changes)


def finish_job(download: str = "", error: str = "") -> None:
    with JOB_LOCK:
        JOB_STATE.update({
            "running": False,
            "download": download,
            "error": error,
            "message": "Concluido" if download else "Erro",
        })


def safe_extract_zip(zip_path: Path, target_dir: Path) -> None:
    with zipfile.ZipFile(zip_path) as archive:
        for member in archive.infolist():
            if member.is_dir():
                continue
            member_path = Path(member.filename)
            if member_path.is_absolute() or ".." in member_path.parts:
                raise ValueError(f"Arquivo invalido dentro do ZIP: {member.filename}")
            if member_path.suffix.lower() != ".pdf":
                continue
            output_path = target_dir / member_path.name
            with archive.open(member) as source, output_path.open("wb") as target:
                shutil.copyfileobj(source, target)


def save_upload(field_name: str, target_path: Path) -> None:
    uploaded = request.files.get(field_name)
    if not uploaded or not uploaded.filename:
        raise ValueError(f"Arquivo obrigatorio nao enviado: {field_name}")
    uploaded.save(target_path)


def save_optional_upload(field_name: str, target_path: Path) -> bool:
    uploaded = request.files.get(field_name)
    if not uploaded or not uploaded.filename:
        return False
    uploaded.save(target_path)
    return True


def create_work_dir(prefix: str) -> Path:
    return Path(tempfile.mkdtemp(prefix=prefix, dir=WORK_DIR))


def cleanup_work_dir(path: Path) -> None:
    try:
        shutil.rmtree(path)
    except OSError:
        # Windows can keep recently processed upload files locked briefly.
        pass


def list_pdf_paths(pdf_dir: Path) -> list[Path]:
    pdfs = [path for path in pdf_dir.glob("*.pdf") if base.PDF_NAME_RE.match(path.name)]
    return sorted(pdfs, key=lambda path: int(base.PDF_NAME_RE.match(path.name).group(1)))


def build_rows_from_pdfs(pdf_paths: list[Path]) -> tuple[list[base.EmendaRow], list[str]]:
    rows: list[base.EmendaRow] = []
    failures: list[str] = []
    workers = max(2, min(8, (os.cpu_count() or 4)))
    base.OCR_PROGRESS["current"] = 0

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = [executor.submit(base.build_row_with_fallback, pdf) for pdf in pdf_paths]
        done = 0
        for future in as_completed(futures):
            row, failure = future.result()
            rows.append(row)
            done += 1
            update_job(done=done, total=len(pdf_paths), message="Lendo e classificando PDFs")
            if failure:
                failures.append(failure)

    return sorted(rows, key=lambda row: row.numero), failures


def copy_reference_if_available(work_dir: Path) -> None:
    if REFERENCE_XLSX.exists():
        shutil.copy2(REFERENCE_XLSX, work_dir / REFERENCE_XLSX.name)


def run_full_consolidado(work_dir: Path, pdf_dir: Path, output_path: Path) -> list[str]:
    pdfs = list_pdf_paths(pdf_dir)
    if not pdfs:
        raise ValueError("Nenhum PDF com nome no padrao '001 - Autor.pdf' foi encontrado no ZIP.")

    update_job(done=0, total=len(pdfs), message="Lendo e classificando PDFs")
    rows, failures = build_rows_from_pdfs(pdfs)
    old_base_dir = base.BASE_DIR
    try:
        base.BASE_DIR = work_dir
        copy_reference_if_available(work_dir)
        update_job(message="Montando planilha Excel")
        wb = base.build_workbook(rows)
    finally:
        base.BASE_DIR = old_base_dir

    update_job(message="Salvando planilha Excel")
    wb.save(output_path)
    return failures


def run_incremental_v2(work_dir: Path, emendas_xlsx: Path, pdf_dir: Path, output_path: Path) -> None:
    existing_rows = v2.load_existing_rows(emendas_xlsx)
    existing_numbers = {row.numero for row in existing_rows}
    new_pdf_paths = [
        path for path in list_pdf_paths(pdf_dir)
        if int(base.PDF_NAME_RE.match(path.name).group(1)) not in existing_numbers
    ]
    update_job(done=0, total=len(new_pdf_paths), message="Lendo novas emendas")
    new_rows = []
    for index, path in enumerate(new_pdf_paths, start=1):
        new_rows.append(v2.build_new_row_from_pdf(path))
        update_job(done=index, total=len(new_pdf_paths), message="Lendo novas emendas")
    update_job(message="Montando planilha Excel")
    wb = v2.build_workbook_v2(existing_rows, new_rows, emendas_xlsx)
    update_job(message="Salvando planilha Excel")
    wb.save(output_path)
    v2.restore_source_emendas_xml(output_path, emendas_xlsx, new_rows)


def run_v1_background(work_dir: Path, zip_path: Path, output_path: Path) -> None:
    try:
        pdf_dir = work_dir / "pdfs"
        pdf_dir.mkdir()
        update_job(message="Extraindo PDFs do ZIP")
        safe_extract_zip(zip_path, pdf_dir)
        failures = run_full_consolidado(work_dir, pdf_dir, output_path)
        error = f"Planilha gerada, mas houve {len(failures)} falha(s) de leitura em PDFs." if failures else ""
        finish_job(download=output_path.name, error=error)
    except Exception as exc:
        finish_job(error=f"Erro ao gerar o consolidado: {exc}")
    finally:
        cleanup_work_dir(work_dir)


def run_v2_background(work_dir: Path, zip_path: Path | None, emendas_xlsx: Path, output_path: Path) -> None:
    try:
        pdf_dir = work_dir / "pdfs"
        pdf_dir.mkdir()
        if zip_path:
            update_job(message="Extraindo PDFs do ZIP")
            safe_extract_zip(zip_path, pdf_dir)
        else:
            update_job(message="Gerando abas a partir da planilha enviada")
        run_incremental_v2(work_dir, emendas_xlsx, pdf_dir, output_path)
        finish_job(download=output_path.name)
    except Exception as exc:
        finish_job(error=f"Erro ao gerar o Consolidado Incremental: {exc}")
    finally:
        cleanup_work_dir(work_dir)


@app.get("/")
def index():
    job = job_snapshot()
    download_name = request.args.get("download") or job.get("download")
    download_url = url_for("download_file", filename=download_name) if download_name else ""
    return render_template_string(PAGE, download_url=download_url, job=job)


@app.post("/v1")
def run_v1():
    job_id = uuid4().hex
    output_path = DOWNLOAD_DIR / f"Consolidado_Emendas_{job_id}.xlsx"
    if not start_job("Consolidado completo", "Recebendo arquivo enviado"):
        flash("Ja existe um processamento em andamento. Aguarde a conclusao antes de iniciar outro.")
        return redirect(url_for("index"))

    work_dir = create_work_dir("emendas_v1_")
    zip_path = work_dir / "emendas.zip"
    try:
        save_upload("pdf_zip", zip_path)
    except Exception as exc:
        cleanup_work_dir(work_dir)
        finish_job(error=f"Erro ao receber o ZIP: {exc}")
        return redirect(url_for("index"))

    thread = threading.Thread(target=run_v1_background, args=(work_dir, zip_path, output_path), daemon=True)
    thread.start()
    return redirect(url_for("index"))


@app.post("/v2")
def run_v2():
    job_id = uuid4().hex
    output_path = DOWNLOAD_DIR / f"Consolidado_Incremental_{job_id}.xlsx"
    if not start_job("Consolidado Incremental", "Recebendo arquivos enviados"):
        flash("Ja existe um processamento em andamento. Aguarde a conclusao antes de iniciar outro.")
        return redirect(url_for("index"))

    work_dir = create_work_dir("emendas_v2_")
    zip_path = work_dir / "novas_emendas.zip"
    emendas_xlsx = work_dir / "Emendas.xlsx"
    try:
        save_upload("emendas_xlsx", emendas_xlsx)
        has_zip = save_optional_upload("pdf_zip", zip_path)
    except Exception as exc:
        cleanup_work_dir(work_dir)
        finish_job(error=f"Erro ao receber os arquivos: {exc}")
        return redirect(url_for("index"))

    selected_zip = zip_path if has_zip else None
    thread = threading.Thread(target=run_v2_background, args=(work_dir, selected_zip, emendas_xlsx, output_path), daemon=True)
    thread.start()
    return redirect(url_for("index"))


@app.get("/download/<filename>")
def download_file(filename: str):
    safe_name = Path(filename).name
    path = DOWNLOAD_DIR / safe_name
    if not path.exists():
        return Response("Arquivo nao encontrado.", status=404)
    return send_file(path, as_attachment=True, download_name=safe_name)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")))
