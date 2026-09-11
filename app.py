#!/usr/bin/env python3
"""SGR tools — servidor Flask (HTML + lógica Python)."""

from __future__ import annotations

import os
import sys
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
os.chdir(ROOT)
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from flask import Flask, Response, jsonify, request, send_from_directory
from flask_cors import CORS

from backend.calendario_global import (
    add_manual_event,
    get_global_calendar,
    move_global_event,
)
from backend.calendario_rentas import get_saved_result, process_position
from backend.conciliacion import (
    conciliar,
    export_conciliacion_excel,
    get_saved_conciliacion,
)
from backend.contingente import (
    export_excel,
    export_pdf,
    get_saved_contingente,
    procesar_contingente,
)
from backend.rendimientos import (
    export_returns_excel,
    export_returns_pdf,
    get_saved_returns,
    process_returns,
)
from backend.manuales import MANUAL_PDF, build_user_manual

app = Flask(__name__, static_folder=None)
CORS(app)


def _parse_float(raw: str | None, default: float) -> float:
    if raw is None or str(raw).strip() == "":
        return default
    return float(str(raw).strip().replace(",", "."))


def _parse_int(raw: str | None, default: int) -> int:
    if raw is None or str(raw).strip() == "":
        return default
    return int(float(str(raw).strip().replace(",", ".")))


# --- API (registrar ANTES del catch-all de archivos estáticos) ---


@app.get("/api/health")
def health():
    return jsonify({"ok": True, "app": "SGR tools"})


@app.post("/api/contingente/procesar")
def api_procesar():
    if "file" not in request.files:
        return jsonify({"error": "Falta el archivo (campo 'file')."}), 400
    f = request.files["file"]
    if not f or not f.filename:
        return jsonify({"error": "Archivo vacío."}), 400
    try:
        dias = _parse_int(
            request.form.get("dias_mesa") or request.form.get("dias_habiles_mesa"), 28
        )
        modo = request.form.get("modo_dias") or "corridos"
        aforo = _parse_float(request.form.get("aforo_pct"), 5.0)
        data = procesar_contingente(
            f.read(),
            f.filename,
            dias_mesa=dias,
            modo_dias=modo,
            aforo_pct=aforo,
        )
        return jsonify(data)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.get("/api/contingente/resultados")
def api_contingente_resultados():
    return jsonify(get_saved_contingente())


@app.post("/api/contingente/export/excel")
def api_export_excel():
    payload = request.get_json(silent=True) or {}
    rows = payload.get("rows") or []
    columns = payload.get("columns") or []
    if not rows or not columns:
        return jsonify({"error": "Sin datos para exportar."}), 400
    content = export_excel(rows, columns)
    return Response(
        content,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": "attachment; filename=contingente_procesado.xlsx"
        },
    )


@app.post("/api/contingente/export/pdf")
def api_export_pdf():
    payload = request.get_json(silent=True) or {}
    rows = payload.get("rows") or []
    columns = payload.get("columns") or []
    title = payload.get("title") or "Contingente avales bancarios"
    if not rows or not columns:
        return jsonify({"error": "Sin datos para exportar."}), 400
    content = export_pdf(rows, columns, title=title)
    return Response(
        content,
        mimetype="application/pdf",
        headers={"Content-Disposition": "attachment; filename=contingente_procesado.pdf"},
    )


@app.post("/api/conciliacion/conciliar")
def api_conciliar():
    if "custodia" not in request.files or "posicion" not in request.files:
        return jsonify({"error": "Faltan archivos: 'custodia' y 'posicion'."}), 400
    cust = request.files["custodia"]
    pos = request.files["posicion"]
    if not cust or not cust.filename or not pos or not pos.filename:
        return jsonify({"error": "Archivos vacíos."}), 400
    try:
        data = conciliar(
            cust.read(),
            pos.read(),
            custodia_name=cust.filename,
            posicion_name=pos.filename,
        )
        return jsonify(data)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.get("/api/conciliacion/resultados")
def api_conciliacion_resultados():
    return jsonify(get_saved_conciliacion())


@app.post("/api/conciliacion/export/excel")
def api_conciliacion_excel():
    payload = request.get_json(silent=True) or {}
    rows = payload.get("rows") or []
    if not rows:
        return jsonify({"error": "Sin datos para exportar."}), 400
    content = export_conciliacion_excel(rows)
    return Response(
        content,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": "attachment; filename=conciliacion_custodia_posicion.xlsx"
        },
    )


@app.get("/api/calendario-rentas/resultados")
def api_calendario_rentas_resultados():
    return jsonify(get_saved_result())


@app.post("/api/calendario-rentas/procesar")
def api_calendario_rentas_procesar():
    if "posicion" not in request.files:
        return jsonify({"error": "Falta el archivo de posición."}), 400
    position = request.files["posicion"]
    if not position or not position.filename:
        return jsonify({"error": "Archivo de posición vacío."}), 400
    try:
        return jsonify(process_position(position.read(), position.filename))
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.get("/api/calendario-global/resultados")
def api_calendario_global_resultados():
    return jsonify(get_global_calendar())


@app.post("/api/calendario-global/mover")
def api_calendario_global_mover():
    payload = request.get_json(silent=True) or {}
    try:
        return jsonify(
            move_global_event(payload.get("event_id"), payload.get("date"))
        )
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.post("/api/calendario-global/manual")
def api_calendario_global_manual():
    payload = request.get_json(silent=True) or {}
    try:
        return jsonify(
            add_manual_event(
                payload.get("type"),
                payload.get("denomination"),
                payload.get("date"),
                payload.get("comment"),
            )
        )
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.get("/api/rendimientos/resultados")
def api_rendimientos_resultados():
    return jsonify(get_saved_returns())


@app.post("/api/rendimientos/procesar")
def api_rendimientos_procesar():
    files = [
        (uploaded.read(), uploaded.filename)
        for uploaded in request.files.getlist("files")
        if uploaded and uploaded.filename
    ]
    fx_files = [
        (uploaded.read(), uploaded.filename)
        for uploaded in request.files.getlist("fx_files")
        if uploaded and uploaded.filename
    ]
    retiro_files = [
        (uploaded.read(), uploaded.filename)
        for uploaded in request.files.getlist("retiro_files")
        if uploaded and uploaded.filename
    ]
    aporte_files = [
        (uploaded.read(), uploaded.filename)
        for uploaded in request.files.getlist("aporte_files")
        if uploaded and uploaded.filename
    ]
    pago_files = [
        (uploaded.read(), uploaded.filename)
        for uploaded in request.files.getlist("pago_files")
        if uploaded and uploaded.filename
    ]
    flow_files = [
        (uploaded.read(), uploaded.filename)
        for uploaded in request.files.getlist("flow_files")
        if uploaded and uploaded.filename
    ]
    try:
        manual = request.form.get("manual") or "[]"
        manual_rows = json.loads(manual)
        if not isinstance(manual_rows, list):
            raise ValueError("Los datos manuales no son válidos.")
        manual_fx = json.loads(request.form.get("manual_fx") or "[]")
        if not isinstance(manual_fx, list):
            raise ValueError("Los tipos de cambio manuales no son válidos.")
        manual_flows = json.loads(request.form.get("manual_flows") or "[]")
        if not isinstance(manual_flows, list):
            raise ValueError("Los movimientos manuales no son válidos.")
        return jsonify(
            process_returns(
                files,
                manual_rows,
                fx_files,
                manual_fx,
                retiro_files,
                aporte_files,
                pago_files,
                flow_files,
                manual_flows,
            )
        )
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.get("/api/rendimientos/export/excel")
def api_rendimientos_excel():
    try:
        data = get_saved_returns()
        if not data.get("monthly"):
            raise ValueError("No hay rendimientos para exportar.")
        return Response(
            export_returns_excel(data),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": "attachment; filename=reporte_rendimientos.xlsx",
                "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
                "Pragma": "no-cache",
                "Expires": "0",
            },
        )
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.get("/api/rendimientos/export/pdf")
def api_rendimientos_pdf():
    try:
        data = get_saved_returns()
        return Response(
            export_returns_pdf(data),
            mimetype="application/pdf",
            headers={
                "Content-Disposition": "attachment; filename=reporte_rendimientos.pdf",
                "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
                "Pragma": "no-cache",
                "Expires": "0",
            },
        )
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


# --- Archivos estáticos (al final) ---


@app.get("/")
def home():
    return send_from_directory(ROOT, "index.html")


@app.get("/docs/manual-sgr-tools.pdf")
def manual_pdf():
    if not MANUAL_PDF.is_file():
        build_user_manual()
    return send_from_directory(
        MANUAL_PDF.parent,
        MANUAL_PDF.name,
        as_attachment=True,
        download_name="manual-sgr-tools.pdf",
    )


@app.get("/<path:path>")
def static_files(path: str):
    if path.startswith(("api/", "data/")):
        return jsonify({"error": "No encontrado"}), 404
    target = ROOT / path
    if target.is_file():
        return send_from_directory(ROOT, path)
    return jsonify({"error": "No encontrado"}), 404


if __name__ == "__main__":
    print("SGR tools -> http://127.0.0.1:5055")
    print("No abras los HTML directo: usa esa URL en el navegador.")
    # En Windows el reloader puede romper imports desde OneDrive.
    use_reload = os.environ.get("FLASK_DEBUG", "1") == "1" and sys.platform != "win32"
    app.run(host="127.0.0.1", port=5055, debug=True, use_reloader=use_reload)
