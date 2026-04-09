# ─────────────────────────────────────────────
#  app.py  —  Pólizas de Provisión de Gastos
# ─────────────────────────────────────────────

import io
import os
import uuid
import threading
from flask import Flask, request, jsonify, send_file, render_template

from motor import procesar_excel, generar_polizas, generar_altas, generar_catalogo_cuentas

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB

JOBS = {}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/procesar", methods=["POST"])
def procesar():
    if "facturas" not in request.files or "catalogo" not in request.files:
        return jsonify({"error": "Se requieren ambos archivos: facturas y catálogo"}), 400

    try:
        num_pol = int(request.form.get("num_poliza", 1))
    except ValueError:
        return jsonify({"error": "Número de póliza inválido"}), 400

    facturas_bytes = request.files["facturas"].read()
    catalogo_bytes = request.files["catalogo"].read()

    job_id = str(uuid.uuid4())
    JOBS[job_id] = {
        "estado": "procesando", "progreso": 0, "total": 0,
        "resultado_polizas": None, "resultado_altas": None, "resultado_cuentas": None,
        "stats": None, "error": None,
    }

    def run():
        try:
            def cb(done, total):
                JOBS[job_id]["progreso"] = done
                JOBS[job_id]["total"]    = total

            facturas, nuevos, stats, ultimo_cod = procesar_excel(
                facturas_bytes, catalogo_bytes, num_pol, cb
            )

            polizas_xlsx = generar_polizas(facturas, num_pol)
            altas_xlsx   = generar_altas(nuevos, ultimo_cod, catalogo_bytes) if nuevos else None
            cuentas_xlsx = generar_catalogo_cuentas(nuevos) if nuevos else None

            JOBS[job_id]["resultado_polizas"] = polizas_xlsx
            JOBS[job_id]["resultado_altas"]   = altas_xlsx
            JOBS[job_id]["resultado_cuentas"] = cuentas_xlsx
            JOBS[job_id]["stats"]  = stats
            JOBS[job_id]["estado"] = "listo"
        except Exception as e:
            import traceback
            JOBS[job_id]["estado"] = "error"
            JOBS[job_id]["error"]  = str(e) + "\n" + traceback.format_exc()

    threading.Thread(target=run, daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/progreso/<job_id>")
def progreso(job_id):
    job = JOBS.get(job_id)
    if not job:
        return jsonify({"error": "Job no encontrado"}), 404
    return jsonify({
        "estado":   job["estado"],
        "progreso": job["progreso"],
        "total":    job["total"],
        "stats":    job["stats"],
        "error":    job["error"],
        "tiene_altas":   job["resultado_altas"]   is not None,
        "tiene_cuentas": job["resultado_cuentas"] is not None,
    })


@app.route("/descargar/polizas/<job_id>")
def descargar_polizas(job_id):
    job = JOBS.get(job_id)
    if not job or job["estado"] != "listo":
        return jsonify({"error": "No disponible"}), 404
    return send_file(
        io.BytesIO(job["resultado_polizas"]),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="Polizas_Provision_Gastos.xlsx",
    )


@app.route("/descargar/cuentas/<job_id>")
def descargar_cuentas(job_id):
    job = JOBS.get(job_id)
    if not job or job["estado"] != "listo" or not job["resultado_cuentas"]:
        return jsonify({"error": "No hay catálogo de cuentas"}), 404
    return send_file(
        io.BytesIO(job["resultado_cuentas"]),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="Alta_Catalogo_Cuentas.xlsx",
    )


@app.route("/descargar/altas/<job_id>")
def descargar_altas(job_id):
    job = JOBS.get(job_id)
    if not job or job["estado"] != "listo" or not job["resultado_altas"]:
        return jsonify({"error": "No hay altas de proveedores"}), 404
    return send_file(
        io.BytesIO(job["resultado_altas"]),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="Alta_Nuevos_Proveedores.xlsx",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5051, debug=False)
