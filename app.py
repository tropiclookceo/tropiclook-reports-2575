#!/usr/bin/env python3
"""
app.py — Flask wrapper for the TropicLook Owner Report engine, NEW MODEL
25/75 after channel costs (OTA / agent).

Deployed as a SEPARATE Render.com service so the old-model pipeline keeps
serving properties that have not been re-signed yet.

Endpoints (same contract as the old service, so the Make.com scenario is a
straight clone pointing to the new URL and new Google Drive folders):

    GET  /health               -> {"status": "ok", "engine": "..."}
    POST /generate             -> multipart 'file' = InputData xlsx,
                                  returns OwnerReport xlsx
    POST /generate-next-input  -> multipart 'file' = InputData xlsx,
                                  returns next-month InputData xlsx

Auth: header X-API-Token must match the API_TOKEN environment variable.
"""

import io
import os
import tempfile

from flask import Flask, request, jsonify, send_file

import tl_report_engine_25_75 as engine

app = Flask(__name__)

API_TOKEN = os.environ.get("API_TOKEN", "tropiclook-2575-change-me")


def _check_token():
    return request.headers.get("X-API-Token") == API_TOKEN


def _save_upload():
    """Save the uploaded InputData file to a temp path, return (path, name)."""
    up = request.files.get("file")
    if up is None or up.filename == "":
        return None, None
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    up.save(path)
    return path, up.filename


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "engine": engine.ENGINE_VERSION})


@app.route("/generate", methods=["POST"])
def generate():
    if not _check_token():
        return jsonify({"error": "unauthorized"}), 401
    path, name = _save_upload()
    if not path:
        return jsonify({"error": "no file uploaded (multipart field 'file')"}), 400
    try:
        report_bytes, warnings = engine.generate_report(path)
        out_name = os.path.basename(name).replace("InputData", "OwnerReport")
        resp = send_file(
            io.BytesIO(report_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=out_name,
        )
        if warnings:
            resp.headers["X-Report-Warnings"] = " | ".join(w[:180] for w in warnings)[:900]
        return resp
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 422
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


@app.route("/generate-next-input", methods=["POST"])
def generate_next_input():
    if not _check_token():
        return jsonify({"error": "unauthorized"}), 401
    path, name = _save_upload()
    if not path:
        return jsonify({"error": "no file uploaded (multipart field 'file')"}), 400
    try:
        next_bytes, meta = engine.generate_next_input_template(path)
        # Derive the next-month name from the ORIGINAL uploaded filename,
        # not from the temp path on the server.
        next_name = engine._next_input_filename(
            os.path.basename(name), meta["current_period"], meta["next_period"]
        )
        resp = send_file(
            io.BytesIO(next_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=next_name,
        )
        resp.headers["X-Next-Period"] = meta["next_period"]
        return resp
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 422
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
