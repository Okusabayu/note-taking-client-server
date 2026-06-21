from flask import Flask, jsonify, request, render_template
import sqlite3

app = Flask(__name__)

def get_db_connection():
    conn = sqlite3.connect("database.db")
    conn.row_factory = sqlite3.Row
    return conn

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/notes", methods=["GET"])
def get_notes():

    conn = get_db_connection()

    notes = conn.execute(
        "SELECT * FROM notes"
    ).fetchall()

    conn.close()

    return jsonify([
        {
            "id": note["id"],
            "title": note["title"],
            "content": note["content"]
        }
        for note in notes
    ])

@app.route("/notes", methods=["POST"])
def create_note():

    data = request.get_json()

    title = data.get("title")
    content = data.get("content")

    if not title or not content:
        return jsonify({
            "error": "Title dan content wajib diisi"
        }), 400

    conn = get_db_connection()

    conn.execute(
        "INSERT INTO notes (title, content) VALUES (?, ?)",
        (title, content)
    )

    conn.commit()
    conn.close()

    return jsonify({
        "message": "Catatan berhasil ditambahkan"
    }), 201

@app.route("/notes/<int:id>", methods=["PUT"])
def update_note(id):

    data = request.get_json()

    title = data.get("title")
    content = data.get("content")

    if not title or not content:
        return jsonify({
            "error": "Title dan content wajib diisi"
        }), 400

    conn = get_db_connection()

    note = conn.execute(
        "SELECT * FROM notes WHERE id = ?",
        (id,)
    ).fetchone()

    if note is None:
        conn.close()
        return jsonify({
            "error": "Catatan tidak ditemukan"
        }), 404

    conn.execute(
        "UPDATE notes SET title = ?, content = ? WHERE id = ?",
        (title, content, id)
    )

    conn.commit()
    conn.close()

    return jsonify({
        "message": "Catatan berhasil diperbarui"
    })

@app.route("/notes/<int:id>", methods=["DELETE"])
def delete_note(id):

    conn = get_db_connection()

    note = conn.execute(
        "SELECT * FROM notes WHERE id = ?",
        (id,)
    ).fetchone()

    if note is None:
        conn.close()
        return jsonify({
            "error": "Catatan tidak ditemukan"
        }), 404

    conn.execute(
        "DELETE FROM notes WHERE id = ?",
        (id,)
    )

    conn.commit()
    conn.close()

    return jsonify({
        "message": "Catatan berhasil dihapus"
    })

if __name__ == "__main__":
    app.run(debug=True)