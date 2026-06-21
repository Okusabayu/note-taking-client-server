// ============================================
// Note Taking App - JavaScript (Fetch API)
// ============================================

const API_URL = "/notes";

// DOM Elements
const noteForm = document.getElementById("noteForm");
const titleInput = document.getElementById("title");
const contentInput = document.getElementById("content");
const submitBtn = document.getElementById("submitBtn");
const cancelBtn = document.getElementById("cancelBtn");
const notesList = document.getElementById("notesList");
const toast = document.getElementById("toast");

// State: track if we are editing
let editingId = null;

// ============================================
// 1. LOAD / GET ALL NOTES
// ============================================
async function loadNotes() {
    notesList.innerHTML = '<div class="loading">Memuat catatan...</div>';

    try {
        const response = await fetch(API_URL);
        const notes = await response.json();

        if (notes.length === 0) {
            notesList.innerHTML = `
                <div class="empty-state">
                    <div class="empty-icon">📋</div>
                    <p>Belum ada catatan. Tambahkan catatan pertamamu!</p>
                </div>
            `;
            return;
        }

        notesList.innerHTML = "";

        notes.forEach(function (note) {
            const noteEl = document.createElement("div");
            noteEl.className = "note-item";
            noteEl.innerHTML = `
                <h3>${escapeHtml(note.title)}</h3>
                <p>${escapeHtml(note.content)}</p>
                <div class="note-actions">
                    <button class="btn btn-sm btn-edit" onclick="startEdit(${note.id}, '${escapeJs(note.title)}', '${escapeJs(note.content)}')">
                        ✏️ Edit
                    </button>
                    <button class="btn btn-sm btn-delete" onclick="deleteNote(${note.id})">
                        🗑️ Hapus
                    </button>
                </div>
            `;
            notesList.appendChild(noteEl);
        });

    } catch (error) {
        notesList.innerHTML = '<div class="loading">Gagal memuat catatan.</div>';
        showToast("Gagal memuat catatan dari server.", "error");
    }
}

// ============================================
// 2. CREATE NOTE (POST)
// ============================================
async function createNote(title, content) {
    try {
        const response = await fetch(API_URL, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ title: title, content: content })
        });

        if (response.ok) {
            showToast("Catatan berhasil ditambahkan!", "success");
            noteForm.reset();
            loadNotes();
        } else {
            const data = await response.json();
            showToast(data.error || "Gagal menambahkan catatan.", "error");
        }
    } catch (error) {
        showToast("Terjadi kesalahan koneksi.", "error");
    }
}

// ============================================
// 3. UPDATE NOTE (PUT)
// ============================================
async function updateNote(id, title, content) {
    try {
        const response = await fetch(API_URL + "/" + id, {
            method: "PUT",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ title: title, content: content })
        });

        if (response.ok) {
            showToast("Catatan berhasil diperbarui!", "success");
            resetForm();
            loadNotes();
        } else {
            const data = await response.json();
            showToast(data.error || "Gagal memperbarui catatan.", "error");
        }
    } catch (error) {
        showToast("Terjadi kesalahan koneksi.", "error");
    }
}

// ============================================
// 4. DELETE NOTE (DELETE)
// ============================================
async function deleteNote(id) {
    if (!confirm("Yakin ingin menghapus catatan ini?")) {
        return;
    }

    try {
        const response = await fetch(API_URL + "/" + id, {
            method: "DELETE"
        });

        if (response.ok) {
            showToast("Catatan berhasil dihapus!", "success");
            loadNotes();
        } else {
            const data = await response.json();
            showToast(data.error || "Gagal menghapus catatan.", "error");
        }
    } catch (error) {
        showToast("Terjadi kesalahan koneksi.", "error");
    }
}

// ============================================
// 5. EDIT MODE
// ============================================
function startEdit(id, title, content) {
    editingId = id;
    titleInput.value = title;
    contentInput.value = content;
    submitBtn.innerHTML = '<span class="btn-icon">💾</span> Simpan Perubahan';
    cancelBtn.style.display = "inline-flex";

    // Scroll ke form
    noteForm.scrollIntoView({ behavior: "smooth" });
    titleInput.focus();
}

function resetForm() {
    editingId = null;
    noteForm.reset();
    submitBtn.innerHTML = '<span class="btn-icon">＋</span> Tambah Catatan';
    cancelBtn.style.display = "none";
}

// ============================================
// 6. FORM SUBMIT HANDLER
// ============================================
noteForm.addEventListener("submit", function (e) {
    e.preventDefault();

    const title = titleInput.value.trim();
    const content = contentInput.value.trim();

    if (!title || !content) {
        showToast("Judul dan isi catatan wajib diisi.", "error");
        return;
    }

    if (editingId) {
        updateNote(editingId, title, content);
    } else {
        createNote(title, content);
    }
});

// Cancel button
cancelBtn.addEventListener("click", function () {
    resetForm();
});

// ============================================
// 7. HELPER FUNCTIONS
// ============================================

// Escape HTML to prevent XSS
function escapeHtml(text) {
    const div = document.createElement("div");
    div.textContent = text;
    return div.innerHTML;
}

// Escape string for inline JS attributes
function escapeJs(text) {
    return text
        .replace(/\\/g, "\\\\")
        .replace(/'/g, "\\'")
        .replace(/"/g, '\\"')
        .replace(/\n/g, "\\n")
        .replace(/\r/g, "\\r");
}

// Toast notification
function showToast(message, type) {
    toast.textContent = message;
    toast.className = "toast " + type + " show";

    setTimeout(function () {
        toast.className = "toast";
    }, 3000);
}

// ============================================
// 8. INITIAL LOAD
// ============================================
document.addEventListener("DOMContentLoaded", function () {
    loadNotes();
});
