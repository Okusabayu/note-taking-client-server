# Note Taking Client-Server Application

A simple Note Taking application developed as the Final Project for the **Client Server Programming** course.

## 📌 Project Overview

This project implements a basic note management system inspired by Google Keep using a Client-Server architecture.

The application allows users to:

* Create notes
* View notes
* Update notes
* Delete notes

Communication between client and server is performed through a REST API built with Flask, while note data is stored using SQLite.

---

## 🏗 Architecture

```text
Browser (Client)
        │
        │ HTTP Request/Response
        ▼
Flask REST API (Server)
        │
        │ SQL Query
        ▼
SQLite Database
```

---

## 🛠 Technologies Used

### Backend

* Python 3
* Flask
* SQLite

### Frontend

* HTML
* CSS
* JavaScript

### API

* RESTful API
* JSON Data Format

---

## 📂 Project Structure

```text
note-taking-client-server/
│
├── app.py
├── init_db.py
├── database.db
├── requirements.txt
│
├── templates/
│   └── index.html
│
├── static/
│   ├── style.css
│   └── script.js
│
└── venv/
```

---

## 🚀 Features

### Create Note

Add a new note to the database.

### Read Notes

Display all saved notes.

### Update Note

Modify existing note content.

### Delete Note

Remove notes from the database.

---

## 🔌 REST API Endpoints

| Method | Endpoint    | Description        |
| ------ | ----------- | ------------------ |
| GET    | /notes      | Retrieve all notes |
| POST   | /notes      | Create a new note  |
| PUT    | /notes/<id> | Update a note      |
| DELETE | /notes/<id> | Delete a note      |

---

## 📦 Installation

### 1. Clone Repository

```bash
git clone https://github.com/USERNAME/note-taking-client-server.git
cd note-taking-client-server
```

### 2. Create Virtual Environment

```bash
python -m venv venv
```

### 3. Activate Virtual Environment

Windows:

```bash
venv\Scripts\activate
```

Linux / macOS:

```bash
source venv/bin/activate
```

### 4. Install Dependencies

```bash
pip install -r requirements.txt
```

### 5. Initialize Database

```bash
python init_db.py
```

### 6. Run Application

```bash
python app.py
```

---

## 🌐 Access Application

Open your browser and visit:

```text
http://127.0.0.1:5000
```

---

## 🧪 Testing

The REST API can be tested using:

* Postman
* cURL
* Browser (for GET requests)

Example endpoint:

```text
GET http://127.0.0.1:5000/notes
```

---

## 📖 Course Information

**Course:** Client Server Programming

**Final Project:** Analysis and Simple Implementation of a Note Taking Application

**Study Case:** Google Keep

---

## 👨‍💻 Author

Name: YOUR NAME

Student ID: YOUR NIM

Informatics Study Program
