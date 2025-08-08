import openai
import os
import json
import re
from docx import Document
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime
import streamlit as st
import PyPDF2

# Optional: For more robust PDF table extraction
try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

# 🔐 API key
openai.api_key = st.secrets["openai_api_key"]

# 📀 Excel config
TASK_ROW_START = 8
TASK_ROW_GAP = 2
TASK_COLS = {'task': 4, 'assigned_to': 5, 'start': 6, 'end': 7}
START_DATE_CELL = 'F3'  # Project Start Date
PROJECT_NAME_CELL = 'F2'  # <-- NEW: write detected project name here

# -------------------- DOC INGEST --------------------
def extract_tables_from_docx(doc):
    """Extract all tables from a docx Document as CSV-style text"""
    tables_text = []
    for t in doc.tables:
        rows = []
        for row in t.rows:
            cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
            rows.append(", ".join(cells))
        tables_text.append("\n".join(rows))
    return "\n\n".join(tables_text)

def read_scope_docx(docx_path):
    doc = Document(docx_path)
    paragraphs = "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
    tables = extract_tables_from_docx(doc)
    combined = paragraphs
    if tables:
        combined += "\n\n[TABLES]\n" + tables
    return combined

def extract_tables_from_pdf(pdf_path):
    """Try extracting tables from PDF using pdfplumber, fallback to none."""
    if HAS_PDFPLUMBER:
        tables = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_tables = page.extract_tables()
                for table in page_tables:
                    rows = []
                    for row in table:
                        if row:
                            cells = [str(cell).strip().replace("\n", " ") if cell else "" for cell in row]
                            rows.append(", ".join(cells))
                    if rows:
                        tables.append("\n".join(rows))
        return "\n\n".join(tables)
    else:
        return ""

def read_scope_pdf(pdf_path):
    text = ""
    with open(pdf_path, "rb") as f:
        pdf = PyPDF2.PdfReader(f)
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text += page_text + "\n"
    tables = extract_tables_from_pdf(pdf_path)
    combined = text
    if tables:
        combined += "\n\n[TABLES]\n" + tables
    return combined

# -------------------- HELPERS --------------------
def extract_json_from_response(content):
    match = re.search(r'\{[\s\S]*\}', content)
    if match:
        return match.group(0)
    return content

def parse_task_number(number):
    """Convert task number to numeric (int/float) if possible, else leave as str."""
    try:
        if isinstance(number, (int, float)):
            return number
        s = str(number)
        if "." in s:
            return float(s)
        return int(s)
    except Exception:
        return number

def parse_date(date_str):
    """Try parsing a date string to a Python date in YYYY-MM-DD. Return None if invalid."""
    if not date_str or date_str == "null":
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d %b %Y", "%d %B %Y"):
        try:
            return datetime.strptime(date_str, fmt).date()
        except Ex
