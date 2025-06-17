

import openai
import os
import json
from docx import Document
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime

# 🔐 Set your OpenAI API key (safe for local testing)
import streamlit as st
openai.api_key = st.secrets["openai_api_key"]
# 📐 Config
TASK_ROW_START = 8
TASK_ROW_GAP = 2
TASK_COLS = {'task': 4, 'assigned_to': 5, 'start': 6, 'end': 7}
START_DATE_CELL = 'F3'

# 📄 Step 1: Read DOCX
def read_scope_docx(docx_path):
    doc = Document(docx_path)
    return "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])

# 🤖 Step 2: Use GPT-3.5 to extract structured task data
def extract_tasks_with_gpt(scope_text):
    prompt = f"""
You are an AI assistant. Based on the project scope below, extract:
1. Project Start Date (if missing, assume today)
2. Extract all tasks with:
   - Task description
   - Person assigned
   - Start date
   - End date

Format your response strictly as JSON:
{{
  "project_start_date": "YYYY-MM-DD",
  "tasks": [
    {{
      "task": "task name",
      "assigned_to": "name",
      "start": "YYYY-MM-DD",
      "end": "YYYY-MM-DD"
    }}
  ]
}}

Project Scope:
\"\"\"
{scope_text}
\"\"\"
"""
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2
    )

    content = response.choices[0].message.content.strip()
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        print("⚠️ GPT returned invalid JSON. Raw response:")
        print(content)
        raise

# 🧪 Helper: Copy row styles for formatting
def copy_row_style(source_row, target_row, ws):
    for col in range(2, 9):  # Columns B to H
        source = ws.cell(row=source_row, column=col)
        target = ws.cell(row=target_row, column=col)

        if source.has_style:
            target.font = source.font.copy()
            target.border = source.border.copy()
            target.fill = source.fill.copy()
            target.number_format = source.number_format
            target.alignment = source.alignment.copy()

# 📊 Step 3: Fill Gantt Excel
def fill_gantt_excel(template_path, output_path, start_date, tasks):
    wb = load_workbook(template_path)
    ws = wb.active

    # Set project start date in cell F3
    ws[START_DATE_CELL] = datetime.strptime(start_date, "%Y-%m-%d").date()

    for i, task in enumerate(tasks):
        row = TASK_ROW_START + i * TASK_ROW_GAP

        # Copy formatting if beyond default 10-task range
        if row > 26:
            copy_row_style(26, row, ws)

        ws.cell(row=row, column=1).value = i + 1  # Column A: No.
        ws.cell(row=row, column=2).value = 0      # Column B: Progress
        ws.cell(row=row, column=TASK_COLS['task']).value = task["task"]
        ws.cell(row=row, column=TASK_COLS['assigned_to']).value = task["assigned_to"]

        start_date_obj = datetime.strptime(task["start"], "%Y-%m-%d").date()
        end_date_obj = datetime.strptime(task["end"], "%Y-%m-%d").date()

        ws.cell(row=row, column=TASK_COLS['start']).value = start_date_obj
        ws.cell(row=row, column=TASK_COLS['end']).value = end_date_obj

        # Column H (Days) = G - F
        col_start = get_column_letter(TASK_COLS['start'])  # F
        col_end = get_column_letter(TASK_COLS['end'])      # G
        ws.cell(row=row, column=8).value = f"={col_end}{row}-{col_start}{row}"

    wb.save(output_path)
    print(f"✅ Excel Gantt chart saved to: {output_path}")

# 🧠 Step 4: Agent Orchestrator
def run_agent(docx_path, template_path):
    print("📄 Reading project scope...")
    scope_text = read_scope_docx(docx_path)

    print("🤖 Extracting structured tasks with GPT-3.5...")
    data = extract_tasks_with_gpt(scope_text)

    output_path = f"filled_gantt_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    fill_gantt_excel(template_path, output_path, data["project_start_date"], data["tasks"])

# ▶️ Run
run_agent("project_scope.docx", "Gantt Chart Template.xlsx")
