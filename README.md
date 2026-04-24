# Hospital-Appointment-Management-Dashboard
Built a web-based Hospital Appointment Management System using Python Flask, Excel, Pandas, OpenPyXL, and Chart.js. It enables appointment booking, tracks records, and uses analytics dashboards to identify causes of appointment delays by comparing fixed 15-minute slots with actual patient consultation needs.

# 🏥 Hospital Appointment Management System

A web-based hospital appointment booking and analytics system built with **Python Flask**, **Microsoft Excel** as the database, and **Chart.js** for visualisations.

> Built as a final project for the Spreadsheets Course

---

## 📸 Screenshots

### Dashboard
<img width="1919" height="1079" alt="image" src="https://github.com/user-attachments/assets/06a182bc-bb6b-4a0c-9d50-e1849b97e5bf" />

### Book Appointment
<img width="1919" height="871" alt="image" src="https://github.com/user-attachments/assets/d718d413-f3e1-457a-a8bc-9e1ea252c4bc" />

### Slot Duration Guide
<img width="1919" height="1079" alt="image" src="https://github.com/user-attachments/assets/df1155a8-11f2-480d-8d4f-275bf13e6007" />

### Appointment Records
<img width="1919" height="864" alt="image" src="https://github.com/user-attachments/assets/2812dfd1-3616-46f5-9fdb-551a8a3993e7" />

---

## 📌 Problem Statement

Most hospitals assign every patient the same **15-minute appointment slot**, regardless of how complex their case is. A routine follow-up fits in 15 minutes. A chronic case does not.

This mismatch causes **cascading delays** — one overrun pushes every appointment after it further behind schedule.

This project addresses that by:
1. Building a functional appointment booking system
2. Visualising doctor performance and patient flow on a live dashboard
3. Analysing how much time each patient type *actually* needs vs what they are allocated

---

## ✨ Features

- **Appointment Booking** — Select a doctor, date, patient type, and available time slot. Booked slots are greyed out in real time so double-booking is prevented
- **Live Dashboard** — 9 charts covering doctor workload, completion rates, patient type breakdown, department load, no-show patterns, and booking trends
- **Slot Duration Guide** — Compares the current 15-min flat allocation against recommended durations per patient type, with delay rate analysis
- **Excel Backend** — All data is stored in `hospital_data.xlsx`. The file is formatted, colour-coded, and usable standalone in Microsoft Excel
- **Export** — Download the live Excel file at any time from the top bar

---

## 🛠 Tech Stack

| Layer | Technology |
|---|---|
| Frontend | HTML, CSS, JavaScript, Chart.js |
| Backend | Python, Flask |
| Database | Microsoft Excel (`.xlsx`) |
| Excel handling | `pandas`, `openpyxl` |

---

## 📁 Project Structure
hospital_system/
├── app.py                  # Flask backend — all routes and Excel logic
├── requirements.txt        # Python dependencies
├── hospital_data.xlsx      # Auto-created on first run (100 seeded rows)
└── templates/
└── index.html          # Complete frontend — dashboard, booking, records

---

## ⚙️ Setup & Installation

**1. Clone the repository**
```bash
git clone https://github.com/your-username/hospital-appointment-system.git
cd hospital-appointment-system
```

**2. Install dependencies**
```bash
pip install -r requirements.txt
```

**3. Run the app**
```bash
python app.py
```

**4. Open in browser**
 * Running on http://127.0.0.1:5000
   
The Excel file `hospital_data.xlsx` is created automatically on first run with 100 pre-seeded appointment records.

---

## 📊 Excel File Structure

The workbook contains two sheets:

**Sheet 1 — `Appointments`**

| Column | Description |
|---|---|
| Appointment_ID | Auto-generated (APT1001, APT1002, …) |
| Date | Appointment date |
| Patient_Name | Name of the patient |
| Doctor | Assigned doctor |
| Department | Doctor's department |
| Patient_Type | New Patient / Follow-up / Chronic Case / Emergency / Consultation |
| Time_Slot | Booked time (e.g. 10:00) |
| Allocated_Min | Slot duration assigned (currently 15 min flat) |
| Actual_Min | Actual consultation duration |
| Delay_Min | Actual − Allocated (positive = overrun) |
| Status | Confirmed / Completed / Cancelled / No-Show |

**Sheet 2 — `Slot_Guide`**

Reference table showing recommended slot durations per patient type, delay rates, and required actions.

---

## 📈 Dashboard Charts

| Chart | Type | What it shows |
|---|---|---|
| Doctor Workload | Horizontal bar | Appointments completed vs confirmed per doctor |
| Doctor Completion Rate | Bar | % of appointments completed, colour-coded |
| Appointment Status Mix | Doughnut | Confirmed / Completed / Cancelled / No-Show split |
| Patient Type Breakdown | Horizontal bar | Count of appointments per patient category |
| Department Load | Polar area | Appointment volume by department |
| Patient Type per Doctor | Stacked bar | Case complexity mix for each doctor |
| Popular Time Slots | Bar | Top 8 most booked appointment times |
| Daily Booking Trend | Line | Bookings over the last 14 days |
| No-Show & Cancellations | Grouped bar | Attendance issues broken down by doctor |

---

## 🕐 Slot Duration Analysis

The **Slot Guide** page is the research contribution of this project. It answers: *how long does each patient type actually need?*

| Patient Type | Current Slot | Recommended | Avg Actual | Delay Rate |
|---|---|---|---|---|
| New Patient | 15 min | 30 min | ~32 min | 77% |
| Follow-up | 15 min | 15 min | ~13 min | 27% |
| Chronic Case | 15 min | 45 min | ~45 min | 100% |
| Emergency | 15 min | 60 min | ~55 min | 100% |
| Consultation | 15 min | 20 min | ~19 min | 43% |

Only **Follow-up** patients fit within the 15-minute slot. Chronic Cases and Emergencies exceed it every single time.

---

## 🔌 API Endpoints

| Endpoint | Method | Description |
|---|---|---|
| `/` | GET | Serves the web app |
| `/api/meta` | GET | Returns doctors, patient types, time slots |
| `/api/dashboard` | GET | Returns all computed metrics for charts |
| `/api/appointments` | GET | Returns recent appointment records |
| `/api/book` | POST | Books an appointment, writes to Excel |
| `/api/slots_taken` | GET | Returns booked slots for a doctor + date |
| `/api/download` | GET | Downloads the Excel file |

---

## 🧑‍💻 How Booking Works

1. User selects a **doctor** and **date**
2. Frontend calls `/api/slots_taken` → Flask reads the Excel file → returns taken slots
3. Taken slots are greyed out in the UI
4. User selects an available slot and confirms
5. Flask writes a new row to `hospital_data.xlsx` via `openpyxl`
6. Appointment ID is returned and shown in a confirmation modal

---

## 📋 Requirements
flask==3.0.3
pandas==2.2.2
openpyxl==3.1.4
Python 3.8 or higher recommended.
