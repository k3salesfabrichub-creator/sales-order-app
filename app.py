from PIL import Image
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import openpyxl
import os
from datetime import datetime

st.set_page_config(page_title="Sales Order", layout="centered")

st.title("📦 Sales Order Generator")

# ================= ORDER NUMBER =================
def get_order_number(prefix):
    file = "order_series.xlsx"

    if not os.path.exists(file):
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.append(["Type","Last Number"])
        sheet.append(["SO",14500])
        sheet.append(["SM",14000])
        sheet.append(["SOT",50000])
        sheet.append(["VFSO",10000])
        sheet.append(["VFSM",10000])
        wb.save(file)

    wb = openpyxl.load_workbook(file)
    sheet = wb.active

    for row in sheet.iter_rows(min_row=2):
        if row[0].value == prefix:
            row[1].value += 1
            num = row[1].value
            wb.save(file)
            return f"{prefix} {num}"

# ================= FORM =================
name = st.text_input("Party Name")
phone = st.text_input("Phone")
address = st.text_input("Address")
salesman = st.text_input("Salesman")

order_type = st.selectbox("Order Type", ["SO","SM","SOT","VFSO","VFSM"])
manual_no = st.text_input("Manual Order No (optional)")

date = datetime.now().strftime("%d/%m/%y")

# ================= FABRIC =================
st.subheader("Fabric Setup")

fabric1 = st.text_input("Fabric 1")
rate1 = st.number_input("Rate 1", min_value=0)

fabric2 = st.text_input("Fabric 2")
rate2 = st.number_input("Rate 2", min_value=0)

fabric3 = st.text_input("Fabric 3")
rate3 = st.number_input("Rate 3", min_value=0)

fabric_options = [f for f in [fabric1,fabric2,fabric3] if f]

# ================= IMAGE UPLOAD =================
uploaded_files = st.file_uploader(
    "Upload Designs",
    type=["jpg","jpeg","png"],
    accept_multiple_files=True
)

design_data = []

if uploaded_files:
    for i, file in enumerate(uploaded_files):

        st.markdown(f"### Design {i+1}")
        st.image(file, width=150)

        pcs = st.number_input(f"PCS {i+1}", min_value=1, key=f"pcs{i}")
        fabric = st.selectbox(f"Fabric {i+1}", fabric_options, key=f"fab{i}")
        cut = st.number_input(f"Cut {i+1}", min_value=0.0, key=f"cut{i}")

        mtr = pcs * cut if cut > 0 else pcs

        design_data.append({
            "file": file,
            "pcs": pcs,
            "fabric": fabric,
            "cut": cut,
            "mtr": mtr
        })

# ================= NOTES =================
notes = st.text_area("IMPORTANT NOTES")

# ================= PDF =================
def create_pdf(data, design_data):

    file = f"{data['order_no']}.pdf"
    c = canvas.Canvas(file, pagesize=A4)

    width, height = A4

    def draw_header():
        y = height - 40
        c.drawString(40, y, f"Date: {data['date']}")
        c.drawString(40, y-20, f"ORDER NO: {data['order_no']}")
        c.drawString(40, y-40, f"Salesman: {data['salesman']}")
        return y-70

    y = draw_header()

    positions = [
        (60, y),
        (300, y),
        (60, y-250),
        (300, y-250)
    ]

    last_y = y

    for i, d in enumerate(design_data):

        if i > 0 and i % 4 == 0:
            c.showPage()
            y = draw_header()

        x, base_y = positions[i % 4]
        last_y = base_y

        # SERIAL
        c.drawString(x, base_y+12, f"{i+1}.")

        # IMAGE (NO STRETCH)
        img = Image.open(d["file"])
        img.thumbnail((180,180))
        img_reader = ImageReader(img)

        c.drawImage(img_reader, x, base_y-180)

        # TEXT
        c.drawString(x, base_y-200, f"{d['fabric']}")
        c.drawString(x, base_y-215, f"PCS: {d['pcs']} | Cut: {d['cut']}")
        c.drawString(x, base_y-230, f"MTR: {d['mtr']}")

    # ===== NOTES SAME PAGE =====
    if notes:
        note_y = last_y - 260

        if note_y < 60:
            c.showPage()
            c.setFillColorRGB(1,0,0)
            c.setFont("Helvetica-Bold", 12)
            c.drawString(50,750,"IMPORTANT NOTES")
            c.drawString(50,730,notes)
        else:
            c.setFillColorRGB(1,0,0)
            c.setFont("Helvetica-Bold", 12)
            c.drawString(50, note_y, "IMPORTANT NOTES")
            c.drawString(50, note_y-20, notes)

    c.save()
    return file

# ================= EXCEL =================
def save_excel(data, design_data):

    file = "orders.xlsx"

    if not os.path.exists(file):
        wb = openpyxl.Workbook()
        sheet = wb.active

        sheet.append([
            "Order No","Party Name","Address","Date","Salesman",
            "Fabric 1","MTR 1","Rate 1",
            "Fabric 2","MTR 2","Rate 2",
            "Fabric 3","MTR 3","Rate 3",
            "Total MTR"
        ])
    else:
        wb = openpyxl.load_workbook(file)
        sheet = wb.active

    summary = {}

    for d in design_data:
        f = d["fabric"]
        summary[f] = summary.get(f,0) + d["mtr"]

    items = list(summary.items())[:3]

    while len(items) < 3:
        items.append(("",0))

    def get_rate(f):
        if f == fabric1: return rate1
        if f == fabric2: return rate2
        if f == fabric3: return rate3
        return 0

    total = sum([i[1] for i in items])

    sheet.append([
        data['order_no'],
        data['name'],
        data['address'],
        data['date'],
        data['salesman'],

        items[0][0],items[0][1],get_rate(items[0][0]),
        items[1][0],items[1][1],get_rate(items[1][0]),
        items[2][0],items[2][1],get_rate(items[2][0]),

        total
    ])

    wb.save(file)

# ================= BUTTON =================
if st.button("🚀 Generate Order"):

    if manual_no:
        order_no = manual_no
    else:
        order_no = get_order_number(order_type)

    data = {
        "order_no": order_no,
        "name": name,
        "address": address,
        "salesman": salesman,
        "date": date
    }

    pdf = create_pdf(data, design_data)
    save_excel(data, design_data)

    st.success("Order Created")

    with open(pdf,"rb") as f:
        st.download_button("Download PDF", f, file_name=pdf)