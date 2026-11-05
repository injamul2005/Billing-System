import streamlit as st
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# =====================================
# MIRA MOBILE SHOP - STREAMLIT BILLING
# Converted from the original Tkinter project
# =====================================

st.set_page_config(
    page_title="MIRA MOBILE SHOP - Billing System",
    page_icon="🛒",
    layout="wide"
)

# Use "cart", not session_state.items (items conflicts with a dict method).
if "cart" not in st.session_state:
    st.session_state["cart"] = []

if "generated_bill" not in st.session_state:
    st.session_state["generated_bill"] = ""

if "bill_number" not in st.session_state:
    st.session_state["bill_number"] = ""

if "bill_date" not in st.session_state:
    st.session_state["bill_date"] = ""

if "discount_rate" not in st.session_state:
    st.session_state["discount_rate"] = 0.0


def get_totals():
    subtotal = sum(product["total"] for product in st.session_state["cart"])
    rate = float(st.session_state["discount_rate"])
    discount = subtotal * rate / 100
    after_discount = subtotal - discount
    gst = after_discount * 18 / 100
    grand_total = after_discount + gst
    return subtotal, discount, after_discount, gst, grand_total


def make_bill():
    if not st.session_state["cart"]:
        st.warning("No items added. Please add at least one product.")
        return

    subtotal, discount, after_discount, gst, grand_total = get_totals()
    bill_number = datetime.now().strftime("%Y%m%d%H%M%S")
    bill_date = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

    lines = [
        "============================================",
        "              MIRA MOBILE SHOP",
        "               BILLING SYSTEM",
        "============================================",
        f"Bill No : {bill_number}",
        f"Date    : {bill_date}",
        "--------------------------------------------",
        f"{'Item':<20}{'Qty':<8}{'Price':<10}{'Total'}",
        "--------------------------------------------",
    ]

    for product in st.session_state["cart"]:
        lines.append(
            f"{product['name'][:19]:<20}"
            f"{product['quantity']:<8}"
            f"{product['price']:<10.2f}"
            f"{product['total']:.2f}"
        )

    lines.extend([
        "--------------------------------------------",
        f"Subtotal          : Rs. {subtotal:.2f}",
        f"Discount ({st.session_state['discount_rate']}%): -Rs. {discount:.2f}",
        f"After Discount    : Rs. {after_discount:.2f}",
        f"GST (18%)         : Rs. {gst:.2f}",
        f"Grand Total       : Rs. {grand_total:.2f}",
        "============================================",
        "          Thank You! Visit Again",
        "============================================",
    ])

    st.session_state["bill_number"] = bill_number
    st.session_state["bill_date"] = bill_date
    st.session_state["generated_bill"] = "\n".join(lines)


def create_excel():
    subtotal, discount, after_discount, gst, grand_total = get_totals()

    wb = Workbook()
    ws = wb.active
    ws.title = "Billing Details"

    ws.merge_cells("A1:D1")
    ws["A1"] = "MIRA MOBILE SHOP"
    ws["A1"].font = Font(size=18, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="1F4E78")
    ws["A1"].alignment = Alignment(horizontal="center")

    ws.merge_cells("A2:D2")
    ws["A2"] = "BILLING SYSTEM"
    ws["A2"].font = Font(size=14, bold=True)
    ws["A2"].alignment = Alignment(horizontal="center")

    ws["A4"] = "Bill Number"
    ws["B4"] = st.session_state["bill_number"]
    ws["A5"] = "Date"
    ws["B5"] = st.session_state["bill_date"]

    headers = ["Item Name", "Quantity", "Price", "Total"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=7, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center")

    row = 8
    for product in st.session_state["cart"]:
        ws.cell(row=row, column=1, value=product["name"])
        ws.cell(row=row, column=2, value=product["quantity"])
        ws.cell(row=row, column=3, value=product["price"])
        ws.cell(row=row, column=4, value=product["total"])
        ws.cell(row=row, column=3).number_format = '₹#,##0.00'
        ws.cell(row=row, column=4).number_format = '₹#,##0.00'
        row += 1

    row += 1
    summary = [
        ("Subtotal", subtotal),
        (f"Discount ({st.session_state['discount_rate']}%)", -discount),
        ("After Discount", after_discount),
        ("GST (18%)", gst),
        ("Grand Total", grand_total),
    ]

    for label, amount in summary:
        ws.cell(row=row, column=3, value=label)
        ws.cell(row=row, column=4, value=amount)
        ws.cell(row=row, column=3).font = Font(bold=True)
        ws.cell(row=row, column=4).font = Font(bold=True)
        ws.cell(row=row, column=4).number_format = '₹#,##0.00'
        if label == "Grand Total":
            for col in (3, 4):
                cell = ws.cell(row=row, column=col)
                cell.fill = PatternFill("solid", fgColor="70AD47")
                cell.font = Font(bold=True, color="FFFFFF")
        row += 1

    for column, width in {"A": 25, "B": 18, "C": 20, "D": 20}.items():
        ws.column_dimensions[column].width = width

    thin = Side(style="thin", color="CCCCCC")
    for row_cells in ws.iter_rows(min_row=7, max_row=row - 1, min_col=1, max_col=4):
        for cell in row_cells:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def clear_all():
    st.session_state["cart"] = []
    st.session_state["generated_bill"] = ""
    st.session_state["bill_number"] = ""
    st.session_state["bill_date"] = ""
    st.session_state["discount_rate"] = 0.0


# =====================================
# USER INTERFACE
# =====================================

st.title("🛒 MIRA MOBILE SHOP")
st.subheader("BILLING SYSTEM")
st.caption("Product billing • Discount • GST • TXT and Excel invoice")
st.divider()

st.header("📦 Add Item")

with st.form("add_product_form", clear_on_submit=True):
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        product_name = st.text_input("Item Name")
    with col2:
        quantity = st.number_input("Quantity", min_value=1, value=1, step=1)
    with col3:
        price = st.number_input("Price (₹)", min_value=0.0, value=0.0, step=1.0)

    submitted = st.form_submit_button("➕ Add Item", use_container_width=True)

if submitted:
    if not product_name.strip():
        st.error("Please enter an item name.")
    elif price < 0:
        st.error("Price cannot be negative.")
    else:
        st.session_state["cart"].append({
            "name": product_name.strip(),
            "quantity": int(quantity),
            "price": float(price),
            "total": int(quantity) * float(price),
        })
        # A changed cart means the previous invoice must be generated again.
        st.session_state["generated_bill"] = ""
        st.success(f"{product_name.strip()} added successfully.")

st.divider()
st.header("🧾 Added Items")

if st.session_state["cart"]:
    st.dataframe(
        st.session_state["cart"],
        use_container_width=True,
        hide_index=True,
        column_config={
            "name": st.column_config.TextColumn("Item Name"),
            "quantity": st.column_config.NumberColumn("Quantity", format="%d"),
            "price": st.column_config.NumberColumn("Price (₹)", format="₹%.2f"),
            "total": st.column_config.NumberColumn("Total (₹)", format="₹%.2f"),
        },
    )

    if st.button("Remove Last Item"):
        st.session_state["cart"].pop()
        st.session_state["generated_bill"] = ""
        st.rerun()
else:
    st.info("No items added yet.")

st.divider()
st.header("🏷️ Discount")

st.number_input(
    "Discount Percentage (%)",
    min_value=0.0,
    max_value=100.0,
    step=1.0,
    key="discount_rate",
)

subtotal, discount, after_discount, gst, grand_total = get_totals()

st.divider()
st.header("💰 Bill Summary")

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Subtotal", f"₹{subtotal:,.2f}")
c2.metric("Discount", f"-₹{discount:,.2f}")
c3.metric("After Discount", f"₹{after_discount:,.2f}")
c4.metric("GST (18%)", f"₹{gst:,.2f}")
c5.metric("Grand Total", f"₹{grand_total:,.2f}")

st.divider()
left, right = st.columns(2)

with left:
    if st.button("🧾 Generate Bill", type="primary", use_container_width=True):
        make_bill()

with right:
    st.button("🗑️ Clear All", on_click=clear_all, use_container_width=True)

if st.session_state["generated_bill"]:
    st.divider()
    st.header("📄 Generated Bill")
    st.code(st.session_state["generated_bill"], language="text")

    file_base = f"Bill_{st.session_state['bill_number']}"
    d1, d2 = st.columns(2)

    with d1:
        st.download_button(
            "⬇️ Download TXT",
            data=st.session_state["generated_bill"],
            file_name=f"{file_base}.txt",
            mime="text/plain",
            use_container_width=True,
        )

    with d2:
        st.download_button(
            "📊 Download Excel",
            data=create_excel(),
            file_name=f"{file_base}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

st.divider()
st.caption("Developed by Injamul Haque | Python & Streamlit")
