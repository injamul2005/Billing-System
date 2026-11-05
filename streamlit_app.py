
import streamlit as st
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from io import BytesIO

# =====================================
# PAGE CONFIGURATION
# =====================================

st.set_page_config(
    page_title="MIRA MOBILE SHOP - Billing System",
    page_icon="🛒",
    layout="wide"
)

# =====================================
# SESSION STATE
# =====================================

if "items" not in st.session_state:
    st.session_state.items = []

if "generated_bill" not in st.session_state:
    st.session_state.generated_bill = ""

if "bill_number" not in st.session_state:
    st.session_state.bill_number = ""

if "bill_date" not in st.session_state:
    st.session_state.bill_date = ""


def clear_bill():
    st.session_state.items = []
    st.session_state.generated_bill = ""
    st.session_state.bill_number = ""
    st.session_state.bill_date = ""
    st.session_state.discount_rate = 0.0


# =====================================
# BILL CALCULATION
# =====================================

def calculate_bill():

    subtotal = sum(item["total"] for item in st.session_state.items)

    discount_rate = st.session_state.get("discount_rate", 0.0)

    discount = subtotal * discount_rate / 100
    after_discount = subtotal - discount

    gst = after_discount * 18 / 100
    grand_total = after_discount + gst

    return subtotal, discount, after_discount, gst, grand_total


# =====================================
# GENERATE BILL
# =====================================

def generate_bill():

    if not st.session_state.items:
        st.warning("Please add at least one item.")
        return

    subtotal, discount, after_discount, gst, grand_total = calculate_bill()

    bill_number = datetime.now().strftime("%Y%m%d%H%M%S")
    date = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

    bill = ""

    bill += "============================================\n"
    bill += "              MIRA MOBILE SHOP\n"
    bill += "               BILLING SYSTEM\n"
    bill += "============================================\n\n"

    bill += f"Bill No : {bill_number}\n"
    bill += f"Date    : {date}\n\n"

    bill += "--------------------------------------------\n"
    bill += f"{'Item':<18}{'Qty':<6}{'Price':<10}{'Total':<10}\n"
    bill += "--------------------------------------------\n"

    for item in st.session_state.items:

        bill += (
            f"{item['name'][:17]:<18}"
            f"{item['quantity']:<6}"
            f"{item['price']:<10.2f}"
            f"{item['total']:.2f}\n"
        )

    bill += "--------------------------------------------\n"

    bill += f"Subtotal          : Rs. {subtotal:.2f}\n"
    bill += f"Discount ({st.session_state.discount_rate}%): -Rs. {discount:.2f}\n"
    bill += f"After Discount    : Rs. {after_discount:.2f}\n"
    bill += f"GST (18%)         : Rs. {gst:.2f}\n"
    bill += f"Grand Total       : Rs. {grand_total:.2f}\n"

    bill += "\n============================================\n"
    bill += "          Thank You! Visit Again\n"
    bill += "============================================\n"

    st.session_state.generated_bill = bill
    st.session_state.bill_number = bill_number
    st.session_state.bill_date = date


# =====================================
# EXCEL GENERATION
# =====================================

def create_excel():

    subtotal, discount, after_discount, gst, grand_total = calculate_bill()

    wb = Workbook()
    ws = wb.active
    ws.title = "Billing Details"

    # Shop heading
    ws.merge_cells("A1:D1")
    ws["A1"] = "MIRA MOBILE SHOP"
    ws["A1"].font = Font(size=18, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="1F4E78")
    ws["A1"].alignment = Alignment(horizontal="center")

    ws.merge_cells("A2:D2")
    ws["A2"] = "BILLING SYSTEM"
    ws["A2"].font = Font(size=14, bold=True)
    ws["A2"].alignment = Alignment(horizontal="center")

    # Bill details
    ws["A4"] = "Bill Number"
    ws["B4"] = st.session_state.bill_number

    ws["A5"] = "Date"
    ws["B5"] = st.session_state.bill_date

    # Table headings
    headers = ["Item Name", "Quantity", "Price", "Total"]

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=7, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center")

    # Products
    row = 8

    for item in st.session_state.items:

        ws.cell(row=row, column=1, value=item["name"])
        ws.cell(row=row, column=2, value=item["quantity"])
        ws.cell(row=row, column=3, value=item["price"])
        ws.cell(row=row, column=4, value=item["total"])

        ws.cell(row=row, column=3).number_format = '₹#,##0.00'
        ws.cell(row=row, column=4).number_format = '₹#,##0.00'

        row += 1

    # Summary
    row += 1

    summary = [
        ("Subtotal", subtotal),
        (f"Discount ({st.session_state.discount_rate}%)", -discount),
        ("After Discount", after_discount),
        ("GST (18%)", gst),
        ("Grand Total", grand_total)
    ]

    for label, amount in summary:

        ws.cell(row=row, column=3, value=label)
        ws.cell(row=row, column=4, value=amount)

        ws.cell(row=row, column=3).font = Font(bold=True)
        ws.cell(row=row, column=4).font = Font(bold=True)
        ws.cell(row=row, column=4).number_format = '₹#,##0.00'

        if label == "Grand Total":
            for col in [3, 4]:
                ws.cell(row=row, column=col).fill = PatternFill(
                    "solid", fgColor="70AD47"
                )
                ws.cell(row=row, column=col).font = Font(
                    bold=True, color="FFFFFF"
                )

        row += 1

    # Column widths
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 20

    # Borders
    thin = Side(style="thin", color="CCCCCC")

    for row_cells in ws.iter_rows(
        min_row=7,
        max_row=row - 1,
        min_col=1,
        max_col=4
    ):
        for cell in row_cells:
            cell.border = Border(
                left=thin,
                right=thin,
                top=thin,
                bottom=thin
            )

    # Return Excel file in memory
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue()


# =====================================
# MAIN USER INTERFACE
# =====================================

st.title("🛒 MIRA MOBILE SHOP")
st.subheader("Billing Management System")

st.caption("Professional Product Billing | Discount | GST | Excel Export")

st.divider()

# =====================================
# ADD PRODUCT
# =====================================

st.header("📦 Add Product")

with st.form("product_form", clear_on_submit=True):

    col1, col2, col3 = st.columns(3)

    with col1:
        name = st.text_input("Product Name")

    with col2:
        quantity = st.number_input(
            "Quantity",
            min_value=1,
            value=1,
            step=1
        )

    with col3:
        price = st.number_input(
            "Price (₹)",
            min_value=0.0,
            value=0.0,
            step=10.0
        )

    add_button = st.form_submit_button(
        "➕ Add Product",
        use_container_width=True
    )

    if add_button:

        if not name.strip():
            st.error("Please enter a product name.")

        elif price <= 0:
            st.error("Please enter a price greater than zero.")

        else:

            total = quantity * price

            st.session_state.items.append({
                "name": name.strip(),
                "quantity": quantity,
                "price": price,
                "total": total
            })

            st.session_state.generated_bill = ""
            st.success(f"{name} added successfully!")

# =====================================
# PRODUCT TABLE
# =====================================

st.divider()
st.header("🧾 Added Products")

if st.session_state.items:

    st.dataframe(
        st.session_state.items,
        use_container_width=True,
        hide_index=True,
        column_config={
            "name": "Product Name",
            "quantity": "Quantity",
            "price": st.column_config.NumberColumn(
                "Price (₹)", format="₹%.2f"
            ),
            "total": st.column_config.NumberColumn(
                "Total (₹)", format="₹%.2f"
            )
        }
    )

else:
    st.info("No products added yet.")

# =====================================
# DISCOUNT
# =====================================

st.divider()
st.header("🏷️ Discount Settings")

discount_rate = st.number_input(
    "Discount Percentage (%)",
    min_value=0.0,
    max_value=100.0,
    value=0.0,
    step=1.0,
    key="discount_rate"
)

# =====================================
# BILL SUMMARY
# =====================================

subtotal, discount, after_discount, gst, grand_total = calculate_bill()

st.divider()
st.header("💰 Bill Summary")

col1, col2, col3, col4, col5 = st.columns(5)

col1.metric("Subtotal", f"₹{subtotal:,.2f}")
col2.metric("Discount", f"-₹{discount:,.2f}")
col3.metric("After Discount", f"₹{after_discount:,.2f}")
col4.metric("GST (18%)", f"₹{gst:,.2f}")
col5.metric("Grand Total", f"₹{grand_total:,.2f}")

# =====================================
# BILL ACTIONS
# =====================================

st.divider()
st.header("⚙️ Bill Actions")

col1, col2 = st.columns(2)

with col1:
    if st.button("🧾 Generate Bill", use_container_width=True):
        generate_bill()

with col2:
    st.button(
        "🗑️ Clear All",
        use_container_width=True,
        on_click=clear_bill
    )

# =====================================
# DISPLAY GENERATED BILL
# =====================================

if st.session_state.generated_bill:

    st.divider()
    st.header("📄 Generated Invoice")

    st.code(st.session_state.generated_bill, language="text")

    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            "⬇️ Download TXT Bill",
            data=st.session_state.generated_bill,
            file_name=f"Bill_{st.session_state.bill_number}.txt",
            mime="text/plain",
            use_container_width=True
        )

    with col2:
        excel_data = create_excel()

        st.download_button(
            "📊 Download Excel Bill",
            data=excel_data,
            file_name=f"Bill_{st.session_state.bill_number}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# =====================================
# FOOTER
# =====================================

st.divider()

st.caption("Developed by Injamul Haque | Python & Streamlit Project")
