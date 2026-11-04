
import tkinter as tk
from tkinter import messagebox, filedialog
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# =====================================
# MIRA MOBILE SHOP - BILLING SYSTEM
# =====================================

items = []


# =====================================
# DISCOUNT VALIDATION
# =====================================

def get_discount():
    try:
        discount_rate = float(discount_entry.get() or 0)

        if discount_rate < 0 or discount_rate > 100:
            raise ValueError

        return discount_rate

    except ValueError:
        messagebox.showerror(
            "Invalid Discount",
            "Enter discount between 0 and 100."
        )
        return None


# =====================================
# CALCULATE TOTAL
# =====================================

def calculate_total():

    subtotal = sum(item["total"] for item in items)

    discount_rate = get_discount()

    if discount_rate is None:
        return

    discount = subtotal * discount_rate / 100
    after_discount = subtotal - discount

    gst = after_discount * 18 / 100
    grand_total = after_discount + gst

    subtotal_label.config(
        text=f"Subtotal: Rs. {subtotal:.2f}"
    )

    discount_label.config(
        text=f"Discount ({discount_rate}%): -Rs. {discount:.2f}"
    )

    gst_label.config(
        text=f"GST (18%): Rs. {gst:.2f}"
    )

    total_label.config(
        text=f"Grand Total: Rs. {grand_total:.2f}"
    )


# =====================================
# ADD ITEM
# =====================================

def add_item():

    name = item_name.get().strip()
    quantity = quantity_entry.get().strip()
    price = price_entry.get().strip()

    if not name or not quantity or not price:
        messagebox.showwarning(
            "Warning",
            "Please enter all item details."
        )
        return

    try:
        quantity = int(quantity)
        price = float(price)

        if quantity <= 0 or price < 0:
            raise ValueError

        total = quantity * price

        item = {
            "name": name,
            "quantity": quantity,
            "price": price,
            "total": total
        }

        items.append(item)

        item_list.insert(
            tk.END,
            f"{name:<20} {quantity:<8} Rs.{price:<10.2f} Rs.{total:.2f}"
        )

        item_name.delete(0, tk.END)
        quantity_entry.delete(0, tk.END)
        price_entry.delete(0, tk.END)

        calculate_total()

    except ValueError:
        messagebox.showerror(
            "Invalid Input",
            "Quantity must be an integer and price must be a number."
        )


# =====================================
# GENERATE BILL
# =====================================

def generate_bill():

    if not items:
        messagebox.showwarning(
            "Warning",
            "No items added."
        )
        return

    discount_rate = get_discount()

    if discount_rate is None:
        return

    bill_number = datetime.now().strftime("%Y%m%d%H%M%S")
    date = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

    subtotal = sum(item["total"] for item in items)

    discount = subtotal * discount_rate / 100
    after_discount = subtotal - discount

    gst = after_discount * 18 / 100
    grand_total = after_discount + gst

    bill = ""

    bill += "============================================\n"
    bill += "              MIRA MOBILE SHOP\n"
    bill += "               BILLING SYSTEM\n"
    bill += "============================================\n"

    bill += f"Bill No : {bill_number}\n"
    bill += f"Date    : {date}\n"

    bill += "--------------------------------------------\n"
    bill += f"{'Item':<20}{'Qty':<8}{'Price':<10}{'Total'}\n"
    bill += "--------------------------------------------\n"

    for item in items:

        bill += (
            f"{item['name']:<20}"
            f"{item['quantity']:<8}"
            f"{item['price']:<10.2f}"
            f"{item['total']:.2f}\n"
        )

    bill += "--------------------------------------------\n"

    bill += f"Subtotal          : Rs. {subtotal:.2f}\n"
    bill += f"Discount ({discount_rate}%): -Rs. {discount:.2f}\n"
    bill += f"After Discount    : Rs. {after_discount:.2f}\n"
    bill += f"GST (18%)         : Rs. {gst:.2f}\n"
    bill += f"Grand Total       : Rs. {grand_total:.2f}\n"

    bill += "============================================\n"
    bill += "          Thank You! Visit Again\n"
    bill += "============================================\n"

    bill_text.delete("1.0", tk.END)
    bill_text.insert(tk.END, bill)


# =====================================
# SAVE BILL AS TEXT FILE
# =====================================

def save_bill():

    bill = bill_text.get("1.0", tk.END).strip()

    if not bill:
        messagebox.showwarning(
            "Warning",
            "Please generate a bill first."
        )
        return

    filename = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text Files", "*.txt")],
        initialfile="Mira_Bill.txt"
    )

    if not filename:
        return

    try:
        with open(filename, "w", encoding="utf-8") as file:
            file.write(bill)

        messagebox.showinfo(
            "Success",
            "Bill saved successfully as TXT!"
        )

    except OSError as error:
        messagebox.showerror("Save Error", str(error))


# =====================================
# SAVE BILL AS EXCEL
# =====================================

def save_excel():

    if not items:
        messagebox.showwarning(
            "Warning",
            "Please add items first."
        )
        return

    discount_rate = get_discount()

    if discount_rate is None:
        return

    subtotal = sum(item["total"] for item in items)

    discount = subtotal * discount_rate / 100
    after_discount = subtotal - discount

    gst = after_discount * 18 / 100
    grand_total = after_discount + gst

    bill_number = datetime.now().strftime("%Y%m%d%H%M%S")
    date = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

    # Create Excel Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Billing Details"

    # Shop Heading
    ws.merge_cells("A1:D1")
    ws["A1"] = "MIRA MOBILE SHOP"
    ws["A1"].font = Font(
        size=18,
        bold=True,
        color="FFFFFF"
    )
    ws["A1"].fill = PatternFill(
        "solid",
        fgColor="1F4E78"
    )
    ws["A1"].alignment = Alignment(
        horizontal="center"
    )

    ws.merge_cells("A2:D2")
    ws["A2"] = "BILLING SYSTEM"
    ws["A2"].font = Font(
        size=14,
        bold=True
    )
    ws["A2"].alignment = Alignment(
        horizontal="center"
    )

    # Bill Information
    ws["A4"] = "Bill Number"
    ws["B4"] = bill_number

    ws["A5"] = "Date"
    ws["B5"] = date

    # Table Heading
    headers = [
        "Item Name",
        "Quantity",
        "Price",
        "Total"
    ]

    for col, header in enumerate(headers, 1):

        cell = ws.cell(row=7, column=col)
        cell.value = header

        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )

        cell.fill = PatternFill(
            "solid",
            fgColor="4472C4"
        )

        cell.alignment = Alignment(
            horizontal="center"
        )

    # Add Items
    row = 8

    for item in items:

        ws.cell(row=row, column=1, value=item["name"])
        ws.cell(row=row, column=2, value=item["quantity"])
        ws.cell(row=row, column=3, value=item["price"])
        ws.cell(row=row, column=4, value=item["total"])

        ws.cell(row=row, column=3).number_format = '₹#,##0.00'
        ws.cell(row=row, column=4).number_format = '₹#,##0.00'

        row += 1

    # Bill Summary
    row += 1

    summary = [
        ("Subtotal", subtotal),
        (f"Discount ({discount_rate}%)", -discount),
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
                    "solid",
                    fgColor="70AD47"
                )

                ws.cell(row=row, column=col).font = Font(
                    bold=True,
                    color="FFFFFF"
                )

        row += 1

    # Column Width
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 20

    # Add Borders
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

    # Choose Save Location
    filename = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        initialfile=f"Bill_{bill_number}.xlsx"
    )

    if not filename:
        return

    try:
        wb.save(filename)

        messagebox.showinfo(
            "Success",
            "Bill saved successfully as Excel file!\n\n" + filename
        )

    except OSError as error:
        messagebox.showerror(
            "Save Error",
            str(error)
        )


# =====================================
# CLEAR ALL DETAILS
# =====================================

def clear_bill():

    items.clear()

    item_list.delete(0, tk.END)
    bill_text.delete("1.0", tk.END)

    item_name.delete(0, tk.END)
    quantity_entry.delete(0, tk.END)
    price_entry.delete(0, tk.END)

    discount_entry.delete(0, tk.END)
    discount_entry.insert(0, "0")

    subtotal_label.config(text="Subtotal: Rs. 0.00")
    discount_label.config(text="Discount (0%): -Rs. 0.00")
    gst_label.config(text="GST (18%): Rs. 0.00")
    total_label.config(text="Grand Total: Rs. 0.00")


# =====================================
# MAIN WINDOW
# =====================================

root = tk.Tk()

root.title("MIRA MOBILE SHOP - Billing System")
root.geometry("850x740")
root.resizable(False, False)


# =====================================
# HEADING
# =====================================

title = tk.Label(
    root,
    text="MIRA MOBILE SHOP",
    font=("Arial", 24, "bold"),
    fg="blue"
)

title.pack(pady=10)

tk.Label(
    root,
    text="BILLING SYSTEM",
    font=("Arial", 15, "bold")
).pack()


# =====================================
# ITEM INPUT SECTION
# =====================================

input_frame = tk.Frame(root)
input_frame.pack(pady=12)

tk.Label(
    input_frame,
    text="Item Name",
    font=("Arial", 11, "bold")
).grid(row=0, column=0, padx=5)

tk.Label(
    input_frame,
    text="Quantity",
    font=("Arial", 11, "bold")
).grid(row=0, column=1, padx=5)

tk.Label(
    input_frame,
    text="Price",
    font=("Arial", 11, "bold")
).grid(row=0, column=2, padx=5)

item_name = tk.Entry(input_frame, width=25)
item_name.grid(row=1, column=0, padx=5, pady=5)

quantity_entry = tk.Entry(input_frame, width=12)
quantity_entry.grid(row=1, column=1, padx=5, pady=5)

price_entry = tk.Entry(input_frame, width=15)
price_entry.grid(row=1, column=2, padx=5, pady=5)

tk.Button(
    input_frame,
    text="Add Item",
    command=add_item,
    width=12,
    bg="lightgreen"
).grid(row=1, column=3, padx=10)


# =====================================
# DISCOUNT SECTION
# =====================================

discount_frame = tk.Frame(root)
discount_frame.pack(pady=5)

tk.Label(
    discount_frame,
    text="Discount Percentage (%):",
    font=("Arial", 11, "bold")
).pack(side=tk.LEFT, padx=5)

discount_entry = tk.Entry(
    discount_frame,
    width=12,
    justify="center"
)

discount_entry.pack(side=tk.LEFT, padx=5)
discount_entry.insert(0, "0")

tk.Button(
    discount_frame,
    text="Apply Discount",
    command=calculate_total,
    bg="lightblue"
).pack(side=tk.LEFT, padx=5)


# =====================================
# ITEM LIST
# =====================================

list_frame = tk.Frame(root)
list_frame.pack(pady=8)

tk.Label(
    list_frame,
    text="Added Items",
    font=("Arial", 13, "bold")
).pack()

item_list = tk.Listbox(
    list_frame,
    width=85,
    height=8,
    font=("Courier New", 10)
)

item_list.pack(pady=5)


# =====================================
# TOTAL SECTION
# =====================================

total_frame = tk.Frame(root)
total_frame.pack(pady=8)

subtotal_label = tk.Label(
    total_frame,
    text="Subtotal: Rs. 0.00",
    font=("Arial", 12, "bold")
)
subtotal_label.pack()

discount_label = tk.Label(
    total_frame,
    text="Discount (0%): -Rs. 0.00",
    font=("Arial", 12, "bold"),
    fg="green"
)
discount_label.pack()

gst_label = tk.Label(
    total_frame,
    text="GST (18%): Rs. 0.00",
    font=("Arial", 12, "bold")
)
gst_label.pack()

total_label = tk.Label(
    total_frame,
    text="Grand Total: Rs. 0.00",
    font=("Arial", 14, "bold"),
    fg="blue"
)
total_label.pack()


# =====================================
# BUTTON SECTION
# =====================================

button_frame = tk.Frame(root)
button_frame.pack(pady=10)

tk.Button(
    button_frame,
    text="Generate Bill",
    command=generate_bill,
    width=15,
    bg="lightgreen"
).grid(row=0, column=0, padx=5)

tk.Button(
    button_frame,
    text="Save TXT",
    command=save_bill,
    width=15,
    bg="lightyellow"
).grid(row=0, column=1, padx=5)

tk.Button(
    button_frame,
    text="Save Excel",
    command=save_excel,
    width=15,
    bg="lightblue"
).grid(row=0, column=2, padx=5)

tk.Button(
    button_frame,
    text="Clear",
    command=clear_bill,
    width=15,
    bg="salmon"
).grid(row=0, column=3, padx=5)


# =====================================
# BILL DISPLAY
# =====================================

tk.Label(
    root,
    text="Generated Bill",
    font=("Arial", 13, "bold")
).pack()

bill_text = tk.Text(
    root,
    width=85,
    height=11,
    font=("Courier New", 10)
)

bill_text.pack(pady=5)


# =====================================
# START APPLICATION
# =====================================

root.mainloop()
