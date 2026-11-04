# 🛒 MIRA MOBILE SHOP – Billing System

A simple and user-friendly **Billing Management System** developed using Python and Tkinter. This application helps manage product billing, calculate discounts, apply GST, generate invoices, and save billing records in Excel and text formats.

## 🚀 Features

- Add products with item name, quantity, and price.
- Automatically calculate subtotal.
- Apply custom discount percentages.
- Automatically calculate 18% GST.
- Calculate the final grand total.
- Generate professional bills.
- Export bills to Excel (`.xlsx`) format.
- Save bills as text (`.txt`) files.
- Clear all billing details with one click.

## 🛠️ Technologies Used

- **Python** – Main programming language.
- **Tkinter** – Graphical User Interface (GUI).
- **OpenPyXL** – Excel file generation.
- **Datetime** – Date and time management.

## 📦 Installation

Install the required library:

```bash
pip install openpyxl
```

## ▶️ How to Run

1. Clone this repository.
2. Open the project folder in VS Code.
3. Install the required library.
4. Run the following command:

```bash
python billing_system.py
```

## 🧮 Billing Calculation

The system calculates:

- Subtotal = Quantity × Price
- Discount = Subtotal × Discount Percentage / 100
- GST = 18% of the discounted amount
- Grand Total = Discounted Amount + GST

## 📊 Export Options

- Excel Export: Saves billing details in `.xlsx` format.
<img width="787" height="457" alt="image" src="https://github.com/user-attachments/assets/38e8cc37-97da-4666-a322-77d427408991" />

- Text Export: Saves the generated invoice in `.txt` format.
<img width="522" height="467" alt="image" src="https://github.com/user-attachments/assets/09bdf444-d76f-4bab-825c-bd28ee30fb7d" />


## 🎯 Project Objective

The main objective of this project is to develop a simple desktop-based billing application that reduces manual calculation errors and makes the billing process faster and easier.

## 👨‍💻 Developer

**Developed by: Injamul Haque**

BCA Student | Python Developer | Aspiring Software Developer

---

⭐ If you find this project useful, feel free to explore the repository!
<img width="1056" height="962" alt="image" src="https://github.com/user-attachments/assets/72fe3824-dc5b-4189-b3ee-23b7672d3ac5" />
