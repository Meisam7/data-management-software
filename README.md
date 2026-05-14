# Customer & Delivery Management Windows App

A Windows desktop application developed with **C# WinForms** for managing customers, delivery records, invoices, and companies/organizations that purchased goods.

This project is designed to help store, view, search, add, edit, and export business records using a local Microsoft Access database.

---

## 📌 Project Description

This application is a Windows-based management system for handling business data related to:

- Customers
- Customer codes
- Delivery notes / Havale records
- Sold goods
- Contracted companies or organizations
- Invoice status
- Provinces and delivery destinations

The application uses a graphical Windows Forms interface and connects to a Microsoft Access database file (`.accdb` or `.mdb`) through `OleDb`.

---

## 🖥️ Technologies Used

- C#
- Windows Forms / WinForms
- .NET Framework
- Microsoft Access Database
- OleDb Connection
- DataGridView
- CSV Export

---

## 🗂️ Main Features

### Customer Management

The application allows the user to manage customer information, including:

- Add new customers
- Edit existing customer information
- Store customer code
- Store customer name
- Display customer list in a table

Example customer fields:

- Customer Code
- Customer Name

---

### Delivery / Havale Management

The application can manage delivery and sales records such as:

- Delivery number
- Customer name
- Contracted company
- Province
- Delivery location
- Delivery amount
- Invoice status

These records are displayed in the main `DataGridView`.

---

### Database Management

The application stores data in a Microsoft Access database.

Supported database formats:

- `.accdb`
- `.mdb`

The last selected database path is saved in application settings, so the program can automatically load the database when it starts again.

---

### Search and Filter

The user can search and filter records based on different fields, such as:

- Customer name
- Customer code
- Delivery number
- Province
- Contracted company
- Delivery destination
- Invoice status
- Delivery amount

This makes it easier to find specific records inside large tables.

---

### Data Export

The application supports exporting the displayed table data to a `.csv` file.

The exported file can be opened with:

- Microsoft Excel
- Google Sheets
- LibreOffice Calc
- Any text editor

The CSV export uses UTF-8 encoding to support Persian text correctly.

---

### Selected Cell Sum

The application can calculate the sum of selected numeric cells in the table.

This is useful when the user wants to quickly calculate totals from selected rows or columns.

Displayed information includes:

- Sum of selected numbers
- Count of selected numeric cells

---

## 🧩 Main Forms

### MainForm

The main form of the application.

Responsibilities:

- Loading database tables
- Showing data in `DataGridView`
- Searching and filtering records
- Exporting data to CSV
- Showing selected numeric sum
- Opening other forms

---

### DBInfoForm

Used for selecting or managing database information.

---

### ClienInfoForm

Used for managing customer information.

Features:

- Add customer
- Edit customer
- Save data to `Table2`
- Refresh customer table after changes

---

### RecordIfoForm

Used for managing main delivery or sales records.

---

## 🗃️ Database Tables

### Table1

Used for main delivery / sales records.

Possible fields:

- Delivery number
- Customer name
- Contracted company
- Province
- Delivery destination
- Delivery amount
- Invoice status

---

### Table2

Used for customer information.

Fields:

| Field Name | Description |
|----------|-------------|
| ردیف | Auto number / row ID |
| کد مشتری | Customer code |
| نام مشتری | Customer name |

---

## ⚙️ How to Run

1. Open the project in Visual Studio.
2. Make sure Microsoft Access Database Engine is installed.
3. Build the project.
4. Run the application.
5. Select or set the database file.
6. Use the menu options to load tables and manage records.

---

## 📤 Exporting Data

To export data:

1. Load the desired table.
2. Click the export option.
3. Choose a file location.
4. Save the file as `.csv`.

The exported file includes only visible columns.

---

## 🔍 Searching Records

The application includes several search boxes and filters.

Examples:

- Search by customer name
- Search by delivery number
- Filter by province
- Search by invoice status
- Search by delivery destination

---

## ✅ Purpose of the Project

The goal of this project is to create a simple and practical Windows application for small business record management.

It helps users manage customers, delivery notes, sold goods, and related companies in one place without needing to work directly inside Microsoft Access.

---

## 🚀 Future Improvements

Possible future improvements:

- Add login system
- Add user roles
- Improve UI design
- Add backup and restore feature
- Add PDF report generation
- Add invoice printing
- Add advanced search
- Add delete confirmation logs
- Add automatic database backup
- Add charts and reports
- Convert the database from Access to SQL Server

---

## 👨‍💻 Developer

Developed with C# Windows Forms.

Project type:

```text
Windows Desktop Application