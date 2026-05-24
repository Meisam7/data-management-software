# Data Management Software

A Windows desktop application built with **C# WinForms** for managing delivery records, customers, contract organizations, bitumen types, invoice status, and dispatch/sales data using a local **Microsoft Access** database.

The application is designed for small business or internal office use where users need a simple desktop tool to manage customer records, government/contract-party organizations, delivery notes, invoice information, remaining balances, and searchable business data.

---

## 📸 Screenshots

Add screenshots to a folder named `screenshots/` in the root of the repository.

### Main Window

![Main Window](screenshots/main-window.png)

### Add / Edit Delivery Record Form

![Record Form](screenshots/record-form.png)

### Database Settings Form

![Database Settings](screenshots/database-settings.png)

### Contract Party Form

![Contract Party Form](screenshots/contract-party-form.png)

### Customer Information Form

![Customer Information Form](screenshots/customer-info-form.png)

---

## 📌 Main Features

### Delivery / Havale Management

The main part of the application is used to manage delivery / havale records stored in `Table1`.

Supported fields include:

- Contract organization / government agency
- Province
- Credit year
- Request ID
- Delivery ID
- Delivery number
- Delivery date
- Customer name
- Bitumen type
- Receiver
- Delivery value
- Delivery amount
- VAT-included unit price
- Sent amount
- Purchased amount
- Final amount
- Remaining delivery balance
- Remaining invoice balance
- Invoice status
- Description

Users can:

- Add new delivery records
- Edit existing records
- Double-click a row ID in the main table to open the record in edit mode
- Update records directly in the Microsoft Access database
- Delete selected `Table1` records using the keyboard `Delete` key with confirmation
- Calculate remaining delivery balance automatically

---

### Add Mode and Edit Mode

The delivery record form is reused for both adding and editing records.

- In **Add Mode**, the **Add** button is enabled and the **Update** button is disabled.
- In **Edit Mode**, the **Update** button is enabled and the **Add** button is disabled.
- When a user double-clicks a `Table1` row in the main `DataGridView`, the selected record is loaded into the record form for editing.

This reduces confusion between creating a new record and updating an existing one.

---

### Customer Management

The application includes a customer management form for storing customer information in `Table2`.

Supported fields:

- Customer code
- Customer name

Users can:

- Add new customers
- Edit existing customers
- Delete customers
- Use saved customer names and customer codes in search filters and record forms

---

### Contract Organization Management

The application includes a form for managing contract organizations / government agencies stored in `Table3`.

Supported field:

- Organization / agency name

These values are used in ComboBoxes for filtering and for selecting the contract organization in delivery records.

---

### Bitumen Type Management

The application can load bitumen names from `Table4`.

Supported field:

- Bitumen name

These values are used in the delivery record form for selecting the bitumen type.

---

## 🔍 Search and Filtering

The main form provides dynamic search/filter controls for the loaded table.

### Table Loading Behavior

When the application starts:

- No table is loaded in the `DataGridView`
- Search/filter group boxes are disabled

When `Table1` is loaded:

- `groupBox1` is enabled
- `groupBox2` is disabled
- Delivery / havale filters become active

When `Table2` is loaded:

- `groupBox1` is disabled
- `groupBox2` is enabled
- Customer filters become active

When `Table3` is loaded:

- Both filter group boxes are disabled

---

### Table1 Filters

When `Table1` is loaded, the user can filter delivery records by:

- Province
- Contract organization / government agency
- Delivery number
- Delivery date range
- Customer name
- Receiver
- Delivery amount
- Invoice status

Some filters use ComboBoxes populated from the database:

- Contract organization values are loaded from `Table3.[نام دستگاه]`
- Customer names are loaded from `Table2.[نام مشتری]`
- Invoice status uses fixed options:
  - `همه`
  - `صادر شده`
  - `صادر نشده`

Selecting `همه` resets the filter and shows all delivery records.

---

### Date Range Filtering

The main form supports date range filtering for delivery records.

The user can filter `Table1` based on:

- Start date only: show records from that date onward
- End date only: show records up to that date
- Both start and end date: show records between the two dates
- No date selected: show all records

Date inputs use masked format:

```text
0000/00/00
```

This keeps delivery dates consistent, for example:

```text
1403/02/15
```

---

### Table2 Filters

When `Table2` is loaded, the user can filter customer records by:

- Customer name
- Customer code

These filters use ComboBoxes populated from:

- `Table2.[نام مشتری]`
- `Table2.[کد مشتری]`

Selecting `همه` resets the filter and shows all customers.

---

## 🧮 Automatic Calculations

### Remaining Delivery Balance

The application calculates the remaining delivery balance using this formula:

```text
مانده حواله = مقدار نهایی - (ارسال + خرید + مقدار تهاتر)
```

This helps users track how much of a delivery record remains after sent, purchased, or offset amounts.

---

### Selected Cell Sum

The main window includes a bottom status label that calculates the sum of selected numeric cells in the `DataGridView`.

This is useful for quickly checking totals without exporting the data.

The `DataGridView` supports cell selection for summing values while still allowing full-row selection for deleting records.

---

## ⌨️ Keyboard and Navigation Features

### Delete Key for Records

When `Table1` is loaded, users can select one or more full rows and press the `Delete` key.

The application shows a confirmation message before deleting:

```text
Are you sure you want to delete these records?
```

Deletion is only active for `Table1`.

If another table is loaded, the application shows a message that deletion is only available for delivery records.

---

### Enter Key Navigation in Record Form

The record form supports custom Enter-key navigation.

After the user finishes one field and presses `Enter`, focus moves to the next logical field based on the custom form order.

This makes data entry faster and easier.

---

## 📤 CSV Export

The application supports exporting the current visible `DataGridView` data to a `.csv` file.

The CSV export uses **UTF-8 with BOM**, which helps Persian text display correctly in tools like Microsoft Excel.

Exported files can be opened with:

- Microsoft Excel
- Google Sheets
- LibreOffice Calc
- Any text editor

---

## 🧩 Forms

### `MainForm`

The main application window.

Responsibilities:

- Load database tables into `DataGridView`
- Display delivery, customer, and contract-party data
- Enable/disable search sections depending on the loaded table
- Search and filter records
- Filter delivery records by date range
- Calculate selected numeric cell totals
- Export data to CSV
- Open add/edit forms
- Handle double-click editing for `Table1`
- Handle delete-key record deletion for `Table1`

---

### `RecordInfoForm` / `Form2`

The delivery record form.

Used for:

- Adding new delivery records
- Editing existing delivery records

Behavior:

- Add mode enables the Add button and disables the Update button
- Edit mode enables the Update button and disables the Add button
- Existing records are loaded from `Table1` using the `ردیف` field
- Delivery date uses a masked date input
- ComboBoxes load values from customer, organization, and bitumen tables
- Enter key navigation moves between fields in the desired order

---

### `ClienInfoForm`

The customer information form.

Used for:

- Adding customers
- Editing customers
- Deleting customers

Data is stored in `Table2`.

---

### `ContractParty`

The contract organization form.

Used for adding government agencies / contract organizations.

Data is stored in `Table3`.

---

### `DBInfoForm`

The database settings form.

Used for selecting the Microsoft Access database file and saving the last selected path in application settings.

---

## 🗃️ Database Structure

The application uses a local Microsoft Access database file:

```text
.accdb / .mdb
```

### `Table1` — Delivery / Havale Records

| Field | Description |
|---|---|
| ردیف | AutoNumber primary key |
| دستگاه طرف قرارداد | Contract organization / government agency |
| استان | Province |
| سال اعتبارات | Credit year |
| شناسه حواله | Delivery ID |
| شناسه درخواست | Request ID |
| معاونت | Department / deputy |
| شماره حواله | Delivery number |
| تاریخ حواله | Delivery date |
| نام مشتری | Customer name |
| نوع قیر | Bitumen type |
| گیرنده | Receiver |
| مبلغ حواله | Delivery value |
| مقدار حواله | Delivery amount |
| فی (با ارزش افزوده) | VAT-included unit price |
| ارسال | Sent amount |
| خرید | Purchased amount |
| مقدار نهایی | Final amount |
| پیمانکار حمل | Transport contractor |
| مبلغ فاکتور | Invoice amount |
| مانده حواله | Remaining delivery balance |
| مانده فاکتور | Remaining invoice balance |
| وضعیت فاکتور | Invoice status |
| توضیحات | Description |

---

### `Table2` — Customers

| Field | Description |
|---|---|
| ردیف | AutoNumber primary key |
| کد مشتری | Customer code |
| نام مشتری | Customer name |

---

### `Table3` — Contract Organizations

| Field | Description |
|---|---|
| ردیف | AutoNumber primary key |
| نام دستگاه | Government agency / contract organization name |

---

### `Table4` — Bitumen Types

| Field | Description |
|---|---|
| ردیف | AutoNumber primary key |
| نام قیر | Bitumen name |

---

## 🖥️ Technologies Used

- C#
- Windows Forms / WinForms
- .NET Framework
- Microsoft Access Database
- OleDb
- DataGridView
- CSV export
- Visual Studio

---

## ⚙️ Requirements

To run the application, the target system needs:

- Windows OS
- .NET Framework compatible with the project
- Microsoft Access Database Engine / OLEDB Provider

For `.accdb` files, the application uses:

```text
Microsoft.ACE.OLEDB.16.0
```

If the application cannot connect to the database, install:

```text
Microsoft Access Database Engine 2016 Redistributable
```

Important:

```text
The application platform target and the Access Database Engine architecture must match.

x64 application  →  64-bit Access Database Engine
x86 application  →  32-bit Access Database Engine
```

---

## 🚀 How to Run

1. Clone the repository.
2. Open the solution in Visual Studio.
3. Build the project.
4. Run the application.
5. Select the Access database file from the database settings form.
6. Load the required table from the menu.
7. Add, edit, search, delete, or export records.

---

## 🧪 Typical Workflow

1. Open the application.
2. Select the Microsoft Access database file.
3. Add customers from the customer form.
4. Add contract organizations from the contract-party form.
5. Add bitumen types to the database.
6. Open the delivery record form and create a new delivery record.
7. Load `Table1` in the main grid.
8. Use filters to search delivery records.
9. Double-click a record ID to edit it.
10. Select full rows and press `Delete` to remove records after confirmation.
11. Select numeric cells to calculate their sum at the bottom of the main window.
12. Export visible data to CSV when needed.

---

## 📁 Suggested Repository Structure

```text
DataManagementSoftware/
│
├── README.md
├── afshin.sln
├── afshin/
│   ├── MainForm.cs
│   ├── MainForm.Designer.cs
│   ├── RecordInfoForm.cs
│   ├── RecordInfoForm.Designer.cs
│   ├── ClienInfoForm.cs
│   ├── ClienInfoForm.Designer.cs
│   ├── ContractParty.cs
│   ├── ContractParty.Designer.cs
│   ├── DBInfoForm.cs
│   ├── DBInfoForm.Designer.cs
│   ├── Class1.cs
│   └── Program.cs
│
└── screenshots/
    ├── main-window.png
    ├── record-form.png
    ├── database-settings.png
    ├── contract-party-form.png
    └── customer-info-form.png
```

---

## 🔮 Future Improvements

- User login system
- User roles and permissions
- Advanced reporting
- PDF invoice generation
- Automatic database backup
- Better UI/UX design
- SQL Server support
- Search result highlighting
- More advanced financial reports
- Safer delete logs / soft delete system
- Audit log for edited and deleted records
- Better validation for Persian date formats
- Installer package for easier deployment

---

## 👨‍💻 Developer

Developed as a C# WinForms desktop application for managing customer, delivery, contract-party, and bitumen sales/dispatch records.

---

## 📄 License

This project is currently intended for personal, educational, or internal business use.
