# Campaign Finance ETL

This project automates the **ETL (Extract, Transform, Load)** process for campaign finance data (contribution and expenditure) extracted from PDF files. It downloads, parses, transforms, and saves structured financial data into Excel files for further analysis or reporting.

---

## 📟 Features

* ✅ **Automated download** of campaign finance PDFs
* 📄 **Dynamic extraction** of contribution or expenditure data based on form content
* 🧹 **Robust transformation** logic to clean names, addresses, dates, and monetary values
* 📊 **Excel export** of clean structured data
* ↻ Entire pipeline is class-based and modular (Download → Extract → Transform)

---

## 🗂️ Project Structure

All code is written inside a single script file:
`campaignETL.py`

It contains 3 main classes:

| Class Name        | Purpose                                                  |
| ----------------- | -------------------------------------------------------- |
| `PDFDownloader`   | Automates the downloading of finance PDFs                |
| `PDFExtractor`    | Extracts raw data (Name, Address, etc.) from PDFs        |
| `DataTransformer` | Transforms and cleans the extracted data into final form |

---

## ⚙️ Setup Instructions

### 1. Clone the Repository

```bash
git clone https://github.com/madhuthakur-0212/campaign-finance-etl.git
cd campaign-finance-etl
```

### 2. Create a Virtual Environment (Recommended)

```bash
python -m venv venv
venv\Scripts\activate   # On Windows
# OR
source venv/bin/activate   # On Mac/Linux
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

---

## 🚀 How to Run

Simply run the Python script:

```bash
python campaignETL.py
```

This will:

1. Download PDF files
2. Identify the type of forms
3. Extract relevant data
4. Clean and transform it
5. Export the final output to:

```bash
transformed_data.xlsx
```

---

## 📁 Output Files

* `contributions_output_combined.xlsx`: Intermediate data after extraction
* `transformed_data.xlsx`: Final structured output file (cleaned and ready to use)

---

## 👍 .gitignore Includes

* Intermediate Excel files
* PDF files
* Virtual environment folders
* System/cache files

---

## 📌 Requirements

The main libraries used:

```txt
pandas
openpyxl
python-dateutil
PyMuPDF==1.22.0
```

---

## 🤝 Contributing

Pull requests and feedback are welcome. If you find bugs or want to suggest features, feel free to open an issue.

---


