# MOORA Method Web App (Streamlit)

This repository contains a simple and user-friendly web application built with **Streamlit** to perform decision analysis using the **MOORA** (Multi-Objective Optimization on the Basis of Ratio Analysis) method.

---

## 🚀 Features

- Upload `.xlsx` file with:
  - Decision matrix
  - Weights
  - Types of criteria (benefit/cost)
- Perform MOORA calculations automatically
- View and download final ranking in Excel format

---

## 📁 Required Excel Structure

### Sheet 1: `DecisionMatrix`

|     | C1 | C2 | C3 |
|-----|----|----|----|
| A1  | 10 | 5  | 8  |
| A2  | 8  | 7  | 6  |
| A3  | 9  | 6  | 7  |

### Sheet 2: `Weights`

| C1  | C2  | C3  |
|-----|-----|-----|
| 0.4 | 0.3 | 0.3 |

> **Note**: Weights must sum to 1.

### Sheet 3: `Types`

| C1     | C2   | C3     |
|--------|------|--------|
| benefit| cost | benefit|

---

## 🛠️ Installation

```bash
git clone https://github.com/yourusername/moora-app.git
cd moora-app
pip install -r requirements.txt
streamlit run app_moora.py
```

---

## 📦 Dependencies

- streamlit
- pandas
- numpy
- openpyxl
- xlsxwriter

---

## 📤 Output

- Final scores and ranks
- Excel file with:
  - Normalized matrix
  - Weighted matrix
  - MOORA results

---

## 👤 Author

- Zahari Md Rodzi (UiTM)
- 📧 zahari@uitm.edu.my

---

## 📄 License

This project is licensed under the BSD-3-Clause License.