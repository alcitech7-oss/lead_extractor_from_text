# 🧠 Lead Extractor from Text

This script extracts names, phone numbers, and emails from a text and saves the results organized in an Excel spreadsheet. It uses **SpaCy** for person detection and **regular expressions** for contacts.

---

## 🚀 What this project does

- Reads a text file (`texto_leitura.txt`) containing raw text
- Uses **SpaCy** (a Natural Language Processing library) to detect person names
- Uses **Regular Expressions (Regex)** to identify:
  - 📧 Email addresses
  - 📞 Phone numbers
- Saves the extracted data into a structured **Excel spreadsheet** (`.xlsx`)

---

## 📁 Project Structure
lead_extractor_from_text/
├── lead_extractor_from_text.py # Main application script
├── texto_leitura.txt # Input file with raw text data
├── .gitignore # Git ignore file
├── requirements.txt # Python dependencies
└── README.md # This file

---

## 🧩 Technologies Used

- Python 3.10+
- **SpaCy** (for person name detection)
- Regular Expressions (`re`)
- Pandas
- OpenPyXL (to generate Excel files)

---

## 📦 How to Run

## 📦 How to Run

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/alictech7-oss/lead_extractor_from_text.git
    cd lead_extractor_from_text
    ```

2.  **Create and activate a virtual environment:**
    ```bash
    python -m venv venv
    venv\Scripts\activate
    ```

3.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Download the SpaCy Portuguese model:**
    ```bash
    python -m spacy download pt_core_news_sm
    ```

5.  **Prepare your data:** Replace the content of `texto_leitura.txt` with your own text.

6.  **Run the script:**
    ```bash
    python lead_extractor_from_text.py
    ```

7.  **Check the results:** The extracted leads will be saved in an Excel file (e.g., `leads_extraidos.xlsx`).
   
   📌 Notes
. The script is designed to be a simple and focused tool for lead extraction.

. You can customize the regex patterns and the SpaCy model in the main script to adapt to different text formats or languages.

🔮 Future improvements
. Add support for more file formats (.docx, .pdf)

. Implement a simple graphical interface with Streamlit

. Integrate with an API to validate emails and phone numbers

🙏 Credits & Original Work
This project was developed by alictech7-oss.

📄 License
MIT — use, modify, and share freely.
