# DRE Financial Automation Engine

An intelligent, AI-powered desktop application designed to automate the extraction, classification, and consolidation of financial transactions into a centralized Google Sheets DRE (Demonstração do Resultado do Exercício).

Developed to solve real-world operational bottlenecks for a local business, this application streamlines the month-end financial closing process by converting unstructured bank statements and credit card bills into clean, categorized data.

---

## 🚀 Core Features & Functionalities

- **Multi-Source Data Extraction:**
  - **Native Integrations:** Built-in parsing logic for standard spreadsheet formats from Bradesco, Caixa PJ, and specific credit card statements (e.g., Nubank, PagVeloz).
  - **AI-Powered Fallback:** For PDFs and unrecognized formats, the system leverages OpenAI (`gpt-4o-mini`) to dynamically read, understand, and extract transaction data.
- **Smart Categorization Engine:**
  - Automatically classifies transactions into Revenue (`RECEITAS`) or Expenses (`DESPESAS`) based on predefined categories fetched directly from the master Google Sheet.
  - **Machine Learning Memory:** Features an adaptive memory system. If a user manually corrects an AI classification in the UI, the engine records this correction to a "Memory Sheet" on Google Drive, improving future automated classifications.
- **Interactive GUI (Tkinter):**
  - A clean, intuitive desktop interface built with native Python Tkinter.
  - Allows users to drag-and-drop or select multiple statement files, securely input their OpenAI API keys, and preview/edit the extracted data before confirming the export.
- **Direct Google Workspace Integration:**
  - Uses Google Service Accounts (`gspread` and OAuth2) to seamlessly write the final, verified transactions in bulk directly to the client's cloud-hosted DRE spreadsheet, negating the need for manual copy-pasting.

## 🛠️ Technical Stack & Architecture

- **Backend / Core Logic:** Python 3
- **GUI Framework:** Tkinter (Custom styled for a modern look)
- **AI & NLP:** OpenAI API (`gpt-4o-mini`) for dynamic text extraction and logical categorization constraint-solving.
- **Document Processing:** 
  - `openpyxl` for Excel (`.xlsx`, `.xlsm`) reading/writing.
  - `pdfplumber` for text extraction from PDF bank statements.
- **Cloud Integration:** Google Sheets API (`gspread`, `google-auth`) for reading categories, storing learned rules, and appending rows securely.
- **Concurrency:** Dedicated background threads handle file processing, API requests, and Google Drive upload events to ensure the UI remains responsive during network I/O operations.

## 📂 Project Structure

```text
├── script.py          # Main application entry point and Tkinter GUI implementation
├── engine.py          # Core DREEngine: Google Sheets I/O, format detection, and native spreadsheet scrapers
├── ai_mapper.py       # OpenAI integration for unstructured extraction and intelligent categorization
├── config.json        # Local user preferences (e.g., cached API keys)
├── .env               # Environment variables
└── credentials.json   # Google Cloud Service Account credentials (ignored in git)
```

## 💡 Engineering Highlights for Recruiters

- **Robust Error Handling & Fallbacks:** The app utilizes a tiered extraction approach. If deterministic parsing (regex/index-based) fails, it cleanly delegates to the AI mapper.
- **Scalable State Management:** The classification memory is decoupled from the local machine and stored on Google Sheets. This ensures that the application's "learned" categorizations persist and are immediately available to other instances or environments.
- **Clean Architecture:** Separation of concerns is maintained. The graphical layer (`script.py`) manages solely the presentation and event handling, the `engine.py` orchestrates I/O and domain logic, while `ai_mapper.py` is an isolated module exclusively handling LLM prompt engineering and schema validation.
- **Prompt Engineering:** Uses structured JSON mode with OpenAI to guarantee strict schema adherence, preventing hallucinated keys or misformatted data during the extraction phase.

## 🚧 Status
Actively in development. Currently deployed locally to assist a business client with their monthly financial reconciliations.
