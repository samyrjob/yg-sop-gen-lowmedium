# 🎓 YES WIZARD: Automated SOP Generator

**YES WIZARD** is an intelligent assistant designed to help students generate personalized, high-quality **Statements of Purpose (SOP)** for university applications, specifically optimized for the German education landscape. 

By leveraging **GPT-4o** and real-time web scraping, the tool creates human-like, data-driven motivation letters that stand out from generic templates.

---

## 🚀 Features

* **Resume Parsing:** Automatically extracts academic and professional data from uploaded `.pdf` or `.docx` files.
* **Dynamic Web Intelligence:** Uses the **Diffbot API** and **DQL** (Diffbot Query Language) to fetch real-time facts about German universities, city attractions, and local job markets.
* **SOP Personalization:** Tailors the narrative based on the student's **CGPA** (using **SpaCy** for entity recognition) and specific program details.
* **Human-Like Tone Control:** Strictly filters out AI "buzzwords" (e.g., *leverage, cutting-edge, foster*) to ensure the output reads naturally and passes admissions filters.
* **Secure Authentication:** Integrated with **Firebase** (Pyrebase) for secure login and **Encrypted Cookie Manager** for session persistence.
* **Document Export:** Generates a professional Word (`.docx`) document with justified alignment and standard typography.

---

## 🛠️ Technical Stack

| Category | Technology |
| :--- | :--- |
| **Frontend** | [Streamlit](https://streamlit.io/) |
| **LLM** | [OpenAI API (GPT-4o)](https://openai.com/) |
| **Data Extraction** | Diffbot API, BeautifulSoup4, Aiohttp |
| **NLP** | SpaCy (Entity Ruler for CGPA detection) |
| **Backend/Auth** | Firebase (Pyrebase), Cookies Manager |
| **File Handling** | PyPDF2, python-docx |

