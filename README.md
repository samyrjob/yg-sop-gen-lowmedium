🎓 YES WIZARD: Automated SOP Generator
YES WIZARD is an intelligent assistant designed to help students generate personalized, high-quality Statements of Purpose (SOP) for university applications (specifically targeted toward studies in Germany). By leveraging GPT-4o and real-time web scraping, the tool creates human-like, data-driven motivation letters.

🚀 Features
Resume Parsing: Automatically extracts academic and professional data from uploaded .pdf or .docx files.

Dynamic Web Intelligence: Uses the Diffbot API and DQL (Diffbot Query Language) to fetch real-time facts about German universities, city attractions, and local job markets.

SOP Personalization: Tailors the narrative based on the student's CGPA (using SpaCy for entity recognition) and specific program details.

Human-Like Tone Control: Strictly avoids AI-typical "buzzwords" to ensure the output passes as human-written.

Secure Authentication: Integrated with Firebase (Pyrebase) for user login and Encrypted Cookie Manager for session persistence.

Document Export: Generates and formats a professional Word (.docx) document with justified text and custom fonts.

🛠️ Technical Stack
Frontend: Streamlit

LLM: OpenAI API (GPT-4o)

Data Extraction: Diffbot API, BeautifulSoup4

NLP: SpaCy (Entity Ruler for CGPA extraction)

Backend/Auth: Firebase (via Pyrebase)

File Handling: PyPDF2, python-docx

📦 Installation & Setup
1. Clone the Repository
Bash
git clone https://github.com/your-username/yes-wizard.git
cd yes-wizard
2. Install Dependencies
Ensure you have Python 3.9+ installed.

Bash
pip install -r requirements.txt
python -m spacy download en_core_web_sm
3. Environment Variables
Create a .env file in the root directory and add your API keys:

Extrait de code
OPENAI_API_KEY=your_openai_key
PYREBASE_API_KEY=your_firebase_key
DIFFBOT_API_KEY=your_diffbot_key
4. Firebase Configuration
Update the firebaseConfig dictionary in the code with your specific Firebase project credentials.

5. Directory Structure
Ensure you have the following folders for the app to function correctly:

/images: Store your yeslogo2.jpg.

/templates4: Store sample SOP PDF files for the AI to use as structural guides.

🖥️ Usage
Start the App:

Bash
streamlit run app.py
Login: Use your Firebase credentials to access the wizard.

Upload: Drop your Resume/CV into the uploader.

Input Details: Enter the target Program Name and University.

Generate: The AI will scrape relevant university data and generate a 7-paragraph SOP.

Download: Click "Download SOP" to get your formatted .docx file.

🔍 How It Works (Internal Logic)
CGPA Detection: The app uses a SpaCy EntityRuler with RegEx patterns to find scores like "9.6 CGPA" or "SGPI 8.0" to decide the tone of the academic paragraph.

Caching: Scraped web data is stored in scraped_data.pkl to reduce API costs and improve speed for repeated queries.

Constraint Logic: The prompt engineering includes a "Negative Constraint" list, preventing the AI from using overused words like “leverage”, “cutting-edge”, or “foster”.
