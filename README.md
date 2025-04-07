# web_scrape_oracle

## 📄 Description

This Python script automates the extraction of text and table names from the Oracle Cloud Human Resources documentation site. It navigates the site's dropdown-based sidebar using Selenium and collects structured information under specific sections, which are then processed with pandas into Excel files.

## ✨ Key Features
- Scrapes dropdown-based TOC (Table of Contents) using Selenium automation
-Handles dynamic iframe content and cookie consent interaction
-Automatically zooms out for better visibility and performance
-Extracts table names from specific sections (e.g., "2 AI")
-Includes dynamic wait handling for smoother execution
-Logs and prints clear step-by-step actions and errors

## 🛠️ Technologies Used
- Python 3.x
- Selenium WebDriver
- ChromeDriver
- WebDriver Manager
- Oracle Cloud Human Resources documentation

## 📂 Project Structure
```bash
oracle-hr-docs-scraper/
├── script.py               # Main Selenium script
├── requirements.txt        # Python dependencies
└── README.md               # Project overview and instructions
```

## ▶️ How to Run the Project
1. Clone this repository:
```bash
git clone https://github.com/your-username/oracle-hr-docs-scraper.git
cd oracle-hr-docs-scraper
```

2. Install Dependencies
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```
```bash
pip install -r requirements.txt
```

3. Run the Script
```bash
python script.py
```

## 📝 Requirements File (requirements.txt)
```bash
selenium
webdriver-manager
```

## 🧾 License
This project is licensed under the MIT License. See the LICENSE file for details.

## 🤝 Contributing
Contributions are welcome! Here's how to help:

1. Fork the repository
2. Create your feature branch (git checkout -b feature/something)
3. Commit your changes (git commit -am 'Add something')
4. Push to the branch (git push origin feature/something)
5. Open a Pull Request















