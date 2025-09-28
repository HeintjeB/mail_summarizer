# From inbox overload to quick insights: My AI mail summarizer project

This repository accompanies the blog post: [From inbox overload to quick insights: My AI mail summarizer project](https://medium.com/@heinburgmans/from-inbox-overload-to-quick-insights-my-ai-mail-summarizer-project-d6a2720c8340).  
The README.md file explains how to configure the project files for your own use.

## 📂 Relevant Files

- **`script.py`**  
  Runs the mail summarizer. **Do not modify this file.**

- **`.env`**  
  Stores the required secret keys.

- **`filter.txt`**  
  Defines which confidential data should be removed from the prompt.

- **`replacing.txt`**  
  Replaces confidential information with dummy values.

- **`config.yaml`**  
  Configuration file where you can tweak settings to your preferences.

---

## ⚙️ Configuration

Below you will find instructions on how to adjust the files to your needs.

- **`.env`**  
  This file is created separately to store keys you don't want to share or commit to GitHub. The keys are:
  - CLIENT_ID: Your Microsoft Client ID
  - AUTHORITY: https://login.microsoftonline.com/<TENANT_ID_OF_YOUR_ORGANISATION>
  - OPENAI_API_KEY: your_openai_api_key  
  <i>Your LLM provider or Google can explain how to retrieve these keys.</i>

- **`.config`**  
  In this file you can configure the script to your needs. Variables include:
    - YOUR_EMAIL_CAPITALIZE: Your email account (as shown in your Outlook app) that you want to use to extract the emails — please be precise!
    - INBOX_FOLDER_NAME: The name of the inbox folder.
    - SUB_FOLDER_NAME: The subfolder you want to create the summaries for.
    - FROM_DAYS: How many days for now today you want to start extracting? 0 is today, 1 is yesterday.
    - DAYS_BACK: Till how far the script want to go.
    - EXCLUDED_SENDERS: A list with senders you want to exclude, for example no-reply mails.
    - SPLITTERS: A list with standard text which mostly indicate the begin of a mail.
    - LANGUAGE: The language you want to have the summary presented in.

- **`replacing.txt`**  
  The replacing is designed as follows:
  - Everything must be written in lowercase.
  - The file is executed line by line.
  - Everything before the comma is replaced with everything after the comma.

- **`filter.txt`**  
  - Add one entry per line with text you want to remove from the email body.

## 🛠️ Installation Instructions

Follow these steps to set up and run the script:

1. **Install Python**  
   - Download and install Python from [python.org](https://www.python.org/) or from the Microsoft Store.  
   - Make sure you check the option **“Add Python to PATH”** during installation (important for using Python from the command line).

2. **Get the project files onto your computer**  
   You can do this in two ways:
   - **If you use GitHub (recommended):**  
     Clone the repository to your computer using:  
     ```bash
     git clone https://github.com/your-username/your-repository.git
     ```
     Then navigate into the project folder:  
     ```bash
     cd your-repository
     ```

   - **If you are not familiar with GitHub or Git:**  
     Simply download the repository as a ZIP file:  
     1. Go to the project page on GitHub.  
     2. Click the green **Code** button.  
     3. Select **Download ZIP**.  
     4. Extract the ZIP file on your computer.  

3. **Create a virtual environment**  
   - A virtual environment keeps the project’s libraries separate from other programs on your computer.  
   - Open a terminal (Command Prompt or PowerShell) in the project folder and run:  
     ```bash
     python -m venv .venv
     ```

4. **Activate the virtual environment**  
   - On **Windows**:  
     ```bash
     .venv\Scripts\activate
     ```  
   - On **Mac/Linux**:  
     ```bash
     source .venv/bin/activate
     ```

5. **Install the required libraries**  
   - Once the environment is active, install all the necessary Python packages:  
     ```bash
     python -m pip install -r requirements.txt
     ```

6. **Configure the project files**  
   - Edit the following files to match your setup:  
     - `config.yaml` → script configuration  
     - `filter.txt` → text you want removed from emails  
     - `replacing.txt` → text replacements  

7. **Run the script**  
   - Start the script by running:  
     ```bash
     python script.py
     ```
---

## 🔧 Optional Automation

8. **Create a `.bat` file** (Windows only)  
   - This file lets you start the script with a double click.  
   - Example content:  
     ```bat
     {source_of_.venv}\.venv\Scripts\python.exe {source_of_script}\script.py
     ```

9. **Schedule the script to run automatically**  
   - Open **Task Scheduler** in Windows.  
   - Create a new task and point it to the `.bat` file.  
   - Set it to run daily (or as often as you like).  
