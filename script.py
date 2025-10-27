
# === IMPORTING LIBRARIES ===
from collections import defaultdict
from datetime import datetime, timedelta
import os
import time
import traceback
import win32com.client

from dotenv import load_dotenv
import openai
import yaml

# === CONFIGURATION ===
load_dotenv()
with open('config.yaml','r') as f:
    config = yaml.safe_load(f)
SCOPES = ["Mail.Read"]
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
YOUR_EMAIL_CAPITALIZE = config['YOUR_EMAIL_CAPITALIZE']  
YOUR_EMAIL = YOUR_EMAIL_CAPITALIZE.lower() 
INBOX_FOLDER_NAME = config['INBOX_FOLDER_NAME'] 
SUB_FOLDER_NAME = config['SUB_FOLDER_NAME']
DAYS_BACK = config['DAYS_BACK']
FROM_DAYS = config['FROM_DAYS']
LANGUAGE = config['LANGUAGE']
EXCLUDED_SENDERS = config['EXCLUDED_SENDERS']
SPLITTERS = config['SPLITTERS']

with open("filter.txt", "r", encoding="utf-8") as f:
    filter_list = [line.strip() for line in f if line.strip()]
replacing_dict = dict((line.strip().split(",")[0], line.strip().split(",")[1]) for line in open('replace.txt', 'r'))

days = list()
for number in range(FROM_DAYS, DAYS_BACK):
    days.append((datetime.now() - timedelta(days=number)).strftime("%d-%m-%Y"))
days.sort()

def extract_last_reply(body):
    if not body:
        return ""
    new_body = body.replace("\r\n", "\n")
    for splitter in SPLITTERS:
        if splitter in new_body:
            new_body = new_body.split(splitter, 1)[0]
            break
    new_body = new_body.strip().lower()
    for secret in filter_list:
        new_body = new_body.replace(secret, "")
    for k, v in replacing_dict.items():
        new_body = new_body.replace(k, v)
    return new_body

# === OPENAI SETUP ===
client = openai.OpenAI(api_key=OPENAI_API_KEY)

# === OUTLOOK SETUP ===
outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")
for account in mapi.Folders:
    if account.Name == YOUR_EMAIL_CAPITALIZE:
        main_account = account

for folder in main_account.Folders:
    if folder.Name == INBOX_FOLDER_NAME:
        main_folder = folder

for subfolder in main_folder.Folders:
    if subfolder.Name == SUB_FOLDER_NAME:
        cc_folder = subfolder

if not cc_folder:
    raise Exception(f"Folder '{SUB_FOLDER_NAME}' not found in Outlook.")
messages = cc_folder.Items
messages.Sort("[ReceivedTime]", True) 

# === SUMMARY BY OPENAI ===
def summarise(messages_with_senders):
    combined = "\n\n".join(f"{name} wrote:\n{body}" for name, body in messages_with_senders)
    prompt = ()
    
    prompt = (
        "Summarize these emails in one clear paragraph, stating who said what. "
        "Mention the sender’s name and keep it short and to the point."
        "Follow standard grammar rules when writing the summary."
        "Leave out irrelevant details and be concise."
        "Combine multiple emails with the same subject into a single summary\n\n"
        f"Write the summary in the following language: {LANGUAGE}"
        f"{combined}"
    )
    response = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt[:8192]}],
        temperature=0.3,
    )
    return response.choices[0].message.content.strip()

# === GROUPING EMAILS BY SUBJECT ===
for day in days:
    threaded_emails = defaultdict(list)
    for msg in messages:
        try:
            date_object = (datetime.strptime(str(msg.ReceivedTime), "%Y-%m-%d %H:%M:%S.%f%z")).strftime("%d-%m-%Y")
        except:
            try:
                date_object = (datetime.strptime(str(msg.ReceivedTime),  "%Y-%m-%d %H:%M:%S%z")).strftime("%d-%m-%Y")
            except:
                print(msg)
        if datetime.strptime(day, "%d-%m-%Y") == datetime.strptime(date_object, "%d-%m-%Y"):
            try:
                sender = msg.SenderEmailAddress.lower()
                sender_name = msg.SenderName.lower()
                for secret in filter_list:
                    sender_name = sender_name.replace(secret, '')
                for replace in replacing_dict:
                    sender_name.replace(replace, replacing_dict[replace])
                if sender in EXCLUDED_SENDERS:
                    continue
                subject = msg.Subject or "(no subject)" # If necessary add 'no subject' in your language
                body = extract_last_reply(msg.Body.strip())
                if not body:
                    continue
                threaded_emails[subject].append((sender_name, body))
            except Exception:
                print(traceback.format_exc()) 

   # === CREATING THE HTML STRUCTURE ===
    html_body = "<html><body>"
    html_body += f"<h2>📬 Summary of your {SUB_FOLDER_NAME}-mails – {day}</h2>"
    for subject_new, messages_new in threaded_emails.items():
        summary = summarise(messages_new)
        html_body += f"<h3>📌 {subject_new}</h3><p>{summary}</p><hr>"
    html_body += "</body></html>"

    # === CREATE & SEND EMAIL WITH OUTLOOK ===
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    mail = outlook_app.CreateItem(0)
    mail.Subject = f"📬 Summary of your {SUB_FOLDER_NAME}-mails – {day}"
    mail.HTMLBody = html_body
    mail.To = YOUR_EMAIL
    mail.Send()
    print("✅ Summary sent to:", YOUR_EMAIL)
    time.sleep(60)
