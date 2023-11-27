import imaplib
from tqdm import tqdm
import email
import os


def get_messsages_gmail(
    user_email, user_password, last_email=-1, email_folder="INBOX", from_email="All"
):
    """This function extracts the messages objects from a gmail account"""

    # connect to gmail
    gmail = imaplib.IMAP4_SSL("imap.gmail.com")

    # sign in with your credentials
    gmail.login(user_email, user_password)

    # select the folder
    gmail.select(email_folder)

    if from_email == "All":
        resp, items = gmail.search(None, from_email)
    else:
        resp, items = gmail.search(None, f"(FROM {from_email})")

    items = items[0].split()
    msgs = []
    for num in tqdm(items[:last_email]):
        typ, message_parts = gmail.fetch(num, "(RFC822)")
        msgs.append(message_parts)

    return msgs


msgs = get_messsages_gmail(
    "name@gmail.com",
    "pass",
    last_email=-1,
    email_folder="INBOX",
    from_email="All",
)
data_folder = './py_attachments'

def get_pdf_attachments(msgs, data_folder):
    """This function extracts the pdf files"""
    for msg_raw in msgs:
        if type(msg_raw[0]) is tuple:
            msg = email.message_from_string(str(msg_raw[0][1], "utf-8"))
            for part in msg.walk():
                if part.get_content_maintype() == "multipart":
                    continue
                if part.get("Content-Disposition") is None:
                    continue

                try:
                    if (".pdf" in part.get_filename()) or (
                        ".PDF" in part.get_filename()
                    ):
                        filename = part.get_filename()
                        file_path = os.path.join(data_folder, filename)

                        # Save the file
                        with open(file_path, "wb") as file:
                            file.write(part.get_payload(decode=True))

                except Exception as error:
                    print(error)
                    pass
get_pdf_attachments(msgs, data_folder)