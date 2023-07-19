import eml_parser
import io
import base64

from email import generator
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from pathlib import Path

from libsynomail import EXT,INV_EXT
from libsynomail.nas import upload_path
#import libsynomail.connection as con

def write_eml(rec,note,path_download):
    msg = MIMEMultipart()
    msg["To"] = rec
    msg["From"] = 'Aes-cr@cardumen.org'
    msg["Subject"] = f"{note.key}/{note.year[2:]}: {note.content}"
    msg.add_header('X-Unsent','1')
    body = ""
    msg.attach(MIMEText(body,"plain"))

    rst = True
    for file in note.files:
        ext = Path(file.name).suffix[1:]
        file_name = f"{Path(file.name).stem}.{INV_EXT[ext]}" if ext in INV_EXT else file.name

        attachment = file.download()
        
        if attachment:
            part = MIMEApplication(attachment.read(),Name=file_name)

            part['Content-Disposition'] = f'attachment; filename = {file_name}'
            msg.attach(part)
        else:
            rst = False

    if rst:
        with open(f"{path_download}/outbox forti/{note.key}.eml",'w') as file:
            emlGenerator = generator.Generator(file)
            emlGenerator.flatten(msg)
            return True
    return False


def read_eml(path_eml,emails = None):
    parsed_eml = eml_parser.parser.decode_email(path_eml,include_attachment_data=True)
    sender = parsed_eml['header']['from']
    if sender == "cg@cardumen.org":
        dest = "/team-folders/Mail cg/Mail from cg"
        bf = 'cg'
    else:
        dest = "/team-folders/Mail r/Mail from r"
        bf = 'r'

    if bf == 'r':
        for r,email in emails.items():
            if sender.lower() == email['email'].lower():
                bf += f"_{r}"
                break

    if 'attachment' in parsed_eml:
        attachments = parsed_eml['attachment']
    
        for file in attachments:
            b_file = io.BytesIO(base64.b64decode(file['raw']))
            b_file.name = f"{bf}_{file['filename']}"
            upload_path(b_file,dest)
