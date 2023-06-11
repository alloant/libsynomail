import eml_parser
import io
import base64

from email import generator
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from pathlib import Path

from libsynomail import EXT,INV_EXT
import libsynomail.connection as con

def write_eml(rec,note,path_download):
    msg = MIMEMultipart()
    msg["To"] = rec
    msg["From"] = 'Aes-cr@cardumen.org'
    msg["Subject"] = f"{note.key}/{note.year[2:]}: {note.content}"
    body = ""
    msg.attach(MIMEText(body,"plain"))

    for file in note.files:
        ext = Path(file.name).suffix[1:]
        file_name = f"{Path(file.name).stem}.{INV_EXT[ext]}" if ext in INV_EXT else file.name

        attachment = con.nas.download_file(file.file_id)
        part = MIMEApplication(attachment.read(),Name=file_name)

        part['Content-Disposition'] = f'attachment; filename = {file_name}'
        msg.attach(part)

    with open(f"{path_download}/outbox forti/{note.key}.eml",'w') as file:
        emlGenerator = generator.Generator(file)
        emlGenerator.flatten(msg)


def read_eml(path_eml):
    parsed_eml = eml_parser.parser.decode_email(path_eml,include_attachment_data=True)
    sender = parsed_eml['header']['from']
    if sender == "cg@cardumen.org":
        dest = "/team-folders/Mail cg/Mail from cg"
    else:
        dest = "/team-folders/Mail r/Mail from r"

    attachments = parsed_eml['attachment']
    
    for file in attachments:
        b_file = io.BytesIO(base64.b64decode(file['raw']))
        b_file.name = file['filename']
        con.nas.upload_file(b_file,dest)
