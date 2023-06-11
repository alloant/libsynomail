#!/bin/python

from datetime import datetime
from pathlib import Path
import logging

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font

from libsynomail import EXT
from libsynomail.classes import File, Note
from libsynomail.scrap_register import Register
from libsynomail.syneml import write_eml

import libsynomail.connection as con

FONT = Font(name= 'Arial',
                size=12,
                bold=False,
                italic=False,
                strike=False,
                underline='none'
                #color='4472C4'
                )

FONT_BOLD = Font(name= 'Arial',
                size=12,
                bold=True,
                italic=False,
                strike=False,
                underline='none'
                #color='4472C4'
                )


def get_notes_in_folders(folders,ctrs,DEBUG = False):
    team_folders = con.nas.get_team_folders()

    files = []
    paths = []
    for key,path in folders.items():
        if key in ['ctr','dr']:
            for ctr in ctrs:
                value = ctr if 'name' not in ctrs[ctr] else ctrs[ctr]['name']
                paths.append([key,ctr,path.replace('@',value)])
        else:
            paths.append([key,'',path])
    
    for path in paths:
        if DEBUG and not (path[1] == 'gul' or path[1] == 'vind1'): continue

        logging.debug(f"Checking path {path}")
        files_in_folder = con.nas.get_file_list(path[2])
        
        if not files_in_folder: continue

        for file in files_in_folder:
            logging.info(f"Found in {path[0]} {path[1]}: {file['name']}")
            source = path[1] if path[0] != 'r' else file['name']
            files.append({"type": path[0],"source": path[1],"file": File(file)})
    
    
    return files

def notes_from_files(files,flow = 'in'):
    notes = {}
    
    files.sort(reverse=True,key = lambda file: f"{file['main']}_{file['file']}")

    for file in files:
        note = Note(file['type'],file['source'],file['num'],flow)
        if not note.key in notes:
            notes[note.key] = note
            
        notes[note.key].addFile(file['file'])

    return notes

def generate_register(path_register,notes):
    wb = Workbook()
    ws = wb.active
    ws.title = "Notes"
    ws.append(['Type','Source','No','Year','Ref','Date','Content','Dept','of_annex','Comments','Archived','Sent to'])
    for cell in ws[ws.max_row]:
        cell.alignment = Alignment(horizontal='center')
        cell.font = FONT_BOLD
    
    ws_data = wb.create_sheet(title="Data")
    ws_data.append(['Key','Folder_id','Permanent_link'])
 
    ws_files = wb.create_sheet(title="Files")
    ws_files.append(['Key','Name','Type','Display_data','File_id','Permanent_link','Original_id'])

    
    for key,note in notes.items():
        ws.append(note.exportExcel())
        ws[ws.max_row][5].number_format = 'dd/mm/yyyy;@'
       
        ws_data.append([note.key,note.folder_id,note.permanent_link])

        for cell in ws[ws.max_row]:
            cell.alignment = Alignment(horizontal='center')
            cell.font = FONT

        for file in note.files:
            ws_files.append([key] + file.exportExcel())

    column_widths = [10,10,10,12,12,15,50,20,12,50,10,20]
    for i, column_width in enumerate(column_widths,1):  # ,1 to start at 1
        ws.column_dimensions[get_column_letter(i)].width = column_width
    
    date = datetime.today().strftime('%Y-%m-%d-%HH-%Mm')
    con.nas.upload_convert_wb(wb,f"register-{date}.xlsx",f"{path_register}") 


def rename_file(source,file,new_name = ''):
    try:
        ext = Path(file.name).suffix[1:]
        if not new_name:
            new_name = file.name if source in file.name and source != 'r' else f"{source}_{file.name}"
            new_name = new_name.strip()
        else:
            new_name += f".{ext}"
        
        new_name = new_name.replace('&','and')

        con.nas.change_name(file.file_id,new_name)
        file.name = new_name
    except Exception as err:
        logging.error(err)
        logging.error(f"Problem changing name to {file.name}")

def manage_files_despacho(path_files,files,is_from_dr = False):
    flow = 'out' if is_from_dr else 'in'
    notes = notes_from_files(files,flow)
    
    if is_from_dr: #We only need this for the dr to know where they are sending the note
        register = Register('out')
        
    for note in notes.values():
        for i,file in enumerate(note.files):
            dest = f"{path_files}"
            
            if is_from_dr:
                note.dept = register.scrap_destination(note.no)
            #else:
            if i == 0: #It is the main file
                main_old_name = Path(file.name).stem
                num = f"0000{note.no}"[-4:]
                #main_name = f"{type}_{note.source}_{num}" if file.type == 'r' else f"{note.source}_{num}"
                main_name = f"{note.key.split('_')[0]}_{num}"
                rename_file(note.source,file,main_name)
            else:
                rename_file(note.source,file,Path(file.name.replace(main_old_name,main_name)).stem)

            if note.of_annex != '':
                note.folder_id,note.permanent_link = con.nas.create_folder(dest,main_name)
                dest += f"/{main_name}"
            #Before everything till here was inside the else

            con.nas.move(file.file_id,dest)
            ext = Path(file.name).suffix[1:]
            if ext in EXT:
                file.original_id = file.file_id
                file.name,file.display_path,file.file_id,file.permanent_link = con.nas.convert_office(file.file_id)

    generate_register(f"{path_files}",notes)
    
    return notes

def read_register_file(path_despacho,flow = 'in'):
    files_in_outbox = con.nas.get_file_list(f"{path_despacho}")
    files_in_outbox.sort(reverse = True,key = lambda file: file['name'])
    notes = ''
    for file in files_in_outbox:
        if file['name'][:9] in ["register-"] and 'osheet' in file['name']:
            notes = register_to_notes(file['display_path'],flow)
            break

    return notes


def register_to_notes(register,flow = 'in'):
    wb = load_workbook(con.nas.download_file(register))

    notes_data = list(wb['Notes'].iter_rows(values_only=True))[1:]
    data = list(wb['Data'].iter_rows(values_only=True))[1:]
    files = list(wb['Files'].iter_rows(values_only=True))[1:]
    
    notes = {}
    
    for i,row in enumerate(notes_data):
        no = row[2].split('","')[1][:-2]
        note = Note(row[0],row[1],no,flow=flow)
        note.year = row[3] if row[3] else ''
        note.ref = row[4] if row[4] else ''
        note.date = row[5] if row[5] else '' 
        note.content = row[6] if row[6] else ''
        note.dept = row[7] if row[7] else '' 
        note.comments = row[9] if row[9] else ''
        note.archived = row[10] if row[10] else ''
        note.sent_to = row[11] if row[11] else ''

        note.folder_id = data[i][1]
        note.permanent_link = data[i][2]

        for file in files:
            if file[0] == note.key:
                note.addFile(File({'name':file[1],'type':file[2],'display_path':file[3],'file_id':file[4],'permanent_link':file[5]},file[6]))

        notes[note.key] = note
    
    return notes


def rec_in_groups(recipients,RECIPIENTS,ctr = True):
    if recipients == 'all':
        for key,rec in RECIPIENTS:
            if 'email' in rec: #es forti
                if not ctr:
                    send_to.append(key)
            else:
                if ctr:
                    send_to.append(key)
        
        return list(set(send_to))


    recs = [rec.strip() for rec in recipients.split(',')]
    send_to = []
    ex_groups = []
    in_groups = []

    for rec in recs:
        if rec in RECIPIENTS:
            send_to.append(rec)
        else: #rec is a groups
            if rec.lower() == rec:
                in_groups.append(rec)
            else:
                ex_groups.append(rec)

    for key,rec in RECIPIENTS.items():
        putin = True
        for gp in ex_groups:
            if not gp.lower() in rec['groups']:
                putin = False
                break

        if putin:
            for gp in in_groups:
                if gp in rec['groups']:
                    if 'email' in rec: #es forti
                        if not ctr:
                            send_to.append(key)
                    else:
                        if ctr:
                            send_to.append(key)


    return list(set(send_to))

def new_mail_ctr(RECIPIENTS,note):
    send_to = rec_in_groups(note.dept,RECIPIENTS,True)
    
    for st in send_to:
        if not st.lower() in note.sent_to.lower():
            for file in note.files:
                con.nas.copy(file.file_id,f"/team-folders/Mailbox {st}/cr to {st}")
            note.sent_to += f",{st}" if note.sent_to else st
    
    return True

def new_mail_ctr_sheet(RECIPIENTS,note):
    wb = Workbook()
    ws = wb.active
    ws.append(['No','Date','Content','Ref'])
    ws.append([note.sheetLink(f"cr {note.no}/{note.year[2:]}"),note.date,note.content,note.ref])
    ws[ws.max_row][1].number_format = 'dd/mm/yyyy;@'

    send_to = rec_in_groups(note.dept,RECIPIENTS,True)
    
    for st in send_to:
        if not st.lower() in note.sent_to.lower():
            note.sent_to += f",{st}" if note.sent_to else st
            con.nas.upload_convert_wb(wb,f"new_mail.xlsx",f"/team-folders/Mailbox {st}/cr to {st}")
    
    return True

def new_mail_eml(RECIPIENTS,note,path_download):
    send_to = rec_in_groups(note.dept,RECIPIENTS,False)
    TO = []
    for st in send_to:
        if not st.lower() in note.sent_to.lower():
            note.sent_to += f",{st}" if note.sent_to else st
            TO.append(RECIPIENTS[st]['email'])
    if TO:
        write_eml(",".join(TO),note,path_download)
    return True

def new_mail_asr(note,path_download):
    if not 'asr' in note.sent_to:
        for file in note.files:
            con.nas.download_file(file.file_id,f"{path_download}/outbox asr",file.name)
        note.sent_to += f",asr" if note.sent_to else 'asr'
    return True

def register_notes(path_despacho,path_archive,RECIPIENTS,is_from_dr = False,path_download = None):
    register_dest = path_despacho if is_from_dr else f"{path_despacho}/Outbox Despacho"
    flow = 'out' if is_from_dr else 'in'
    notes = read_register_file(register_dest,flow)
    
    if not notes: return None

    for note in notes.values():
        if not note.archived and note.dept != '':
            if is_from_dr:
                dest = f"{path_archive}/{note.archive_folder}"
            else:
                dest = f"{path_archive}/{note.archive_folder}"
            
            con.nas.move(note.synology_id,dest)
            note.archived = True

            if len(note.files) == 1:
                if note.files[0].original_id:
                    con.nas.move(note.files[0].original_id,dest)
            
        if note.archived and note.dept != '':
            if is_from_dr:
                if note.no < 250:
                    rst = new_mail_eml(RECIPIENTS,note,path_download)
                elif note.no < 1000:
                    rst = new_mail_asr(note,path_download)
                elif note.no < 2000:
                    rst = new_mail_ctr(RECIPIENTS,note)
                else:
                    rst = new_mail_eml(RECIPIENTS,note,path_download)
            else:
                depts = [dep.lower().strip() for dep in note.dept.split(',')]
                for dep in depts:
                    if not dep.lower() in note.sent_to.lower():
                        rst = con.nas.send_message(dep,RECIPIENTS,note.message)
                        if rst: note.sent_to += f",{dep}" if note.sent_to else dep

    generate_register(register_dest,notes)
