#!/bin/python

from getpass import getpass
from datetime import datetime
import time
import re
from pathlib import Path
import logging
import ast

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

from libsynomail import CONFIG, EXT, DEBUG
import libsynomail.connection as con

def folder_in_teams(folder,teams):
    fds = folder.split("/")[1:]
    for team in teams:
        tms = team['folder'].split("/")[1:]
        same = True
        key = team['type']
        for i,fd in enumerate(fds):
            if '@' in tms[i]:
                pt = tms[i].replace('@','')
                if fds[i][:len(pt)] == pt:
                    key = fds[i][len(pt):]
                else:
                    same = False
                    break
            else:
                if fds[i] != tms[i]:
                    same = False
                    break

        if same:
            nt = team.copy()
            nt['folder'] = team['folder'].replace('@',key)
            return same,key,nt

    return False,'asd',''


def get_notes_in_folders():
    year = datetime.today().strftime('%Y')

    team_folders = con.nas.get_team_folders()
    mail_sources = [ast.literal_eval(tm) for tm in CONFIG['teams'].split('|')]

    reg_notes = {}
    for folder in team_folders:
        mail_folder,key,team = folder_in_teams(f"/team-folders/{folder}",mail_sources)
        
        if mail_folder and (key == 'gul' or not DEBUG):
            logging.debug(f"Checking folder {team['folder']}")
            notes = con.nas.get_file_list(team['folder'])
            for note in notes:
                logging.info(f"Found note {note['name']} in {team['folder']}")
                reg_notes[note['name']] = team.copy()|{'source':key,'Original':'','p_link':note['permanent_link']}

    return reg_notes

def create_despacho(ws,reg_notes):
    year = datetime.today().strftime('%Y')
    ws.append(['type','source','No','Year','Ref','Date','Content','Dept','Name','Original','Comments'])
    
    for name,note in reg_notes.items():
        num = re.findall('\d+',name.replace(note['source'],''))
        num = num[0] if num else ''
        
        #num = f"000{num[0]}"[-4:] if num else ''
        #if num and note['type'] == 'ctr in': num = num[1:]
        
        note['num'] = num

        note['year'] = year

        if note['type'] == 'r in':
            src = re.findall('\D+',name)
            note['source'] = src[0] if src else ''
        
        nm = note['link'] if 'link' in note else name

        ws.append([note['type'],note['source'],num,year,'','','','',nm,note['Original'],''])    
        ws[ws.max_row][5].value = datetime.today()
        ws[ws.max_row][5].number_format = 'dd/mm/yyyy;@'

        column_widths = [10,10,10,10,12,12,50,12,20,20]
        for i, column_width in enumerate(column_widths,1):  # ,1 to start at 1
            ws.column_dimensions[get_column_letter(i)].width = column_width

        for row in ws[1:ws.max_row]:  # skip the header
            for i,col in enumerate(column_widths):
                cell = row[i]             # column H
                cell.alignment = Alignment(horizontal='center')

def change_names(notes):
    new_names = []
    for name,note in notes.items():
        try:
            new_name = name if note['source'] in name and note['source'] != 'r' else f"{note['source']}_{name}"
            new_name = new_name.strip()
            #new_name = new_name.replace(' ','_')
            new_names.append([new_name,name])
            con.nas.change_name(f"{note['folder']}/{name}",new_name)
        except Exception as err:
            logging.error(err)
            logging.warning(f"Cannot change name of {name}")

    for new in new_names:
        notes[new[0]] = notes.pop(new[1])


def move_to_despacho(notes):
    for name,note in notes.items():
        ext = Path(name).suffix[1:]
        name_link = name
                
        con.nas.move(f"{note['folder']}/{name}",f"{CONFIG['despacho']}/Inbox Despacho")
                
        note['folder'] = f"{CONFIG['despacho']}/Inbox Despacho"

        chain = 'oo/r' if ext in EXT.values() else 'd/f'
                
        p_link = note['p_link']
        link = f'=HYPERLINK("#dlink=/{chain}/{p_link}", "{name_link}")' if p_link != '' else name_link
        note['link'] = link


def convert_files(notes):
    new_names = []
    for name,note in notes.items():
        name_link = name
        # Here I check if I can convert the file to synology with the extension
        ext = Path(name).suffix[1:]
        if ext in EXT:
            f_path,f_id,p_link = con.nas.convert_office(f"{CONFIG['despacho']}/Inbox Despacho/{name}")
                    
            note['Original'] = name
            new_name = f"{name[:-len(ext)]}{EXT[ext]}"
            new_names.append([new_name,name])
            if p_link != '':
                note['link'] = f'=HYPERLINK("#dlink=/oo/r/{p_link}", "{new_name}")'


    for new in new_names:
        notes[new[0]] = notes.pop(new[1])

def upload_register(wb,name,dest):
    con.nas.upload_convert_wb(wb,name,dest) 
            

def init_get_mail():
    logging.info('Starting searching new mail --------------------')
    reg_notes = get_notes_in_folders()
    
    date = datetime.today().strftime('%Y-%m-%d-%HH-%Mm')
    name = f"despacho-{date}.xlsx"

    if reg_notes != {}:
        wb = Workbook()
        ws = wb.active
        
        try:
            change_names(reg_notes)
            move_to_despacho(reg_notes)
            convert_files(reg_notes)
            create_despacho(ws,reg_notes)
            upload_register(wb,name,f"{CONFIG['despacho']}/Inbox Despacho")
        except:
            create_despacho(ws,reg_notes)
            upload_register(wb,name,f"{CONFIG['despacho']}/Inbox Despacho")
    
    logging.info('Finish searching new mail ~~~~~~~~~~~~~~~~~~~~~~')
    
    

def main():
    init_get_mail()
    input("Pulse Enter to continue")


if __name__ == '__main__':
    main()
