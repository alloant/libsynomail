from pathlib import Path
from attrdict import AttrDict
from libsynomail import INV_EXT
from datetime import datetime

import libsynomail.connection as con

class File(AttrDict):
    def __init__(self,data,original_name = ''):
        self.name = data['name']
        self.type = data['type']
        self.path = str(Path(data['display_path']).parent)
        self.file_id = data['file_id']
        self.permanent_link = data['permanent_link']
        self.original_name = original_name
    
    def __str__(self): 
         return self.name
    
    @property
    def display_path(self):
        return f"{self.path}/{self.name}"

    @property
    def ext(self):
        return Path(self.name).suffix[1:]

    @property
    def chain_link(self):
        if self.type == 'dir' or not self.ext in INV_EXT:
            return 'd/f'
        else:
            return 'oo/r'

    def getLinkSheet(self,text = None):
        link_text = text if text else self.name
        return f'=HYPERLINK("#dlink=/{self.chain_link}/{self.permanent_link}", "{link_text}")'

    def getLinkMessage(self,text = None):
        link_text = text if text else self.name
        return f'<https://nas.prome.sg:5001/{self.chain_link}/{self.permanent_link}|{link_text}>'

    def exportExcel(self):
        return [self.name,self.type,self.display_path,self.file_id,self.permanent_link,self.original_name]

    def move(self,dest):
        con.nas.move(self.display_path,dest)
        if self.original_name != '':
            con.nas.move(f"{self.path}/{self.original_name}",dest)
        self.path = dest

    def copy(self,dest):
        con.nas.copy(self.display_path,dest)

    def convert(self):
        self.original_name = self.name
        name,path,fid,p_link = con.nas.convert_office(self.display_path)
        self.name = name
        self.file_id = fid
        self.permanent_link = p_link

    def rename(self,new_name):
        con.nas.change_name(self.display_path,new_name)
        self.name = new_name

    def download(self,dest = None):
        return con.nas.download_file(self.display_path,dest)

class Note(AttrDict):
    def __init__(self,tp,source,no,flow='in',ref='',date=None,content='',dept='',comments='',year=None):
        self.type = tp
        self.source = source
        self._no = no
        self.ref = ref
        self.date = date if date else datetime.today()
        self.content = content
        self.dept = dept
        self.comments = comments
        self.year = year if year else datetime.today().strftime('%Y')
        self.files = []
        self.permanent_link = ''
        self.folder_id = ''
        self.folder_path = ''
        self.archived = ''
        self.sent_to = ''
        self.flow = flow
   

    @property
    def no(self):
        return int(self._no)

    @no.setter
    def no(self,value):
        self._no = value

    @property
    def key(self):
        if self.flow == 'in':

            if self.type in ['r','ctr']:
                key = f"{self.source}_"
            else:
                key = f"{self.type}_"

            key += f"{self.no}"
        else:
            tp = self.type_from_no

            if tp == 'cg':
                key = f'Aes_{self.no}'
            elif tp == 'asr':
                key = f"cr-asr_{self.no}"
            elif tp == 'ctr':
                key = f"cr_{self.no}"
            elif tp == 'r':
                key = f"Aes-r_{self.no}"
        
        return key
    
    @property
    def type_from_no(self):
        if self.flow == 'out':
            if self.no < 250:
                tp = 'cg'
            elif self.no < 1000:
                tp = 'asr'
            elif self.no < 2000:
                tp = 'ctr'
            else:
                tp = 'r'

        return tp

    @property
    def synology_id(self):
        return self.folder_id if self.folder_id else self.files[0].file_id
    
    @property
    def archive_folder(self):
        if self.flow == 'in':
            return f"{self.type} {self.flow} {self.year}"
        else:
            return f"{self.type_from_no} {self.flow} {self.year}"
   
    @property
    def message(self):
        message = f"Content: `{self.content}` \nLink: {self.messageLink()} \nAssigned to: *{self.dept}*"
        
        if self.ref != '':
            message += f"\nRef: _{self.ref}_"
        
        if self.comments != '':
            message += f"\nComment: _{self.comments}_"
 
        message +=  f"\nRegistry date: {self.date}"

        return message



    @property
    def of_annex(self):
        annex = len(self.files) - 1

        return annex if annex > 0 else ''
 
    def __str__(self):
         return self.key

    def addFile(self,file):
        self.files.append(file)

    def sheetLink(self,text):
        if self.permanent_link:
            return f'=HYPERLINK("#dlink=/d/f/{self.permanent_link}", "{text}")'
        else:
            return self.files[0].getLinkSheet(text)

    def messageLink(self):
        if self.type == 'cg':
            text = f"{self.no}/{self.year[2:]}"
        elif self.type == 'asr':
            text = f"asr {self.no}/{self.year[2:]}"
        else:
            text = f"{self.source} {self.no}/{self.year[2:]}"
        
        if self.permanent_link:
            return f'<https://nas.prome.sg:5001/d/f/{self.permanent_link}|{text}>'
        else:
            return self.files[0].getLinkMessage(text)


    def exportExcel(self):
        return [self.type,self.source,self.sheetLink(self.no),self.year,self.ref,self.date,self.content,self.dept,self.of_annex,self.comments,self.archived,self.sent_to]

    def move(self,dest):
        if self.folder_path != '' and self.folder_path != None:
            con.nas.move(self.folder_path,dest)
        else:
            if self.files:
                self.files[0].move(dest)

