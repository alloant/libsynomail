from pathlib import Path
from attrdict import AttrDict
from libsynomail import EXT, DEBUG

class File(AttrDict):
    def __init__(self,data):
        self.name = data['name']
        self.type = data['type']
        self.display_path = data['display_path']
        self.file_id = data['file_id']
        self.permanent_link = data['permanent_link']

            
    @property
    def ext(self):
        return Path(self.name).suffix[1:]

    @property
    def chain_link(self):
        if self.type == 'dir' or self.ext in EXT:
            return 'd/f'
        else:
            return 'oo/r'

    def getLinkSheet(self,text = None):
        link_text = text if text else self.name
        return f'=HYPERLINK("#dlink=/{self.chain_link}/{self.permanent_link}", "{link_text}")'

    def getLinkMessage(self,text = None):
        link_text = text if text else self.name
        return f'<https://nas.prome.sg:5001/{self.chain_link}/{self.permanent_link}|{link_text}>'

