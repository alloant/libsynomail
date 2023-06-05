from synology_drive_api.drive import SynologyDrive
from synochat.webhooks import IncomingWebhook

from tempfile import NamedTemporaryFile
import ast

import time
import logging
from pathlib import Path

from synomail import CONFIG, EXT, INV_EXT

class prome:
    def __init__(self,PASS=None):
        self.PASS = PASS

    def get_team_folders(self):
        try:
            with SynologyDrive(CONFIG['user'],self.PASS,"nas.prome.sg",dsm_version='7') as synd:
                return synd.get_teamfolder_info()
            return None
        except Exception as err:
            logging.error(err)
            logging.warning(f'Cannot get team folders')

    def get_file_list(self,folder):
        try:
            with SynologyDrive(CONFIG['user'],self.PASS,"nas.prome.sg",dsm_version='7') as synd:
                return synd.list_folder(folder)['data']['items']
            return None
        except Exception as err:
            logging.error(err)
            logging.warning(f'Cannot get files from {folder}')


    def change_name(self,file,new_name):
        try:
            with SynologyDrive(CONFIG['user'],self.PASS,"nas.prome.sg",dsm_version='7') as synd:
                synd.rename_path(new_name,file)
                return True
        except Exception as err:
            logging.error(err)
            logging.warning(f'Cannot change name of {file}')
            return False



    def move(self,path,new_path):
        try:
            with SynologyDrive(CONFIG['user'],self.PASS,"nas.prome.sg",dsm_version='7') as synd:
                logging.debug(f'Sending synology command to move {path}')
                rst = synd.move_path(path,new_path)
                logging.debug('Command to move sent')

                task_id = rst['data']['async_task_id']
        
                logging.debug("Waiting for synology to move")
                rst = synd.get_task_status(task_id)
        
                while(rst['data']['result'][0]['data']['progress'] < 100 or rst['data']['has_fail']):
                    time.sleep(0.2)
                    rst = synd.get_task_status(task_id)

                logging.debug("{path} was moved to {new_path}")

                rst_data = rst['data']['result'][0]['data']['result']
        
                if not 'targets' in rst_data:
                    logging.error(f'Synology cannot move the file {path} to {new_path}')
                    return False
                return True
        except Exception as err:
            logging.error(err)
            logging.warning(f'Cannot move the file {path} to {new_path}')
            return False

    def copy(self,path,dest):
        try:
            with SynologyDrive(CONFIG['user'],self.PASS,"nas.prome.sg",dsm_version='7') as synd:
                ext = Path(path).suffix[1:]
                if ext in INV_EXT:
                    synd.copy(path,dest)
                else:
                    tmp_file = synd.download_file(path)
                    synd.upload_file(tmp_file,dest)
        except Exception as err:
            logging.error(err)
            logging.error(f"Cannot copy file {path} to {dest}")


    def convert_office(self,file_path,delete = False):
        logging.debug(f"Converting {file_path}...")        
        try:
            with SynologyDrive(CONFIG['user'],self.PASS,"nas.prome.sg",dsm_version='7') as synd:
                rst = synd.convert_to_online_office(file_path,delete_original_file=delete)
                task_id = rst['data']['async_task_id']
    
                rst = synd.get_task_status(task_id)
                while(not rst['data']['has_fail'] and rst['data']['result'][0]['data']['status'] == 'in_progress'):
                    time.sleep(1)
                    rst = synd.get_task_status(task_id)
                
                ext = Path(file_path).suffix[1:]
                name = file_path.replace(ext,EXT[ext])
    
                new_file = synd.get_file_or_folder_info(name)
                file_id = new_file['data']['file_id']
                permanent_link = new_file['data']['permanent_link']
                file_path = new_file['data']['display_path'] 

                return file_path,file_id,permanent_link

        except Exception as err:
            logging.error(err)
            logging.warning(f'Cannot convert {file_path}')
            return '','',''


    def download_file(self,file_path,dest=None):
        logging.debug(f"Downloading {file_path}")
        try:
            with SynologyDrive(CONFIG['user'],self.PASS,"nas.prome.sg",dsm_version='7') as synd:                
                ext = Path(file_path).suffix[1:]
                if ext in INV_EXT:
                    ext = INV_EXT[ext]
                    bio = synd.download_synology_office_file(file_path)
                else:
                    bio = synd.download_file(file_path)
                
                if dest:
                    file_name = Path(file_path).stem
        
                    with open(f'{dest}/{file_name}.{ext}','wb') as f:
                        f.write(bio.read())
                else:
                    return bio
        
        except Exception as err:
            raise
            logging.error(err)
            logging.error(f"Cannot download {file_path}")


    def upload_file(self,file,dest):
        try:
            with SynologyDrive(CONFIG['user'],self.PASS,"nas.prome.sg",dsm_version='7') as synd: 
                logging.debug(f"Uploading {file.name}")
                ret_upload = synd.upload_file(file, dest_folder_path=dest)
        except Exception as err:
            logging.error(err)
            logging.error("Cannot upload {file.name}")


    def upload_convert_wb(self,wb,name,dest):
        try:
            with SynologyDrive(CONFIG['user'],self.PASS,"nas.prome.sg",dsm_version='7') as synd: 
                file = NamedTemporaryFile()
                wb.save(file)
                file.seek(0)
                file.name = name

                logging.debug(f"Uploading {file}")
                ret_upload = synd.upload_file(file, dest_folder_path=dest)
                uploaded = True
        except Exception as err:
            logging.error(err)
            logging.error("Cannot upload register")
            wb.save(file.name)
            uploaded = False

        if uploaded:
            try:
                file_path,file_id,permanent_link = self.convert_office(ret_upload['data']['display_path'],delete=False)
                #with SynologyDrive(CONFIG['user'],self.PASS,"nas.prome.sg",dsm_version='7') as synd: 
                #    ret_convert = synd.convert_to_online_office(ret_upload['data']['display_path'],
                #    delete_original_file=False,
                #    conflict_action='autorename')

                return file_path,file_id,permanent_link
            except Exception as err:
                logging.error(err)
                logging.warning("Cannot convert register to Synology Office")


    def create_folder(self,path,folder):
        files = self.get_file_list(path)
        folder_exists = False
        for fl in files:
            if fl['name'] == folder:
                folder_exists = True
        
        try:
            with SynologyDrive(CONFIG['user'],self.PASS,"nas.prome.sg",dsm_version='7') as synd: 
                if folder_exists:
                    p_link = synd.get_file_or_folder_info(f"{path}/{folder}")['data']['permanent_link']
                else:
                    #path = path[1:] if path[0] == "/" else path
                    rst = synd.create_folder(folder,path)
                    #rst = synd.create_folder(f"{path}/{folder}")

                    p_link = rst['data']['permanent_link']
        except Exception as err:
            logging.error(err)
            logging.error(f"Cannot create folder {path}/{folder}")
            return '',''

        return p_link


    def send_message(self,dep,message):
        tokens = ast.literal_eval(CONFIG['tokens'])
        try:
            webhook = IncomingWebhook('nas.prome.sg', tokens[dep], port=5001)
            webhook.send(message)
            return 1
        except Exception as err:
            raise
            logging.error(err)
            logging.error(f"Cannot send message to {dep}")
            return 0
