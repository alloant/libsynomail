#!/bin/python

import logging

from libsynomail import EXT
import libsynomail.connection as con


def change_names(groups):
    for group,ctrs in groups.items():
        notes = con.nas.get_file_list(f"/mydrive/ToSend/{group}")
        
        if not notes: continue

        for note in notes:
            if note['name'][0].isdigit():
                con.nas.change_name(f"{note['display_path']}",f"cr{note['name']}")


def send_to_all(groups):
    for group,ctrs in groups.items():
        notes = con.nas.get_file_list(f"/mydrive/ToSend/{group}")
        
        if not notes: continue

        for note in notes:
            for ctr in ctrs.split(","):
                if ctr == ctrs.split(",")[-1]:
                    con.nas.move(note['display_path'],f"/team-folders/Mailbox {ctr}/cr to {ctr}")
                else:
                    con.nas.copy(note['display_path'],f"/team-folders/Mailbox {ctr}/cr to {ctr}")
                    

def init_send_mail(groups):
    logging.info('Starting to send mail to ctr')
    change_names(groups)
    send_to_all(groups)
    logging.info('Finish to send mail to ctr')
