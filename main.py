import win32com.client
import os
from datetime import datetime, timedelta
from time import sleep

class mail_methods():
    def __init__(self, current_items):
        self.current_items = current_items
        self.use_filter = False

    def turn_in(self, to_tuple=False):
        if to_tuple:
            return self.have_filter()
    
    def have_filter(self):
        if self.use_filter:
            return self.filtred_items
        
    def filter(self, subject=None, un_read=None):
        #! Adicionar filtros para remetente
        self.use_filter = True
        self.filtred_items = []

        for msg in self.current_items:
            conditions = []
            
            if subject is not None:
                conditions.append(subject in msg.Subject)
            
            if un_read is not None:
                conditions.append(msg.UnRead == un_read)
            
            # Se não há filtros, adiciona tudo
            if not conditions:
                self.filtred_items.append(msg)
            # Se há filtros, adiciona apenas se TODOS forem verdadeiros (AND logic)
            elif all(conditions):
                self.filtred_items.append(msg)

        return self

    def reply_mails(self, email_text, attachments=False, send=False):
        for msg in self.have_filter():
            if msg.Class == 43:
                reply = msg.ReplyAll()
                reply.Display()

                reply.CC = msg.CC

                # Digitar e sair a assinatura padrão
                awnser = reply.GetInspector.WordEditor.Application.Selection

                print(f'\nWriting E-mail -+- {msg.Subject}', end='')

                awnser.HomeKey(Unit=6)
                awnser.TypeText(email_text)

                if attachments:
                    for attachment in msg.Attachments:
                        if str(attachment.FileName).endswith('.pdf'):
                            temp_dir = os.path.join(os.environ['TEMP'], attachment.FileName)
                            attachment.SaveAsFile(temp_dir)
                            reply.Attachments.Add(temp_dir)

                if send:
                    reply.Send()
                    print(' -+- Sended [OK]')
    
    def un_read(self, turn_to=True):
        for msg in self.have_filter():
            if msg.Class == 43:
                msg.UnRead = turn_to
                print(f'\nMail "{msg.Subject}" has been edited')

class folder_methods():
    """ SUBCLASSE para classe principal """

    def __init__(self, selected_folder):
        self.selected_folder = selected_folder

    def list_items(self):
        """ Extrai os itens da pasta informada, conjunto com o 'select_folder' """
        folder_items = self.selected_folder.Items
        folder_items.Sort("[ReceivedTime]", True)
        return mail_methods(folder_items)
    
    def move_mails_to(self, folder_path: str):
        to_folder = Outlook().select_folder(folder_path).selected_folder

        mails_to_move = [msg for msg in self.list_items()]
        for mail in mails_to_move:
            mail.Move(to_folder)
            print(f'\nMail "{mail.Subject}" has moved')



class Outlook():
    def __init__(self):
        self.outlook = win32com.client.Dispatch('Outlook.Application')

    def select_folder(self, folder_path:str):
        """ Encontra pasta de acordo com a tupla informada, a ordem das pastas deve ser informada de forma decrescente hierarquicamente. """
        path = folder_path.replace('//','').split('/')

        is_first = True
        for folder in path:
            if is_first:
                selected_folder = self.outlook.GetNamespace('MAPI').Folders[folder]
                is_first = False
            else:
                selected_folder = selected_folder.Folders[folder]
        
        print(f'\nFolder "{folder}" has selected')

        return folder_methods(selected_folder)
    
    def write_email(self, mail_text, mail_subject, to_addres, attachments_dir:tuple = [], cc_address = None, send=False, secound_mail_text=None, copy_paste=False):
        mail = self.outlook.CreateItem(0)
        mail.Display()

        inspector = mail.GetInspector

        while inspector.WordEditor is None:
            sleep(.2)

        mail.To = to_addres
        mail.Subject = mail_subject

        writer = inspector.WordEditor.Application.Selection
        writer.HomeKey(Unit=6)
        writer.TypeText(mail_text)

        if copy_paste:
            writer.TypeText('\n\n')
            writer.Paste()
            writer.TypeText('\n')

        if secound_mail_text:
            writer.TypeText(secound_mail_text)

        if cc_address:
            mail.CC = cc_address
        
        if len(attachments_dir) > 0:
            for att in attachments_dir:
                mail.Attachments.Add(att)

        if send:
            mail.Send()
        
        mail = None
        writer = None