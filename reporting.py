import os
import re
import pythoncom
import win32com.client
from outlook_report_automation import plt

glob_path = "C:\Projects"


class Handler_Class(object):
    def OnNewMailEx(self, receivedItemsIDs):
        """
        This function is used for reading received emails from outlook inbox folder.
        It checks a subject of the Email and if the subject contains trigger phrase -> extract attachment/save/
        execute gen_report function to generate reports from attached files.
        :param receivedItemsIDs:
        :return: None
        """
        # RecrivedItemIDs is a collection of mail IDs separated by a ",".
        for ID in receivedItemsIDs.split(","):
            mail = outlook.Session.GetItemFromID(ID)
            subject = mail.Subject
            print("getting a new email------")
            try:
                if re.search('grafana reporting', subject.lower()) != None:
                    subject = re.sub('grafana reporting', ' ', subject.lower())
                    subject = re.sub('[^0-9a-zA-Z]+', ' ', subject.lower())
                    subject = re.sub('ext', '', subject.lower())
                    attachments = mail.Attachments
                    if attachments != None:
                        attachment = attachments.Item(1)
                        fullpath = os.path.join(glob_path, 'inputs', attachment.FileName)
                        attachment.SaveAsFile(fullpath)
                        plt.gen_report(fullpath, subject)

            except Exception as e:
                print("Error " + str(e))


outlook = win32com.client.DispatchWithEvents("Outlook.Application", Handler_Class)

pythoncom.PumpMessages()
