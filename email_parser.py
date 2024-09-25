import win32com.client as win32
from datetime import datetime
import time
import os
import ctypes
import traceback

"""
Read the mailing list as a pandas dataframe
"""
mailing_list = "mailing_list.xlsx"

mailing_list_df = pd.read_excel(mailing_list)

# Fill the any NA values in the mailing list
mailing_list_df = mailing_list_df.fillna(method="ffill", axis=0)

# Define custom Exception raised when Outlook is shut
class ClosedOutlookException(Exception):
    "Raised when the Outlook application is closed"
    def __init__(self):
        print(
            "Outlook may not be running. Starting Outlook and resuming the background process...")
        print("\n")
        self.handle_exception()

    def handle_exception(self):
        monitor_outlook_inbox()  # Restart parser


def monitor_outlook_inbox():
    print(
        "===================================================================================\n" +
        "KRI Email Inbox Parser started!\n" +
        "==================================================================================="
    )
    print("\n")
    try:
        while True:  # keeps the program running until termination

            # At start of execution, check if Outlook is open, and if not, start the process
            if not outlook_is_running():
                os.startfile("outlook")

            outlook = win32.Dispatch("Outlook.Application")  # Outlook app
            namespace = outlook.GetNameSpace("MAPI")  # Outlook API
            folder = namespace.Folders.Item("example@example.com") # Outlook email folder to be monitored
            inbox = folder.Folders.Item("Inbox")  # Main inbox
            messages = inbox.Items  # Get the inbox items

            for message in messages:

                """
                Step I: Iterate over unread emails with subject containing the defined keyword: Appollo
                """
                if (message.Unread == True) and ("Appollo" in message.Subject):

                    subject = message.Subject # Retrieve the subject of the email
                    print("Email found!")
                    print(f"Subject: {subject}")

                    if (message.SenderEmailType == "EX"):
                        sender = message.Sender.GetExchangeUser().PrimarySmtpAddress
                    else:  # email type SMTP
                        sender = message.SenderEmailAddress

                    date_and_time = message.SentOn.strftime("%Y-%m-%d--%H-%M") # Retrieve the date and time the email was sent
                    print(f"Date & time received: {date_and_time}")
                    print(f"Sender: {sender}")

                    """
                    Step II: Archive the email and any attachments in the defined archive location
                    """
                    # Define the path to the archiving location
                    archive_location = "C:\Users\user\Desktop"
                  
                    # Eliminate any special characters in the email subject
                    new_subject = str(date_and_time) + " " + subject.replace("RE:", "")
                    filename = new_subject + ".msg"
                    message.SaveAs(os.path.join(archive_location,filename))
                    message.unread = False  # set the email as Read

                    """
                    Step III: Archive the attachments in the email
                    """

                    # Get the current timestamp and include it in the attachment file name when saving
                    today = datetime.today().strftime("%Y-%m-%d--%H-%M")
                    
                    for att in message.Attachments:
                        attachement_filename = today + att.FileName
                        att.SaveAsFile(os.path.join(archive_location, attachement_filename))

                    """
                    Step IV: Send a notification informing that an email has been detected. This step is executed only after the email has been successfully archived
                    """
                    recipients = "recipient@example.com"
                    notification = outlook.CreateItem(0)
                    notification.Subject = "Notification subject"
                    notification.Body = "Hello,\n\nThis is an automated notification informing you that an email with \"Appollo\" in subject title has been detected. This email has been successfully archived in the defined target folder"
                    notification.To = recipients
                    notification.SentOnBehalfOfName = "example@example.com"
                    notification.Send()
                    print("Notification successfully sent!")
                    print("\n")
            # Halt for 5 seconds to handle the case where multiple emails arrive at the same time
            time.sleep(5)
    except Exception as e:
        error_message = str(e)
        exception_type = e.__class__.__name__
        print(exception_type)
        if "<unknown>.Unread" in error_message or "Outlook.Application.GetNameSpace" in error_message or exception_type == "com_error":
            raise ClosedOutlookException from None
        else:
            traceback.print_exc()
            ctypes.windll.user32.MessageBoxW(
                0, "An error has occured. Please check the terminal window for more information and restart the program", "ERROR", 1)
            # keep the terminal window open
            input("Press enter to close the program")


"""
Helper functions
"""

# Method that checks whether outlook app is running
def outlook_is_running():
    import win32ui
    time.sleep(3)  # pauses for 3 seconds
    try:
        win32ui.FindWindow(None, "Microsoft Outlook")
        return True
    except win32ui.error:
        return False


monitor_outlook_inbox()
