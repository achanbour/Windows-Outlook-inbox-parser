import win32com.client as win32
from datetime import datetime
import time
import os
import ctypes
import traceback

"""
Create a mapping between the full month name and month keyword
"""
months_map = {
    "January": "Jan",
    "February": "Feb",
    "March": "Mar",
    "April": "Apr",
    "May": "May",
    "June": "Jun",
    "July": "Jul",
    "August": "Aug",
    "September": "Sep",
    "October": "Oct",
    "November": "Nov",
    "December": "Dec",
}

"""
Read the mailing list as a pandas dataframe
"""
mailing_list = "mailing_list.xlsx"

mailing_list_df = pd.read_excel(mailing_list)

# Fill the any NA values in the mailing list dataframe
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
            folder = namespace.Folders.Item("myemail@test.com")
            inbox = folder.Folders.Item("Inbox")  # Main inbox
            messages = inbox.Items  # Get the inbox items

            for message in messages:

                """
                Step I: Iterate over unread emails with subject containing the defined keyword: test
                """
                if (message.Unread == True) and ("test" in message.Subject):

                    subject = message.Subject
                    print("Email found!")
                    print(f"Subject: {subject}")

                    if (message.SenderEmailType == "EX"):
                        sender = message.Sender.GetExchangeUser().PrimarySmtpAddress
                    else:  # email type SMTP
                        sender = message.SenderEmailAddress

                    date_and_time = message.SentOn.strftime("%Y-%m-%d--%H-%M")
                    print(f"Date & time received: {date_and_time}")
                    print(f"Sender: {sender}")

                    """
                    Step II: Archive the email and any attachments in the defined archive location
                    """
                    # Create year folder in the Emails Received directory
                    archive_location = "C:\Users\user\Desktop"
                  
                    # Eliminate any special characters in the name
                    new_subject = str(date_and_time) + " " + \
                        subject.replace("RE:", "")
                  
                    name = new_subject + ".msg"
                    message.SaveAs(archive_location + "\\" + name)
                    message.unread = False  # set the email as Read

                    """
                    Step III: Archive the attachments in the email
                    """

                    # Get the current timestamp and include it in the attachment file name when saving

                    today = datetime.today().strftime("%Y-%m-%d--%H-%M")

                    for att in message.Attachments:
                        file = today + att.FileName
                        att.SaveAsFile(archive_location + "/" + file)

                    """
                    Step IV: Send a notification informing that an email has been detected only after the archive is complete
                    """
                    recipients = "recipient@test.com"
                    notification = outlook.CreateItem(0)
                    notification.Subject = "Notification subject"
                    notification.Body = "Hello,\n\nThis is an automated notification informing you that an email with \"test\" in subject has been detected. This email has been successfully archived in the defined target location"
                    notification.To = recipients
                    notification.SentOnBehalfOfName = "myemail@test.com"
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
