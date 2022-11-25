import win32com.client
import pandas as pd

# file = pd.ExcelFile('Vouchers\\track_table\\Emails_Vouchers_0_22-11-2022-1610.xlsx') #Establishes the excel file you wish to import into Pandas
file = pd.ExcelFile('Emails_Vouchers_1_22-11-2022-1610.xlsx')

df = file.parse('Sheet1') #Uploads Sheet1 from the Excel file into a dataframe

for index, row in df.iterrows(): #Loops through each row in the dataframe
    email = (row['email'])  #Sets dataframe variable, 'email' to cells in column 'Email Addresss'
    subject = ('Thank you for your participation in our survey') #Sets dataframe variable, 'subject' to cells in column 'Subject'
    content = str((row['content'])) #Sets dataframe variable, 'body' to cells in column 'Email HTML Body'
    #with the voucher variable

    if (pd.isnull(email) or pd.isnull(subject) or pd.isnull(content)): #Skips over rows where one of the cells in the three main columns is blank
        continue

    olMailItem = 0x0 #Initiates the mail item object
    obj = win32com.client.Dispatch("Outlook.Application") #Initiates the Outlook application
    send_account = None #Sets the send account to none
    for account in obj.Session.Accounts:
        if account.DisplayName == 'humanwildlife@sutd.edu.sg': #Sets the send account to the account with the name '
            send_account = account
            break

    newMail = obj.CreateItem(olMailItem) #Creates an Outlook mail item

    newMail._oleobj_.Invoke(*(64209, 0, 8, 0, send_account)) #Sets the send account to the account with the name

    newMail.Subject = subject #Sets the mail's subject to the 'subject' variable
    newMail.HTMLbody = (content)

    newMail.To = email #Sets the mail's To email address to the 'email' variable
    # newMail.BCC = 'brigid_trenerry@sutd.edu.sg' #Sets the mail's CC email address to the 'email' variable
    #newMail.display() #Displays the mail as a draft email
    
    # To send mail, uncomment below
    newMail.Send()