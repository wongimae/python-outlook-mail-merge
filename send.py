import win32com.client
import pandas as pd

file = pd.ExcelFile('./data.xlsx') #Establishes the excel file you wish to import into Pandas

df = file.parse('Sheet1') #Uploads Sheet1 from the Excel file into a dataframe

for index, row in df.iterrows(): #Loops through each row in the dataframe
    email = (row['email'])  #Sets dataframe variable, 'email' to cells in column 'Email Addresss'
    subject = (row['subject']) #Sets dataframe variable, 'subject' to cells in column 'Subject'
    content = str((row['content'])) #Sets dataframe variable, 'body' to cells in column 'Email HTML Body'
    # voucher = str((row['voucher'])) #Sets dataframe variable, 'voucher' to cells in column 'voucher'

    if (pd.isnull(email) or pd.isnull(subject) or pd.isnull(content)): #Skips over rows where one of the cells in the three main columns is blank
        continue

    olMailItem = 0x0 #Initiates the mail item object
    obj = win32com.client.Dispatch("Outlook.Application") #Initiates the Outlook application
    newMail = obj.CreateItem(olMailItem) #Creates an Outlook mail item
    newMail.Subject = subject #Sets the mail's subject to the 'subject' variable
    newMail.HTMLbody = "(<HTML>content</HTML>)"
    # print(content)
    # newMail.HTMLbody = ("" + content + "\r\n" + voucher) #Sets the mail's body to 'body' variable
    # print(type(newMail.HTMLbody))


    newMail.To = email #Sets the mail's To email address to the 'email' variable
    newMail.display() #Displays the mail as a draft email
    
    # newMail.Send()
