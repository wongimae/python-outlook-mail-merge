import win32com.client
import pandas as pd

file = pd.ExcelFile('./data.xlsx') #Establishes the excel file you wish to import into Pandas

df = file.parse('Sheet1') #Uploads Sheet1 from the Excel file into a dataframe

for index, row in df.iterrows(): #Loops through each row in the dataframe
    email = (row['email'])  #Sets dataframe variable, 'email' to cells in column 'Email Addresss'
    subject = (row['subject']) #Sets dataframe variable, 'subject' to cells in column 'Subject'
    content = str((row['content'])) #Sets dataframe variable, 'body' to cells in column 'Email HTML Body'
    voucher1 = str((row['voucher_1'])) #Sets dataframe variable, 'voucher' to cells in column 'voucher'
    voucher2 = str((row['voucher_2']))
    voucher3 = str((row['voucher_3']))
    voucher4 = str((row['voucher_4']))
    voucher5 = str((row['voucher_5']))

    if (pd.isnull(email) or pd.isnull(subject) or pd.isnull(content)): #Skips over rows where one of the cells in the three main columns is blank
        continue

    olMailItem = 0x0 #Initiates the mail item object
    obj = win32com.client.Dispatch("Outlook.Application") #Initiates the Outlook application
    newMail = obj.CreateItem(olMailItem) #Creates an Outlook mail item
    newMail.Subject = subject #Sets the mail's subject to the 'subject' variable
    newMail.HTMLbody = (content)
    # print(content)
    # newMail.HTMLbody = ("" + content + "\r\n" + voucher) #Sets the mail's body to 'body' variable
    # print(type(newMail.HTMLbody))
    newMail.HTMLbody = ("<HTML> <p> Dear " 
                + "Sir/Madam" + ","
                + "<br> Thank you for participating in our survey. <br> <br> Please find your vouchers below. There are 5 x $5 vouchers so $25 in total. Please find the links below - you can use the QR code to redeem in-store. <br> <br>"
                + voucher1 +"<br>" + voucher2 + "<br>" + voucher3 + "<br>" + voucher4 + "<br>" + voucher5 + "<br><br>"
                + "IMPORTANT: Please reply to this email to confirm you have received the vouchers. <br> <br>"
                + "Thank you again for your participation in this research. <br> Sincerely, <br> LKYCIC Human Wildlife Team"
                + "</HTML>")

    newMail.To = email #Sets the mail's To email address to the 'email' variable
    newMail.display() #Displays the mail as a draft email
    
    # To send mail, uncomment below
    newMail.Send()
