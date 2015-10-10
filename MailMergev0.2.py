######################################################################################################
## MailMerge Version 0.2
## 
## Date - 10 Oct 2015
## Authors - Sushant Kudagi, Rohit Shinde, Sudarshan Shanbhag, Sachin Mahendrakar
##
## New Features:
## 	- Added mail server option for yahoo
##	- Added Reset Program Button
##	- Removed obsolete package xlwt
######################################################################################################

#!/usr/bin/python

##Import all require packages
import sys
import gtk                      #GTK Package for GUI
import Tkinter,tkFileDialog     #Package for file open dialog
import smtplib
import xlrd
import re
from time import sleep
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText

##Assign global variables to avoid memory leakages
attachpath = 0                    
namespath = 0
temppath = 0
temppath1 = 0
nicknames = []
emails = []
username = 0
password = 0
subject = 0
message_text = 0
emailcombobox = gtk.combo_box_new_text()
emailcombobox.append_text("Select a Server:")
emailcombobox.append_text("Gmail")
emailcombobox.append_text("Yahoo")
emailcombobox.set_active(0)

#Program Window Declaration & Attributes
win = gtk.Window()
win.set_title("MailMerge Beta")
win.connect("destroy", gtk.main_quit)

#Function to choose a file from a dialog box
def mail(to, subject, text, usern, pwdd):
    msg = MIMEMultipart()
    
    msg['From'] = usern
    msg['To'] = to
    msg['Subject'] = subject
    
    msg.attach(MIMEText(text))
    
    tempmail = emailcombobox.get_active()
    if(tempmail == 1):
        mailServer = smtplib.SMTP("smtp.gmail.com", 587)
    elif(tempmail == 2):
        mailServer = smtplib.SMTP("smtp.mail.yahoo.com", 587)
    mailServer.ehlo()
    mailServer.starttls()
    mailServer.ehlo()
    mailServer.login(usern, pwdd)
    mailServer.sendmail(usern, to, msg.as_string())
    mailServer.close()
    print "Email sent to %s" % to
    
#Definition to choose a txt Message file
def on_clicked(temp_button):
        global temppath
        global message_text
        root = Tkinter.Tk();
        root.update()
        root.withdraw()
        temppath= tkFileDialog.askopenfilename(filetypes=[("Text files","*.txt")])
        f = open(temppath, "r")
        message_text = ''.join(f.readlines())
        f.close()

#Definition to choose a contact xls or xlsx file
def on_clicked1(temp1_button):
        global temppath1
        root = Tkinter.Tk();root.withdraw()
        temppath1 = tkFileDialog.askopenfilename(filetypes=[("Excel 2003","*.xls"),("Excel","*.xlsx")])

#Definition for Reset Button
def on_reset(reset_button):
    labelfinal.set_text("Program Reset")
    emailcombobox.set_active(0)
    input23.set_text("")
    input24.set_text("")
    inputsubject.set_text("")    

#Definition for Send Email button
def on_clicke(email_button):
    labelfinal.set_text("Sending Mails")
    username = input23.get_text()
    password = input24.get_text()
    
    if(emailcombobox.get_active() == 1):
        server = smtplib.SMTP("smtp.gmail.com:587")
    elif(emailcombobox.get_active() == 2):
        server = smtplib.SMTP("smtp.mail.yahoo.com:587")
    server.starttls()
    try:
        server.login(username, password)
        server.close()
    except smtplib.SMTPAuthenticationError:
        labelfinal.set_text("Email Authentication Error")
    finally:
        server.close()

      
    subject = inputsubject.get_text()
    #open up the workbook
    book = xlrd.open_workbook(temppath1)
    sheet = book.sheet_by_index(0) 
 
    #check all the rows
    for row_number in xrange(1, sheet.nrows):
        nickname = sheet.cell_value(row_number, 0)
        email = sheet.cell_value(row_number, 1)
 
        if nickname and email and re.match(r"[^@]+@[^@]+\.[^@]+", email):     #Email regex validation
            nicknames.append(nickname)
            emails.append(email)
            
    
    counter = 0
    for email in emails:
            GREETING = "Dear %s,\n\n" % nicknames[counter]
            mail(email,
            subject,
            GREETING+message_text,
            username,
            password)
            counter += 1
            sleep(1)
    
    labelfinal.set_text("All Mails Sent")
           
#Declaration and addition of GUI elements	
vbox=gtk.VBox()
vbox1=gtk.VBox()
vbox2=gtk.VBox()
hbox1=gtk.HBox();hbox2=gtk.HBox();hbox3=gtk.HBox();
hbox11=gtk.HBox();hbox12=gtk.HBox();hbox13=gtk.HBox();

temp_button=gtk.Button("Choose Template")
temp1_button=gtk.Button("Choose Email list")
email_button=gtk.Button("Email all")
reset_button=gtk.Button("Reset")

label=gtk.Label("MailMerge v0.2")
label11=gtk.Label("Mail Settings")
label12=gtk.Label("Template file")
label13=gtk.Label("Name list file")
label21=gtk.Label("Account Login Details")
label23=gtk.Label("User Name")
label24=gtk.Label("Password")
labelsubject=gtk.Label("Subject")
label25=gtk.Label("")
labelfinal=gtk.Label("")

input23=gtk.Entry()
input24=gtk.Entry()
input24.set_visibility(False)
inputsubject=gtk.Entry()

vbox.add(label);

vbox1.add(label11);hbox11.add(label12);hbox11.add(temp_button);vbox1.add(hbox11);hbox12.add(label13);hbox12.add(temp1_button);vbox1.add(hbox12);hbox13.add(labelsubject);hbox13.add(inputsubject);vbox1.add(hbox13);vbox1.add(email_button);
vbox2.add(label21);vbox2.add(emailcombobox);hbox1.add(label23);hbox1.add(input23);hbox2.add(label24);hbox2.add(input24);
vbox2.add(hbox1);vbox2.add(hbox2);vbox2.add(label25);vbox2.add(reset_button);

hbox3.add(vbox1);hbox3.add(vbox2);

vbox.add(hbox3);
vbox.add(labelfinal);


email_button.connect("clicked",on_clicke)
temp_button.connect("clicked",on_clicked)
temp1_button.connect("clicked",on_clicked1)
reset_button.connect("clicked",on_reset)

win.add(vbox)
win.show_all()
gtk.main()

##End of Program

