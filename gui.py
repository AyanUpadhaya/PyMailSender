#mail sender graphical app by Ayan
#use creds.py file to store your credentials like email and password
#this app is for gmail smtp server, 
#so make sure turn on less secure apps: https://myaccount.google.com/lesssecureapps
#if you have two step authenticator than you might need an app password. Use the app password as your password in creds file.
#user can send only one file at a time, so the script is under development mode
#but can be used to send plain text or html decorated emails
#filetype supported for attachment: doc,xlsx,pdf,txt,jpeg,jpg,png,gif
#contact: ayanu881@gmail.com

from tkinter import*  
import smtplib,ssl,email
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tkinter import filedialog as fd
import time
from creds import*

smtp_server="smtp.gmail.com"
port=587 
EMAIL_ADDRESS=sender_email
EMAIL_PASSWORD=password
context=ssl.create_default_context()

#root decoration
root=Tk()
root.title("PyMailSender GUI Application")
root.geometry('640x460')
root.configure(background="#2f3232")
root.resizable(width=0, height=0)

#global variable
attachments=[]
filename=''
filetype=''

"""---BACKEND-START---"""

#file attachment
def file_attach():
    file_name=fd.askopenfilename(initialdir="/home/ayan/Desktop", title="Select File")
    attachments.append(file_name)
    print(file_name)
    report.config(text="Attached "+str(len(attachments))+" File")


#mail sending
def mailsend():
    reciever_email=send_to.get(1.0,END)
    message_subject=subject_box.get(1.0,END)
    msg=msg_text.get(1.0,END)

    message=MIMEMultipart("alternative")
    message['Subject']=message_subject
    message['From']=sender_email
    message['To']=reciever_email
    part1 = MIMEText(msg, "html")


    if len(attachments)>0:
        filename=attachments[0]
        filetype=attachments[0].split('.')[-1]
        finalname=filename.split('/')[-1]

        #for document object
        if filetype=='pdf' or filetype=='txt' or filetype=='doc' or filetype=="py" or filetype=="xlsx":
            with open(filename,'rb') as f:
                part2=MIMEBase("application","octet-stream")
                part2.set_payload(f.read())
            encoders.encode_base64(part2)
            part2.add_header("Content-Disposition",f"attachment;filename={finalname}")
            message.attach(part2)

        #for image object
        if filetype=='jpeg' or filetype=='jpg' or filetype=='png' or filetype=="gif":
            with open(filename,'rb') as f:
                image=MIMEBase("image",filetype,filename=filename)
                image.set_payload(f.read())
            encoders.encode_base64(image)
            image.add_header("Content-Disposition",f"attachment;filename={finalname}")
            message.attach(image)

    message.attach(part1)
    

    try:
        server=smtplib.SMTP(smtp_server,port,timeout=10)
        server.ehlo()
        server.starttls(context=context)
        server.ehlo()
        server.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
        server.sendmail(sender_email,reciever_email,message.as_string())
        report.configure(text="Mail Has been sent")
        root.update()
        time.sleep(3)
        report.configure(text='')
        root.update()
    except Exception as e:
        report.configure(text=e)
        root.update()
        time.sleep(6)
    finally:
        server.quit()
"""---BACKEND-ENDS--"""



"""---FRONTEND-START--"""
#frane
first_frame=Frame(root,width=500,height=500,bg="yellow", pady=20,padx=20)

#heading 
heading=Label(first_frame,text="PyMailSender",font=('Helvetica' ,16 ,"bold"),bg='yellow',fg="#2f3232")
heading.grid(row=1,column=2)

#tolabel
tolabel=Label(first_frame,text="To",bg="yellow", fg="#2f3232",font=('Helvetica',15,'bold'))
tolabel.grid(row=3,columnspan=1)

#to:
send_to=Text(first_frame,width=50,height=1,bg="white",fg="black")
send_to.grid(row=3,column=2,padx=5,pady=10)

#subjectlabel
subjectlabel=Label(first_frame,text="Subject",bg="yellow", fg="#2f3232",font=('Helvetica',15,'bold'))
subjectlabel.grid(row=4,columnspan=1,pady=10)

#subject:
subject_box=Text(first_frame,width=50,height=1,bg="white",fg="black")
subject_box.grid(row=4,column=2,pady=10,padx=5)

#msglabel
msglabel=Label(first_frame,text="Message",bg="yellow", fg="#2f3232",font=('Helvetica',15,'bold'))
msglabel.grid(row=5,columnspan=1,padx=5)

#msgbox

msg_text=Text(first_frame,width=50,height=10,bg="white",fg="black")
msg_text.grid(row=5,column=2,pady=20)

#attach a file and send_mail button

send_mail=Button(first_frame,text="Send Email",command=mailsend)
send_mail.grid(row=6,column=2)
attach_file=Button(first_frame,text="Attach File",command=file_attach)
attach_file.grid(row=0,column=3)

first_frame.pack()

#notification
report=Label(root,fg="white")
report.pack(side=LEFT)

"""---FRONTEND-ENDS---"""

#mainloop
root.mainloop()