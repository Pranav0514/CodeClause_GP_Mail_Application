from tkinter import *
from tkinter import messagebox, filedialog
import pandas
from pygame import mixer
import speech_recognition
from email.message import EmailMessage
import smtplib
import os
import imghdr
check = False

def browse():
    global final_emails
    path=filedialog.askopenfilename(initialdir='c/:',title='Select Excel File')
    if path=='':
        messagebox.showerror('Error','Please Select An Excel File')

    else:
        data=pandas.read_excel(path)
        if 'Email' in data.columns:
            emails=list(data['Email'])
            final_emails=[]
            for i in emails:
                if pandas.isnull(i)==False:
                    final_emails.append(i)

            if len(final_emails)==0:
                messagebox.showerror('Error','File Does Not Contain Any Mail Addresses')
            else:
                toEntryField.config(state=NORMAL)
                toEntryField.insert(0,os.path.basename(path))
                toEntryField.config(state='readonly')
                totalLable.config(text='Total: '+str(len(final_emails)))
                sentLable.config(text='Sent: ')
                leftLable.config(text='Left: ')
                failedLable.config(text='Failed: ')

def button_check():
    if choice.get()=='multiple':
        browseButton.config(state=NORMAL)
        toEntryField.config(state='readonly')

    if choice.get()=='single':
        browseButton.config(state=DISABLED)
        toEntryField.config(state=NORMAL)

def attachment():
    global filename, filetype, filepath, check
    check = True

    filepath = filedialog.askopenfilename(initialdir='c:/', title='Select Attachment')
    filetype = filepath.split('.')
    filetype = filetype[1]
    filename = os.path.basename(filepath)
    textarea.insert(END, f'\n{filename}\n')


def sendingEmail(toAddress, subject, body):
    f = open('credentials.txt', 'r')
    for i in f:
        credentials = i.split(',')

    message = EmailMessage()
    message['subject'] = subject
    message['to'] = toAddress
    message['from'] = credentials[0]
    message.set_content(body)
    if check:
        if filetype == 'png' or filetype == 'jpg' or filetype == 'jpeg':
            f = open(filepath, 'rb')
            file_data = f.read()
            subtype = imghdr.what(filepath)

        else:
            f = open(filepath, 'rb')
            file_data = f.read()
            message.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=filename)

    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(credentials[0], credentials[1])
    s.send_message(message)
    x = s.ehlo()
    if x[0] == 250:
        return 'Sent'
    else:
        return 'Failed'



def send_email():
    if toEntryField.get() == '' or subjectEntryField.get() == '' or textarea.get(1.0, END) == '\n':
        messagebox.showerror('Error', 'All Fields Are Required', parent=root)

    else:
        if choice.get() == 'single':
            result=sendingEmail(toEntryField.get(), subjectEntryField.get(), textarea.get(1.0, END))
            if result=='Sent':
                messagebox.showinfo('Success', 'Mail Has Been Sent Successfully')

            if result=='Failed':
                messagebox.showerror('Failed','Failed To Send Mail')

        if choice.get() == 'multiple':
            sent=0
            failed=0
            for x in final_emails:
                result=sendingEmail(x,subjectEntryField.get(), textarea.get(1.0, END))
                if result=='Sent':
                    sent+=1
                if result=='Failed':
                    failed+=1

                totalLable.config(text='')
                sentLable.config(text='Sent: '+str(sent))
                leftLable.config(text='Left: '+str(len(final_emails)-(sent+failed)))
                failedLable.config(text='Failed: '+str(failed))

                totalLable.update()
                sentLable.update()
                leftLable.update()
                failedLable.update()

            messagebox.showinfo('Success','Mails Have Been Sent Successfully.')


def settings():
    def clear1():
        fromEntryField.delete(0, END)
        passwordEntryField.delete(0, END)

    def save():
        if fromEntryField.get() == '' or passwordEntryField.get() == '':
            messagebox.showerror('Error', 'All Fields Are Required', parent=root1)

        else:
            f = open('credentials.txt', 'w')
            f.write(fromEntryField.get() + ',' + passwordEntryField.get())
            f.close()
            messagebox.showinfo('Information', 'Credentials Stored Successfully', parent=root1)

    root1 = Toplevel()
    root1.title('Setting')
    root1.geometry('650x340+350+90')
    root1.config(bg='slategray4')

    Label(root1, text='Credentials', image=logoImage, compound=LEFT, font=('Goudy Old Style', 40, 'bold'),
          fg='white', bg='slategray4').grid(padx=150)
    fromLabelFrame = LabelFrame(root1, text='From (Mail Id)', font=('times new roman', 16, 'bold'), bd=5, fg='white',
                                bg='slategray4')
    fromLabelFrame.grid(row=1, column=0, pady=20)
    fromEntryField = Entry(fromLabelFrame, font=('times new roman', 18, 'bold'), width=30)
    fromEntryField.grid(row=0, column=0)

    passwordLabelFrame = LabelFrame(root1, text='Password', font=('times new roman', 16, 'bold'), bd=5, fg='white',
                                    bg='slategray4')
    passwordLabelFrame.grid(row=2, column=0, pady=20)
    passwordEntryField = Entry(passwordLabelFrame, font=('times new roman', 18, 'bold'), width=30, show='*')
    passwordEntryField.grid(row=0, column=0)

    Button(root1, text='Save', font=('times new roman', 18, 'bold'), cursor='hand2', bg='ghost white', fg='black',
           command=save).place(x=210, y=280)
    Button(root1, text='Clear', font=('times new roman', 18, 'bold'), cursor='hand2', bg='ghost white', fg='black',
           command=clear1).place(x=340, y=280)

    f = open('credentials.txt', 'r')
    for i in f:
        credentials = i.split(',')

    fromEntryField.insert(0, credentials[0])
    passwordEntryField.insert(0, credentials[1])

    root1.mainloop()


def speak():
    mixer.init()
    mixer.music.load('music1.mp3')
    mixer.music.play()
    sr = speech_recognition.Recognizer()
    with speech_recognition.Microphone() as m:
        try:
            sr.adjust_for_ambient_noise(m, duration=0.2)
            audio = sr.listen(m)
            text = sr.recognize_google(audio)
            textarea.insert(END, text + '.')

        except:
            pass


def iexit():
    result = messagebox.askyesno('Notification', 'Confirm to exit?')
    if result:
        root.destroy()
    else:
        pass


def clear():
    toEntryField.delete(0, END)
    subjectEntryField.delete(0, END)
    textarea.delete(1.0, END)


root = Tk()
root.title('MailApplication')
root.geometry('780x620+100+50')
root.resizable(0, 0)
root.config(bg='slategray4')

titleFrame = Frame(root, bg='white')
titleFrame.grid(row=0, column=0)
logoImage = PhotoImage(file='email.png')
titleLabel = Label(titleFrame, text='  Mail App', image=logoImage, compound=LEFT, font=('Goudy Old Style', 28, 'bold'),
                   bg='white', fg='grey11')
titleLabel.grid(row=0, column=0)
settingImage = PhotoImage(file='setting.png')

Button(titleFrame, image=settingImage, bd=0, bg='white', cursor='hand2', activebackground='white'
       , command=settings).grid(row=0, column=1, padx=20)

chooseFrame = Frame(root, bg='slategray4')
chooseFrame.grid(row=1, column=0, pady=10)
choice = StringVar()

singleRadioButton = Radiobutton(chooseFrame, text='Single', font=('times new roman', 20, 'bold'),
                                variable=choice, value='single', bg='slategray4', activebackground='slategray4',
                                  command=button_check)
singleRadioButton.grid(row=0, column=0, padx=20)
multipleRadioButton = Radiobutton(chooseFrame, text='Multiple', font=('times new roman', 20, 'bold'),
                                  variable=choice, value='multiple', bg='slategray4', activebackground='slategray4',
                                  command=button_check)
multipleRadioButton.grid(row=0, column=1, padx=20)
choice.set('single')

toLabelFrame = LabelFrame(root, text='To (Mail Id)', font=('times new roman', 16, 'bold'), bd=5, fg='white',
                          bg='slategray4')
toLabelFrame.grid(row=2, column=0, padx=100)

toEntryField = Entry(toLabelFrame, font=('times new roman', 18, 'bold'), width=30)
toEntryField.grid(row=0, column=0)

browseImage = PhotoImage(file='browse.png')

browseButton=Button(toLabelFrame, text='Browse', image=browseImage, compound=LEFT, font=('arial', 12, 'bold'),
       cursor='hand2', bd=0, bg='slategray4', activebackground='slategray4', state=DISABLED,command=browse)
browseButton.grid(row=0, column=1,padx=20)

subjectLabelFrame = LabelFrame(root, text='Subject', font=('times new roman', 16, 'bold'), bd=5, fg='white',
                               bg='slategray4')
subjectLabelFrame.grid(row=3, column=0, pady=10)

subjectEntryField = Entry(subjectLabelFrame, font=('times new roman', 18, 'bold'), width=30)
subjectEntryField.grid(row=0, column=0)

emailLabelFrame = LabelFrame(root, text='Compose', font=('times new roman', 16, 'bold'), bd=5, fg='white',
                             bg='slategray4')
emailLabelFrame.grid(row=4, column=0, padx=20)

micImage = PhotoImage(file='mic.png')
Button(emailLabelFrame, text=' Speak', image=micImage, compound=LEFT, font=('arial', 12, 'bold'),
       cursor='hand2', bd=0, bg='slategray4', activebackground='slategray4', command=speak).grid(row=0, column=0)
attachImage = PhotoImage(file='attachments.png')
Button(emailLabelFrame, text=' Attach', image=attachImage, compound=LEFT, font=('arial', 12, 'bold'),
       cursor='hand2', bd=0, bg='slategray4', activebackground='slategray4', command=attachment).grid(row=0, column=1)

textarea = Text(emailLabelFrame, font=('times new roman', 14,), height=8)
textarea.grid(row=1, column=0, columnspan=2)

sendImage = PhotoImage(file='send.png')
Button(root, image=sendImage, bd=0, bg='slategray4', cursor='hand2', activebackground='slategray4',
       command=send_email).place(x=490, y=540)

clearImage = PhotoImage(file='clear.png')
Button(root, image=clearImage, bd=0, bg='slategray4', cursor='hand2', activebackground='slategray4'
       , command=clear).place(x=590, y=550)

exitImage = PhotoImage(file='exit.png')
Button(root, image=exitImage, bd=0, bg='slategray4', cursor='hand2', activebackground='slategray4'
       , command=iexit).place(x=690, y=550)

totalLable = Label(root, font=('times new roman', 18, 'bold'), bg='slategray4', fg='white')
totalLable.place(x=10, y=560)
sentLable = Label(root, font=('times new roman', 18, 'bold'), bg='slategray4', fg='white')
sentLable.place(x=100, y=560)
leftLable = Label(root, font=('times new roman', 18, 'bold'), bg='slategray4', fg='white')
leftLable.place(x=190, y=560)
failedLable = Label(root, font=('times new roman', 18, 'bold'), bg='slategray4', fg='white')
failedLable.place(x=280, y=560)

root.mainloop()
