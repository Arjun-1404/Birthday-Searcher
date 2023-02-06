import os
from tkinter import *
from tkinter import messagebox, filedialog, simpledialog
import pandas as pd
from datetime import date
import smtplib
from email.message import EmailMessage

win = Tk()
win.title("Birthday Searcher")
win.geometry("1150x650")
win.resizable(False, False)

bg = PhotoImage(file="img.gif")
C1 = Canvas(win, width=1150, height=978)
C1.place(x=0, y=0)
C1.create_image(433, 200, image=bg)


C2 = Canvas(win, width=240, height=290)
C2.place(x=10, y=10)
C2.create_image(305, 7, image=bg)
C2.create_text(120, 148, text="\t**NOTE**\n\nPlace the excel sheet in\nC drive by creating a\nBirthday folder in it.\n"
               "Name the excel sheet\nBirthday only. \nAlso the file\nshouldn't be open in\nbackground, "
               "else\nfile won't be read by\napplication",
               font=("commissars", 15, "italic", "bold"))


t1 = Text(win, font=("calibre", 14, "bold" ), height=12, width=38)
t1.place(x=270, y=10)
r = "Name the columns of excel as below only: \n\nBirthdate : DOB\nCollege name : College\nEmail : Email" \
    "\nName : Name\nMobile number : Contact\nDepartment : Dept\n" \
    "\nTake care of spaces while naming the\ncolumns no space should be given\nbefore or after the column name in excel"
t1.insert(END, r)

t2 = Text(win, font=("commissars", 14), height=12, width=38)
t2.place(x=710, y=10)
u = "Template of email : \n\nWish you a very happy and blessed birthday !!\n" \
    "May God fulfill all desire of your heart and bestow you with health happiness and peace."
t2.insert(END, u)

t3 = Text(win, font=("commissars", 15, "bold"), height=12, width=36)
t3.place(x=230, y=325)
file = "C:\\Birthday\\Birthday.xlsx"


def check():
    if os.path.exists(file):
        try:
            t3.delete(1.0, END)
            today = date.today()
            x = int(today.day)
            y = int(today.month)
            if y <= int(9):
                y = "0" + str(y)
            if x <= int(9):
                x = "0" + str(x)

            z = str(y) + "-" + str(x)

            df = pd.read_excel(file)
            df["Date"] = pd.to_datetime(df["DOB"]).dt.strftime("%Y-%m-%d")
            df[["year", "month", "day"]] = df["Date"].str.split("-", expand=True)
            cols = ["month", "day"]
            df['M-D'] = df[cols].apply(lambda x: '-'.join(x.values.astype(str)), axis="columns")

            for i in range(2, len(df)):

                dt1 = df["M-D"][i]
                clg = df["College"][i]
                em = df["Email"][i]
                name = df["Name"][i]
                num = df["Contact"][i]
                dept = df["Dept"][i]
                if z == dt1:
                    x = "Name : {}\nCollege : {} ,{}\nEmail: {}\nNo : {}\n\n".format(name, clg, dept, em, num)
                    t3.insert(END, x)
                    continue
            if len(t1.get("1.0", END)) == 1:
                messagebox.showinfo("Info", "No Birthday Found")
        except:
            messagebox.showinfo("Error", "Something went wrong. Please check names of columns or try again.")
    else:
        messagebox.showerror("Error", "Birthday excel sheet unavailable")


def email():
    x1 = simpledialog.askstring("Info", "Enter sender's Email ID")
    if x1:
        x2 = simpledialog.askstring("Info", "Enter password")
        if x2:
            x3 = simpledialog.askstring("Info", "Enter receiver's Email ID")
            if x3:
                a = messagebox.askyesno("Info", "Do you want to send an email to {}?".format(x3))
                if a:
                    c = t2.get(3.0, END)
                    msg = EmailMessage()
                    msg.set_content(c)

                    msg['Subject'] = 'Birthday email '
                    msg['From'] = x1
                    msg['To'] = x3
                    try:
                        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
                        server.login(x1, x2)
                        server.send_message(msg)
                        server.quit()
                        messagebox.showinfo("Info", "Email sent successfully")
                    except:
                        messagebox.showerror("Info", "Email not sent. Please check the SMTP permission else try again")
            else:
                messagebox.showinfo("Info", "No email entered. Please enter receiver's email id to send an email")
        else:
            messagebox.showinfo("Info", "No password entered. Please enter password to send an email")
    else:
        messagebox.showinfo("Info", "No email entered. Please enter sender's email id to send an email")


b2 = Button(win, text="Check for Birthdays", font=("commissars", 16), fg="green", bg="White", command=check)
b2.place(x=650, y=400)
b3 = Button(win, text="Send Email", font=("commissars", 16), fg="green", bg="White", command=email)
b3.place(x=680, y=480)


win.mainloop()