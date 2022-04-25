from tkinter import *
from PIL import ImageTk
from tkinter import messagebox, filedialog
import os
import pandas as pd
import email_function
import time


class BulkEmail:
    def __init__(self, root):
        self.root = root
        self.root.title("Bulk Email Application")
        self.root.geometry("1000x550+200+50")
        self.root.config(bg="white")

        # Icons
        self.email_icon = ImageTk.PhotoImage(file="images/email.png")
        self.settings_icon = ImageTk.PhotoImage(file="images/settings.png")

        # Title
        title = Label(
            self.root,
            text="Send Bulk Emails!",
            image=self.email_icon,
            padx=10,
            compound=LEFT,
            font=("Goudy Old Style", 48, "bold"),
            bg="#222A35",
            fg="white",
            anchor="w",
        ).place(x=0, y=0, relwidth=1)
        description = Label(
            self.root,
            text="Upload an Excel File to send bulk emails!",
            padx=10,
            compound=LEFT,
            font=("Calibri (Body)", 14),
            bg="#fac000",
            fg="#262626",
        ).place(x=0, y=80, relwidth=1)
        btn_setting = Button(
            self.root,
            image=self.settings_icon,
            activebackground="#222A35",
            cursor="hand2",
            command=self.setting_window,
        ).place(x=900, y=5)

        # Radio Buttons
        self.var_choice = StringVar()
        single = Radiobutton(
            self.root,
            text="Single",
            value="single",
            variable=self.var_choice,
            activebackground="white",
            font=("consolas", 30, "bold"),
            bg="white",
            fg="#262626",
            command=self.check_single_or_bulk,
        ).place(x=50, y=150)
        bulk = Radiobutton(
            self.root,
            text="Bulk",
            value="bulk",
            variable=self.var_choice,
            activebackground="white",
            font=("consolas", 30, "bold"),
            bg="white",
            fg="#262626",
            command=self.check_single_or_bulk,
        ).place(x=250, y=150)
        self.var_choice.set("single")

        # The Actual Content
        to = Label(self.root, text="To ", font=("consolas", 18), bg="white").place(
            x=50, y=250
        )
        subject = Label(
            self.root, text="Subject ", font=("consolas", 18), bg="white"
        ).place(x=50, y=300)
        message = Label(
            self.root, text="Message ", font=("consolas", 18), bg="white"
        ).place(x=50, y=350)

        self.txt_to = Entry(self.root, font=("consolas", 14), bg="#fac000")
        self.txt_to.place(x=200, y=250, width=350, height=30)

        self.btn_browser = Button(
            self.root,
            command=self.browse_file,
            text="Browse",
            font=("consolas", 18, "bold"),
            bg="#fac000",
            fg="black",
            activebackground="#fac000",
            activeforeground="black",
            cursor="hand2",
            state=DISABLED,
        )
        self.btn_browser.place(x=670, y=165, width=120, height=30)

        self.txt_subject = Entry(self.root, font=("consolas", 14), bg="#fac000")
        self.txt_subject.place(x=200, y=300, width=450, height=30)

        self.txt_message = Text(self.root, font=("consolas", 14), bg="#fac000")
        self.txt_message.place(x=200, y=350, width=750, height=120)

        self.lbl_total = Label(self.root, font=("consolas", 18), bg="white")
        self.lbl_total.place(x=50, y=490)

        self.lbl_sent = Label(self.root, font=("consolas", 18), bg="white", fg="green")
        self.lbl_sent.place(x=250, y=490)

        self.lbl_remaining = Label(
            self.root, font=("consolas", 18), bg="white", fg="orange"
        )
        self.lbl_remaining.place(x=50, y=550)

        self.lbl_failed = Label(self.root, font=("consolas", 18), bg="white", fg="red")
        self.lbl_failed.place(x=250, y=550)

        btn_send = Button(
            self.root,
            command=self.send_email,
            text="Send",
            font=("consolas", 18, "bold"),
            bg="#000000",
            fg="white",
            activebackground="#000000",
            activeforeground="white",
            cursor="hand2",
        ).place(x=700, y=490, width=120, height=30)
        btn_clear = Button(
            self.root,
            command=self.clear1,
            text="Clear",
            font=("consolas", 18, "bold"),
            bg="#FF0000",
            fg="white",
            activebackground="#FF0000",
            activeforeground="white",
            cursor="hand2",
        ).place(x=830, y=490, width=120, height=30)
        self.check_if_file_exists()

    def browse_file(self):
        op = filedialog.askopenfile(
            initialdir="/",
            title="Select csv file for emails.",
            filetypes=(("All Files", "*.*"), ("Excel Files", ".xlsx")),
        )
        if op != None:
            data = pd.read_excel(op.name, engine="openpyxl")
            if "Email" in data.columns:
                self.emails = list(data["Email"])
                email_collection = []
                for email in self.emails:
                    if pd.isnull(email) == False:
                        email_collection.append(email)
                self.emails = email_collection
                if len(self.emails) > 0:
                    self.txt_to.config(state=NORMAL)
                    self.txt_to.delete(0, END)
                    self.txt_to.insert(0, str(op.name.split("/")[-1]))
                    self.txt_to.config(state="readonly")
                    self.lbl_total.config(text="Total: " + str(len(self.emails)))
                    self.lbl_sent.config(text="Sent: ")
                    self.lbl_remaining.config(text="Remaining: ")
                    self.lbl_failed.config(text="Failed: ")
                else:
                    messagebox.showerror(
                        "Error", "This file doesn't have any emails!", parent=self.root
                    )
            else:
                messagebox.showerror(
                    "Error",
                    "Please select a csv file which has the 'Email' column in it!",
                    parent=self.root,
                )

    def send_email(self):
        x = len(self.txt_message.get("1.0", END))
        if self.txt_to.get() == "" or self.txt_subject.get() == "" or x == 1:
            messagebox.showerror("Error", "All fields are required!", parent=self.root)
        else:
            if self.var_choice.get() == "single":
                status = email_function.email_sent_function(
                    self.txt_to.get(),
                    self.txt_subject.get(),
                    self.txt_message.get("1.0", END),
                    self.email,
                    self.passcode,
                )
                if status == "s":
                    messagebox.showinfo(
                        "Success!", "Email has been sent, yay!", parent=self.root
                    )
                elif status == "f":
                    messagebox.showerror(
                        "Failed", "Email not sent, gah!", parent=self.root
                    )
            elif self.var_choice.get() == "bulk":
                self.failed = []
                self.s_count = 0
                self.f_count = 0
                for x in self.emails:
                    status = email_function.email_sent_function(
                        x,
                        self.txt_subject.get(),
                        self.txt_message.get("1.0", END),
                        self.email,
                        self.passcode,
                    )
                    if status == "s":
                        self.s_count += 1
                    elif status == "f":
                        self.f_count += 1
                    self.status_bar()
                messagebox.showinfo(
                    "Success!",
                    "Email has been sent, please check the status!",
                    parent=self.root,
                )

    def status_bar(self):
        self.lbl_total.config(text="Status: " + str(len(self.emails)) + "=>>")
        self.lbl_sent.config(text="Sent: " + str(self.s_count))
        self.lbl_remaining.config(
            text="Remaining: " + str(len(self.emails) - (self.s_count + self.f_count))
        )
        self.lbl_failed.config(text="Failed: " + str(self.f_count))
        self.lbl_total.update()
        self.lbl_sent.update()
        self.lbl_remaining.update()
        self.lbl_failed.update()

    def check_single_or_bulk(self):
        if self.var_choice.get() == "single":
            self.btn_browser.config(state=DISABLED)
            self.txt_to.config(state=NORMAL)
            self.txt_to.delete(0, END)
            self.clear1()
        elif self.var_choice.get() == "bulk":
            self.btn_browser.config(state=NORMAL)
            self.txt_to.delete(0, END)
            self.txt_to.config(state="readonly")

    def clear1(self):
        self.txt_to.config(state=NORMAL)
        self.txt_to.delete(0, END)
        self.txt_subject.delete(0, END)
        self.txt_message.delete("1.0", END)
        self.var_choice.set("single")
        self.btn_browser.config(state=DISABLED)
        self.lbl_total.config(text="")
        self.lbl_sent.config(text="")
        self.lbl_remaining.config(text="")
        self.lbl_failed.config(text="")

    def setting_window(self):
        self.check_if_file_exists()
        self.root2 = Toplevel()
        self.root2.title("Settings")
        self.root2.geometry("650x300+350+90")
        self.root2.focus_force()
        self.root2.grab_set()
        self.root2.config(bg="white")
        title2 = Label(
            self.root2,
            text="Credentials Settings",
            image=self.settings_icon,
            padx=10,
            compound=LEFT,
            font=("Goudy Old Style", 48, "bold"),
            bg="#222A35",
            fg="white",
            anchor="w",
        ).place(x=0, y=0, relwidth=1)
        description2 = Label(
            self.root2,
            text="Enter your email address and password to login.",
            padx=10,
            compound=LEFT,
            font=("Calibri (Body)", 14),
            bg="#fac000",
            fg="#262626",
        ).place(x=0, y=80, relwidth=1)

        email = Label(
            self.root2, text="Email: ", font=("consolas", 18), bg="white"
        ).place(x=50, y=125)
        passcode = Label(
            self.root2, text="Password: ", font=("consolas", 18), bg="white"
        ).place(x=50, y=175)

        self.txt_email = Entry(self.root2, font=("consolas", 14), bg="#fac000")
        self.txt_email.place(x=175, y=125, width=350, height=30)

        self.txt_passcode = Entry(
            self.root2, font=("consolas", 14), bg="#fac000", show="*"
        )
        self.txt_passcode.place(x=175, y=175, width=350, height=30)

        btn_save2 = Button(
            self.root2,
            command=self.save_setting,
            text="Save",
            font=("consolas", 18, "bold"),
            bg="#000000",
            fg="white",
            activebackground="#000000",
            activeforeground="white",
            cursor="hand2",
        ).place(x=275, y=225, width=120, height=30)
        btn_clear2 = Button(
            self.root2,
            command=self.clear2,
            text="Clear",
            font=("consolas", 18, "bold"),
            bg="#FF0000",
            fg="white",
            activebackground="#FF0000",
            activeforeground="white",
            cursor="hand2",
        ).place(x=405, y=225, width=120, height=30)
        self.txt_email.insert(0, self.email)
        self.txt_passcode.insert(0, self.passcode)

    def clear2(self):
        self.txt_email.delete(0, END)
        self.txt_passcode.delete(0, END)

    def check_if_file_exists(self):
        if os.path.exists("important.txt") == False:
            f = open("important.txt", "w")
            f.write(",")
            f.close()
        f2 = open("important.txt", "r")
        self.credentials = []
        for i in f2:
            self.credentials.append([i.split(",")[0], i.split(",")[1]])
        self.email = self.credentials[0][0]
        self.passcode = self.credentials[0][1]

    def save_setting(self):
        if self.txt_email.get() == "" or self.txt_passcode.get() == "":
            messagebox.showerror("Error", "All fields are required!", parent=self.root2)
        else:
            f = open("important.txt", "w")
            f.write(self.txt_email.get() + "," + self.txt_passcode.get())
            f.close()
            messagebox.showinfo(
                "Success!", "Username, password saved successfully!", parent=self.root2
            )
            self.check_if_file_exists()


root = Tk()
obj = BulkEmail(root)
root.mainloop()
