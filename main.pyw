# Programm zur Verwaltung der Schülerunterlagen an bayerischen Berufsschulen
# (c) by Claus Gumbmann
# Das Programm ist unter der Mozilla Public License (MPL) veröffentlich

#Verwendete Drittsoftware: Python 3.8, SQLite3, pywin32 und pyXML

import sqlite3
from tkinter import Tk,mainloop,Label,Frame,BOTH,Pack,Grid,RIDGE,Scrollbar,RIGHT,LEFT,Text,END,YES,Canvas,N,E,NO,Y,X,Entry,NW,W,S,Checkbutton
from tkinter import IntVar,DISABLED,FLAT,NORMAL,messagebox,GROOVE,Button,CENTER
from pathlib import Path
import datetime
import platform
if platform.system() == "Windows":
    import win32com.client as win32
else:
    win32 = None
import xml.etree.ElementTree as ET

tree = ET.parse('./config/headers.xml')
root = tree.getroot()

#import getpass
#print(getpass.getuser())

#Preprocessor-Sytle
fobj_in = open("./config/studentspath.txt")
DATEI = fobj_in.readline()
fobj_in = open("./config/databasepath.txt")
DB = fobj_in.readline()
tree = ET.parse('./config/headers.xml')
root = tree.getroot()

THCOLUMS = []
for child in root:
    THCOLUMS.append(child.text)

if len(THCOLUMS) < 15:
    BREITE = 257 + (len(THCOLUMS) * 65)
else:
    BREITE = 264 + (len(THCOLUMS) * 65)

HÖHE = 800
ICONPATH = "./img/favicon.ico"

class MyButtons():
    def __init__(self, klasse, anzahl):
        self.klasse = klasse
        self.anzahl = anzahl
    
    def mail_sender(self, event):
        #print("Klasse: {}, Schüler: {}".format(self.klasse,self.anzahl))
        html = """Sehr geehrter Klassenleiter der Klasse {},<p>
von folgenden Schülern fehlen noch Unterlagen.<p>
<table style='border:1px solid gray;'>
<tr>
<th style='border:1px solid gray;'>Name</th>
<th style='border:1px solid gray;'>Vorname</th>
<th style='border:1px solid gray;'>Geburtsdatum</th>
<th style='border:1px solid gray;'>fehlende Unterlagen</th>
</tr>
""".format(self.klasse)

        #UNTERLAGEN = ["Anmeldung","Ausbildungsvertrag","Datenkontrollblatt",
                    #"Merkblatt ärtzliche Leistung" ,"Merkblatt Infektionenschutz",
                    #"Einwilligung Datenschutz", "Abschlusszeugnis"]

        conn = sqlite3.connect(DB)
        db = conn.cursor()
        sql = "SELECT studentsid FROM students WHERE class='{}'".format(self.klasse)
        db.execute(sql)
        rows = db.fetchall()
        for row in rows:
            sql = "SELECT "
            for unbr in range(1,len(THCOLUMS) + 1):
                sql += "u{}".format(unbr)
                if unbr < len(THCOLUMS):
                    sql += ","
            sql += " FROM documents WHERE studentsid={}".format(row[0])
            db.execute(sql)
            ary = db.fetchone()
            if sum(ary) < len(THCOLUMS):
                helper = 0
                sql = "SELECT name,vorname,geb FROM students WHERE studentsid='{}'".format(row[0])
                db.execute(sql)
                element = db.fetchone()
                html += "<tr>\n"
                for x in element:
                    html += "<td style='border:1px solid gray;'>{}</td>\n".format(x)
                html += "<td style='border:1px solid gray;'>"
                for fehlt in range(0,len(ary)):
                    if helper == 0:
                        if ary[fehlt] == 0:
                            html += "{}".format(THCOLUMS[fehlt])
                            helper += 1
                    else:
                        if ary[fehlt] == 0:
                            html += ", {}".format(THCOLUMS[fehlt])
                html += "</td>\n</tr>\n"
        
        conn.close()
        #print(len(html))
        if len(html) > 360:
            if win32:
                outlook = win32.Dispatch("outlook.application")
                mail = outlook.CreateItem(0)
                #mail.To = "" #recipient
                mail.Subject = "Erinnerung an fehlende Schülerunterlagen" #subject
                html += """
            else:
                 messagebox.showinfo("Info", "E-Mail-Funktion steht unter Linux nicht zur Verfügung.")
            </table>
            <p>Vielen Dank für die Unterstützung.
            <P>Mit freundlichen Grüßen
            <p><p>
            """
            mail.HtmlBody = html #text
            #mail.send
            mail.Display(True)
        else:
             event.widget['text'] = "Keine Mail erfolderlich!"

    def over_cursor(self, event):
        event.widget['cursor'] = "hand2"
        event.widget['text'] = "Klassenleiter per E-Mail an fehlende Unterlagen erinnern!"

    def leave_cursor(self, event):
        event.widget['cursor'] = 'arrow'
        event.widget['text'] = "{}  ( {} Schüler )".format(self.klasse,self.anzahl)

    def render_line(self,Tk):
        bt = Button(content)
        bt.configure(text="{}  ( {} Schüler )".format(self.klasse,self.anzahl), background="deep sky blue",relief=FLAT)
        bt.pack()
        bt.bind('<Button-1>', self.mail_sender)
        bt.bind('<Enter>', self.over_cursor)
        bt.bind('<Leave>', self.leave_cursor)

class Student():

    def __init__(self, klasse, name, vorname, geb, studentid, u1, u2, u3, u4, u5, u6, u7, u8, u9, u10, u11, u12, u13, u14, u15):
        self.klasse = klasse
        self.name = name
        self.vorname = vorname
        self.geb = geb
        self.studentid = studentid
        self.u1 = u1
        self.u2 = u2
        self.u3 = u3
        self.u4 = u4
        self.u5 = u5
        self.u6 = u6
        self.u7 = u7
        self.u8 = u8
        self.u9 = u9
        self.u10 = u10
        self.u11 = u11
        self.u12 = u12
        self.u13 = u13
        self.u14 = u14
        self.u15 = u15
        
    def get_student(self):
        return "{}, {} ({})".format(self.name,self.vorname,self.geb)

    def sql_update(self, which, what):
        conn = sqlite3.connect(DB)
        db = conn.cursor()
        sql = "UPDATE documents SET {} = {} WHERE studentsid = {}".format(which, what, self.studentid)
        db.execute(sql)
        conn.commit()
        conn.close()
    
    def change(self,which):
        if which == "u1":
            if self.u1 == 1:
                self.u1 = 0
            else:
                self.u1 = 1
            self.sql_update(which,self.u1)

        if which == "u2":
            if self.u2 == 1:
                self.u2 = 0
            else:
                self.u2 = 1
            self.sql_update(which,self.u2)

        if which == "u3":
            if self.u3 == 1:
                self.u3 = 0
            else:
                self.u3 = 1
            self.sql_update(which,self.u3)
        
        if which == "u4":
            if self.u4 == 1:
                self.u4 = 0
            else:
                self.u4 = 1
            self.sql_update(which,self.u4)

        if which == "u5":
            if self.u5 == 1:
                self.u5 = 0
            else:
                self.u5 = 1
            self.sql_update(which,self.u5)
        
        if which == "u6":
            if self.u6 == 1:
                self.u6 = 0
            else:
                self.u6 = 1
            self.sql_update(which,self.u6)

        if which == "u7":
            if self.u7 == 1:
                self.u7 = 0
            else:
                self.u7 = 1
            self.sql_update(which,self.u7)

        if which == "u8":
            if self.u8 == 1:
                self.u8 = 0
            else:
                self.u8 = 1
            self.sql_update(which,self.u8)

        if which == "u9":
            if self.u9 == 1:
                self.u9 = 0
            else:
                self.u9 = 1
            self.sql_update(which,self.u9)

        if which == "u10":
            if self.u10 == 1:
                self.u10 = 0
            else:
                self.u10 = 1
            self.sql_update(which,self.u10)

        if which == "u11":
            if self.u11 == 1:
                self.u11 = 0
            else:
                self.u11 = 1
            self.sql_update(which,self.u11)
        
        if which == "u12":
            if self.u12 == 1:
                self.u12 = 0
            else:
                self.u12 = 1
            self.sql_update(which,self.u12)
        
        if which == "u13":
            if self.u13 == 1:
                self.u13 = 0
            else:
                self.u13 = 1
            self.sql_update(which,self.u13)
        
        if which == "u14":
            if self.u14 == 1:
                self.u14 = 0
            else:
                self.u14 = 1
            self.sql_update(which,self.u14)
        
        if which == "u15":
            if self.u15 == 1:
                self.u15 = 0
            else:
                self.u15 = 1
            self.sql_update(which,self.u15)
        
    def render_line(self,Tk,zahl,color):
        content = Frame(canvas)

        l = Label(content)
        l.configure(width=31, text=self.get_student(), anchor=W, background=color)
        l.pack(side=LEFT)

        LBREITE = 6

        if 1 <= len(THCOLUMS):
            self.cb1var = IntVar()
            self.cb1var.set(self.u1)
            cb1 = Checkbutton(content, var=self.cb1var, command=lambda: self.change("u1"))
            cb1.configure(width=LBREITE)
            cb1.configure(background=color, borderwidth=0)
            cb1.pack(side=LEFT)

        if 2 <= len(THCOLUMS): 
            self.cb2var = IntVar()
            self.cb2var.set(self.u2)
            cb2 = Checkbutton(content, var=self.cb2var, command=lambda: self.change("u2"))
            cb2.configure(width=LBREITE, borderwidth=0)
            cb2.configure(background=color)
            cb2.pack(side=LEFT)
        
        if 3 <= len(THCOLUMS):
            self.cb3var = IntVar()
            self.cb3var.set(self.u3)
            cb3 = Checkbutton(content, var=self.cb3var, command=lambda: self.change("u3"))
            cb3.configure(width=LBREITE, borderwidth=0)
            cb3.configure(background=color)
            cb3.pack(side=LEFT)

        if 4 <= len(THCOLUMS):
            self.cb4var = IntVar()
            self.cb4var.set(self.u4)
            cb4 = Checkbutton(content, var=self.cb4var, command=lambda: self.change("u4"))
            cb4.configure(width=LBREITE, borderwidth=0)
            cb4.configure(background=color)
            cb4.pack(side=LEFT)

        if 5 <= len(THCOLUMS):
            self.cb5var = IntVar()
            self.cb5var.set(self.u5)
            cb5 = Checkbutton(content, var=self.cb5var, command=lambda: self.change("u5"))
            cb5.configure(width=LBREITE, borderwidth=0)
            cb5.configure(background=color)
            cb5.pack(side=LEFT)

        if 6 <= len(THCOLUMS):
            self.cb6var = IntVar()
            self.cb6var.set(self.u6)
            cb6 = Checkbutton(content, var=self.cb6var, command=lambda: self.change("u6"))
            cb6.configure(width=LBREITE, borderwidth=0)
            cb6.configure(background=color)
            cb6.pack(side=LEFT)

        if 7 <= len(THCOLUMS):
            self.cb7var = IntVar()
            self.cb7var.set(self.u7)
            cb7 = Checkbutton(content, var=self.cb7var, command=lambda: self.change("u7"))
            cb7.configure(width=LBREITE, borderwidth=0)
            cb7.configure(background=color)
            cb7.pack(side=LEFT)

        if 8 <= len(THCOLUMS):
            self.cb8var = IntVar()
            self.cb8var.set(self.u8)
            cb8 = Checkbutton(content, var=self.cb8var, command=lambda: self.change("u8"))
            cb8.configure(width=LBREITE, borderwidth=0)
            cb8.configure(background=color)
            cb8.pack(side=LEFT)

        if 9 <= len(THCOLUMS):
            self.cb9var = IntVar()
            self.cb9var.set(self.u9)
            cb9 = Checkbutton(content, var=self.cb9var, command=lambda: self.change("u9"))
            cb9.configure(width=LBREITE, borderwidth=0)
            cb9.configure(background=color)
            cb9.pack(side=LEFT)

        if 10 <= len(THCOLUMS):
            self.cb10var = IntVar()
            self.cb10var.set(self.u10)
            cb10 = Checkbutton(content, var=self.cb10var, command=lambda: self.change("u10"))
            cb10.configure(width=LBREITE, borderwidth=0)
            cb10.configure(background=color)
            cb10.pack(side=LEFT)

        if 11 <= len(THCOLUMS):
            self.cb11var = IntVar()
            self.cb11var.set(self.u11)
            cb11 = Checkbutton(content, var=self.cb11var, command=lambda: self.change("u11"))
            cb11.configure(width=LBREITE, borderwidth=0)
            cb11.configure(background=color)
            cb11.pack(side=LEFT)

        if 12 <= len(THCOLUMS):
            self.cb12var = IntVar()
            self.cb12var.set(self.u12)
            cb12 = Checkbutton(content, var=self.cb12var, command=lambda: self.change("u12"))
            cb12.configure(width=LBREITE, borderwidth=0)
            cb12.configure(background=color)
            cb12.pack(side=LEFT)

        if 13 <= len(THCOLUMS):
            self.cb13var = IntVar()
            self.cb13var.set(self.u13)
            cb13 = Checkbutton(content, var=self.cb13var, command=lambda: self.change("u13"))
            cb13.configure(width=LBREITE, borderwidth=0)
            cb13.configure(background=color)
            cb13.pack(side=LEFT)
        
        if 14 <= len(THCOLUMS):
            self.cb14var = IntVar()
            self.cb14var.set(self.u14)
            cb14 = Checkbutton(content, var=self.cb14var, command=lambda: self.change("u14"))
            cb14.configure(width=LBREITE, borderwidth=0)
            cb14.configure(background=color)
            cb14.pack(side=LEFT)

        if 15 == len(THCOLUMS):
            self.cb15var = IntVar()
            self.cb15var.set(self.u15)
            cb15 = Checkbutton(content, var=self.cb15var, command=lambda: self.change("u15"))
            cb15.configure(width=LBREITE, borderwidth=0)
            cb15.configure(background=color)
            cb15.pack(side=LEFT)
        
        canvas.create_window(0,zahl*20,window=content, anchor=NW)

conn = sqlite3.connect(DB)
db = conn.cursor()

sql = "SELECT b.class, b.name, b.vorname, b.geb, a.* FROM documents a, students b WHERE a.studentsid = b.studentsid ORDER BY class,name ASC"
db.execute(sql)
rows = db.fetchall()
students = []
for line in rows:
    students.append(Student(line[0],line[1],line[2],line[3],line[4],
                            line[5],line[6],line[7],line[8],line[9],
                            line[10],line[11],line[12],line[13],line[14],
                            line[15],line[16],line[17],line[18],line[19]))

sql = "SELECT class FROM students Group by class Order by class ASC"
db.execute(sql)
rows = db.fetchall()
classes = []
for line in rows:
    classes.append(line[0])

root = Tk()
root.geometry("{}x{}+0+0".format(BREITE,HÖHE))
root.resizable(width=False,height=False)
root.title("Copyright by Claus Gumbmann")
import platform
if platform.system() == "Windows":
    root.iconbitmap(ICONPATH)
root.configure(background="white")

header = Frame(root)
header.pack(fill=BOTH, padx=5, pady=5)

hlabel = Label(header)
hlabel.pack(fill=BOTH)

if len(THCOLUMS)*8 > 72:
    FONT = 72
else:
    FONT =  len(THCOLUMS)*8

hlabel.configure(text="Schülerunterlagen", background="white", font=("Arial",FONT))

tableheader = Frame(root)
tableheader.pack(fill=BOTH, padx=5, pady=0)
tableheader.configure(background="white")

THDESIGN = GROOVE
thlabel = Label(tableheader)
thlabel.pack(side=LEFT, padx=0, pady=0)
thlabel.configure(text="Name, Vorname (Geburtsdatum)", background="lightgray", font=("Arial",8), relief=THDESIGN, width=36, anchor=W)

for element in THCOLUMS:
    thlabel = Label(tableheader)
    thlabel.pack(side=LEFT, padx=0, pady=0)
    thlabel.configure(text=" {} ".format(element), background="lightgray", font=("Arial",8), relief=THDESIGN, width=10, anchor=N)

canvas = Canvas(root)
canvas.pack(fill=BOTH, expand=YES, padx=5, pady=5)

sb =Scrollbar(canvas)
sb.pack(side=RIGHT,fill=Y)

scrollable = (len(students)+len(classes)) * 20
if scrollable < 640:
    scrollable = 0

canvas.configure(background="white", yscrollcommand=sb.set ,scrollregion=(0,0,0,scrollable), highlightthickness=0)
sb.configure(command=canvas.yview)

def OnMouseWheel(event):
    if scrollable != 0:
        canvas.yview("scroll",int(event.delta/120) * -1, "units")

def fast_OnMouseWheel(event):
    if scrollable != 0:
        canvas.yview("scroll",int(event.delta/30) * -1, "units")

def superfast_OnMouseWheel(event):
    if scrollable != 0:
        canvas.yview("scroll",event.delta * -1, "units")

root.bind("<MouseWheel>", OnMouseWheel)
root.bind("<Control-MouseWheel>", fast_OnMouseWheel)
root.bind("<Control-Shift-MouseWheel>", superfast_OnMouseWheel)

durchlauf = 0
for i in range(0,len(classes)):
    content = Frame(canvas)
    content.configure(background="deep sky blue")
    canvas.create_window(BREITE/2-26,durchlauf*20,window=content, anchor=N, width=BREITE)

    sql = "SELECT COUNT(class) FROM students WHERE class='{}'".format(classes[i])
    db.execute(sql)
    counter = db.fetchone()
    #print(counter)

    mybutton = MyButtons(classes[i],counter[0])
    mybutton.render_line(content)

    durchlauf += 1
    for x in range(0, len(students), 2):
        if classes[i] == students[x].klasse:
            students[x].render_line(canvas,durchlauf,"white")
            durchlauf += 1
        try:    
            if classes[i] == students[x + 1].klasse:
                students[x+1].render_line(canvas,durchlauf,"lavender")
                durchlauf += 1
        except:
            pass

conn.close()         
def on_closing():
    conn = sqlite3.connect(DB)
    db = conn.cursor()
    sql = "SELECT * FROM 'data'"
    db.execute(sql)
    row = db.fetchone()
    conn.close()

    MESSAGE = "Möchten Sie das Programm beenden?\n"
    
    path = Path(DATEI)
    LASTSAVED = datetime.datetime.fromtimestamp(path.stat().st_mtime).strftime("%d.%m.%Y (%H:%M)")

    if LASTSAVED != row[1]:
        MESSAGE += "\nUpdate erforderlich!\n"
        MESSAGE += "Datenquelle (alt) vom:\t{}\n".format(row[1])
        MESSAGE += "Datenquelle (neu) vom:\t{}".format(LASTSAVED)

    if messagebox.askokcancel("DATENSTAND: {}".format(row[0]),MESSAGE):
        root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()