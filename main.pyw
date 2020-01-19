import sqlite3
from tkinter import Tk,mainloop,Label,Frame,BOTH,Pack,Grid,RIDGE,Scrollbar,RIGHT,LEFT,Text,END,YES,Canvas,N,E,NO,Y,X,Entry,NW,W,S,Checkbutton
from tkinter import IntVar,DISABLED,FLAT,NORMAL,messagebox,GROOVE,Button,CENTER
from pathlib import Path
import datetime
import win32com.client as win32  

#import getpass
#print(getpass.getuser())

#Preprocessor-Sytle
fobj_in = open("./config/studentspath.txt")
DATEI = fobj_in.readline()
fobj_in = open("./config/databasepath.txt")
DB = fobj_in.readline()

BREITE = 760
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

        UNTERLAGEN = ["Anmeldung","Ausbildungsvertrag","Datenkontrollblatt",
                    "Merkblatt ärtzliche Leistung" ,"Merkblatt Infektionenschutz",
                    "Einwilligung Datenschutz", "Abschlusszeugnis"]
        conn = sqlite3.connect(DB)
        db = conn.cursor()
        sql = "SELECT studentsid FROM students WHERE class='{}'".format(self.klasse)
        db.execute(sql)
        rows = db.fetchall()
        for row in rows:
            sql = "SELECT anmeldung,av,edv,mb,infekt,daten,zeugnis FROM documents WHERE studentsid={}".format(row[0])
            db.execute(sql)
            ary = db.fetchone()
            if sum(ary) < 7:
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
                            html += "{}".format(UNTERLAGEN[fehlt])
                            helper += 1
                    else:
                        if ary[fehlt] == 0:
                            html += ", {}".format(UNTERLAGEN[fehlt])
                html += "</td>\n</tr>\n"
        
        conn.close()
        #print(len(html))
        if len(html) > 360:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            #mail.To = "" #recipient
            mail.Subject = "Erinnerung an fehlende Schülerunterlagen" #subject
            html += """
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
    def __init__(self, klasse, name, vorname, geb, studentid, anmeldung, av, edv, mb, infekt, daten, zeugnis):
        self.klasse = klasse
        self.name = name
        self.vorname = vorname
        self.geb = geb
        self.studentid = studentid
        self.anmeldung = anmeldung
        self.av = av
        self.edv = edv
        self.mb = mb
        self.infekt = infekt
        self.daten = daten
        self.zeugnis = zeugnis
    
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
        if which == "anmeldung":
            if self.anmeldung == 1:
                self.anmeldung = 0
            else:
                self.anmeldung = 1
            self.sql_update(which,self.anmeldung)
        
        if which == "av":
            if self.av == 1:
                self.av = 0
            else:
                self.av = 1
            self.sql_update(which,self.av)

        if which == "edv":
            if self.edv == 1:
                self.edv = 0
            else:
                self.edv = 1
            self.sql_update(which,self.edv)
        
        if which == "mb":
            if self.mb == 1:
                self.mb = 0
            else:
                self.mb = 1
            self.sql_update(which,self.mb)
        
        if which == "infekt":
            if self.infekt == 1:
                self.infekt = 0
            else:
                self.infekt = 1
            self.sql_update(which,self.infekt)
        
        if which == "daten":
            if self.daten == 1:
                self.daten = 0
            else:
                self.daten = 1
            self.sql_update(which,self.daten)
        
        if which == "zeugnis":
            if self.zeugnis == 1:
                self.zeugnis = 0
            else:
                self.zeugnis = 1
            self.sql_update(which,self.zeugnis)
        
    def render_line1(self,Tk,zahl,color):
        content = Frame(canvas)

        l = Label(content)
        l.configure(width=31, text=self.get_student(), anchor=W, background=color)
        l.pack(side=LEFT)

        LBREITE = 7

        self.cb1var = IntVar()
        self.cb1var.set(self.anmeldung)
        cb1 = Checkbutton(content, var=self.cb1var, command=lambda: self.change("anmeldung"))
        cb1.configure(width=LBREITE)
        cb1.configure(background=color, borderwidth=0)
        cb1.pack(side=LEFT)
        
        self.cb2var = IntVar()
        self.cb2var.set(self.av)
        cb2 = Checkbutton(content, var=self.cb2var, command=lambda: self.change("av"))
        cb2.configure(width=LBREITE, borderwidth=0)
        cb2.configure(background=color)
        cb2.pack(side=LEFT)
        
        self.cb3var = IntVar()
        self.cb3var.set(self.edv)
        cb3 = Checkbutton(content, var=self.cb3var, command=lambda: self.change("edv"))
        cb3.configure(width=LBREITE, borderwidth=0)
        cb3.configure(background=color)
        cb3.pack(side=LEFT)

        self.cb4var = IntVar()
        self.cb4var.set(self.mb)
        cb4 = Checkbutton(content, var=self.cb4var, command=lambda: self.change("mb"))
        cb4.configure(width=LBREITE, borderwidth=0)
        cb4.configure(background=color)
        cb4.pack(side=LEFT)

        self.cb5var = IntVar()
        self.cb5var.set(self.infekt)
        cb5 = Checkbutton(content, var=self.cb5var, command=lambda: self.change("infekt"))
        cb5.configure(width=LBREITE, borderwidth=0)
        cb5.configure(background=color)
        cb5.pack(side=LEFT)

        self.cb6var = IntVar()
        self.cb6var.set(self.daten)
        cb6 = Checkbutton(content, var=self.cb6var, command=lambda: self.change("daten"))
        cb6.configure(width=LBREITE, borderwidth=0)
        cb6.configure(background=color)
        cb6.pack(side=LEFT)

        self.cb7var = IntVar()
        self.cb7var.set(self.zeugnis)
        cb7 = Checkbutton(content, var=self.cb7var, command=lambda: self.change("zeugnis"))
        cb7.configure(width=LBREITE, borderwidth=0)
        cb7.configure(background=color)
        cb7.pack(side=LEFT)
        
        canvas.create_window(0,zahl*20,window=content, anchor=NW)

conn = sqlite3.connect(DB)
db = conn.cursor()

sql = "SELECT b.class, b.name, b.vorname, b.geb, a.* FROM documents a, students b WHERE a.studentsid = b.studentsid ORDER BY class,name ASC"
db.execute(sql)
rows = db.fetchall()
students = []
for line in rows:
    students.append(Student(line[0],line[1],line[2],line[3],line[4],line[5],line[6],line[7],line[8],line[9],line[10],line[11]))

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
root.iconbitmap(ICONPATH)
root.configure(background="white")

header = Frame(root)
header.pack(fill=BOTH, padx=5, pady=5)

hlabel = Label(header)
hlabel.pack(fill=BOTH)
hlabel.configure(text="Schülerunterlagen", background="white", font=("Arial",68))

tableheader = Frame(root)
tableheader.pack(fill=BOTH, padx=5, pady=0)
tableheader.configure(background="white")

THDESIGN = GROOVE
thlabel = Label(tableheader)
thlabel.pack(side=LEFT, padx=0, pady=0)
thlabel.configure(text="Name, Vorname (Geburtsdatum)", background="lightgray", font=("Arial",8), relief=THDESIGN, width=36, anchor=W)

thcolumns = [
    ["Anmeldung", 11],
    ["AV", 11],
    ["EDV-Kontr.", 11],
    ["MB är.Lei.", 11],
    ["Inf.schutz", 11],
    ["Dat.schutz", 11],
    ["Zeug.kopie", 11]
]

for element in thcolumns:
    thlabel = Label(tableheader)
    thlabel.pack(side=LEFT, padx=0, pady=0)
    thlabel.configure(text=" {} ".format(element[0]), background="lightgray", font=("Arial",8), relief=THDESIGN, width=element[1], anchor=N)

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
            students[x].render_line1(canvas,durchlauf,"white")
            durchlauf += 1
        try:    
            if classes[i] == students[x + 1].klasse:
                students[x+1].render_line1(canvas,durchlauf,"lavender")
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