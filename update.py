# Programm zur Verwaltung der Schülerunterlagen an bayerischen Berufsschulen
# (c) by Claus Gumbmann
# Das Programm ist unter der Mozilla Public License (MPL) veröffentlich.

# Dieses Programm fürht Änderungen an der Datenbank durch.
# Bisher nicht enthaltene Schüler werden in die Datenbank aufgenommen.
# Die Vorbelegung der Unterlagen ist 0, d.h. keine Unterlagen vorhanden. 
# Schüler, bei der sich die Klasse geändert hat, werden aktuallisiert.
# Schüler, die nicht mehr in der students.txt stehen, werden vollständig aus der Datenbank gelöscht.

#Verwendete Drittsoftware: Python 3.8, SQLite3, pywin32 und pyXML

print("Es geht los...")
import sqlite3
import os
import datetime
from pathlib import Path
import sys

#Preprocessor-Sytle
try:
    fobj_in = open("./config/studentspath.txt")
    DATEI = fobj_in.readline()
    if Path(DATEI).stat().st_size < 80000:
        print("Die Dateigröße der Schülerdaten scheint etwas klein zu sein!")
        if input("Schülerdatei überprüft? (j)a/(n)ein? ") != "j":
            sys.exit(1)
    fobj_in.close()
except:
    print("Sollte die Datei in Ordnung sein, dann Update wiederholen und j drücken.")
    print("Um die Plausibilitätsprüfung der Schülerdaten anzupassen, muss eine Änderung im Programmcode vorgenommen werden.")
    input("Enter drücken, um das Fenster zu schließen...")
    sys.exit()
    

try:
    fobj_in = open("./config/databasepath.txt")
    DB = fobj_in.readline()
    fobj_in.close()
except:
    print(DB)
    input()

try:
    fobj_in = open("./config/withoutclass.txt")
    AUSLASSUNGEN = fobj_in.readline().strip().split(",")
    fobj_in.close()
except:
    print(AUSLASSUNGEN)
    input()
#import getpass
#print(getpass.getuser())

#Verbidnung zur Textdatei hertsellen Klasse/tName/tVorname/tGeburtsdatum/t
#restlichen Werte werden nicht benötigt
fobj_in = open(DATEI,encoding="ANSI")

#Datenbankverbindung öffnen
conn = sqlite3.connect(DB)

#Zähler auf null, für die Ausgabe am Ende, wie viele Datensätze geändert wurden
#!= gelöschte Elemente, die werden weiter unten ermittelt 
counter = 0

#Liste für die studentensids in der Textdatei
stud_ids = []

#Liste für die gelöschten Schüler
deleted_students = []

#Beginn des Abgleichprozesses
for Zeile, line_txt in enumerate(fobj_in.readlines()):
    #Statusanzeige der offenen Command Prompt Konsole
    if counter == 250:
        os.system('cls')
        print("250 Schüler eingelesen")
    elif counter == 500:
        os.system('cls')
        print("500 Schüler eingelesen")
    elif counter == 750:
        os.system('cls')
        print("750 Schüler eingelesen")
    elif counter == 1000:
        os.system('cls')
        print("1000 Schüler eingelesen")
    elif counter > 1500:
        os.system('cls')
        print("jetzt kann es aber nicht mehr lange dauern...")

    ary = line_txt.strip().split('\t') #Seperator in der Textdatei ist die Tab-Taste
    #ary[2] = ary[2].strip.split(' ')
    
    if ary[0] not in AUSLASSUNGEN:
        sql = "SELECT class, studentsid FROM students WHERE name='{}' AND vorname='{}' AND geb='{}'".format(ary[1],ary[2],ary[3])
        db = conn.cursor()
        db.execute(sql)
        row = db.fetchone()
        if row != None:
            stud_ids.append(row[1])

        if (row == None):
            sql = "INSERT INTO students(class,name,vorname,geb) VALUES ('{}','{}','{}','{}')".format(ary[0],ary[1],ary[2],ary[3])
            #db = conn.cursor()
            db.execute(sql)
            conn.commit()
            sql = "SELECT studentsid FROM students WHERE name='{}' AND vorname='{}' AND geb='{}'".format(ary[1],ary[2],ary[3])
            #db = conn.cursor()
            db.execute(sql)
            row = db.fetchone()
            stud_ids.append(row[0])
            sql = "INSERT INTO documents(studentsid) VALUES ('{}')".format(row[0])
            #db = conn.cursor()
            db.execute(sql)
            conn.commit()
            counter += 1
        else:
            if (row[0] != ary[0]):
                sql = "UPDATE students SET class='{}' WHERE studentsid={}".format(ary[0],row[1])
                #db = conn.cursor()
                db.execute(sql)
                conn.commit()
                counter += 1

#Verbindung zur Textdatei schließen
fobj_in.close()

sql = "SELECT studentsid FROM students ORDER BY studentsid ASC"
db = conn.cursor()
db.execute(sql)
row = db.fetchall()

db_stud_ids = []
for x in row:
    db_stud_ids.append(x[0])

dif = set(sorted(db_stud_ids)) - set(sorted(stud_ids))
for x in dif:
    sql = "SELECT class,name,vorname,geb FROM students WHERE studentsid={}".format(x)
    db.execute(sql)
    row = db.fetchone()
    deleted_students.append(row)
    sql = "DELETE FROM students WHERE studentsid={}".format(x)
    #db = conn.cursor()
    db.execute(sql)
    sql = "DELETE FROM documents WHERE studentsid={}".format(x)
    #db = conn.cursor()
    db.execute(sql)
    conn.commit()

#LastSaved-Datum ermitteln und in der Datenbank speichern
path = Path(DATEI)
sql = "UPDATE 'data' SET 'lastSaved'='{}'".format(datetime.datetime.fromtimestamp(path.stat().st_mtime).strftime("%d.%m.%Y (%H:%M)"))
#db = conn.cursor()
db.execute(sql)
conn.commit()

#Updatedatum in der Datenbank speichern
sql = "UPDATE 'data' SET 'update'='{}'".format(datetime.datetime.now().strftime("%d.%m.%Y (%H:%M)"))
#db = conn.cursor()
db.execute(sql)
conn.commit()

#Datenbankverbindung schließen
conn.close()

#Ausgabe der Veränderungen
if (counter==1):
    print("fertig! 1 Datensatz geschrieben und {} Schüler gelöscht".format(len(dif)))
    if len(deleted_students) > 0:
        print("")
        print("Folgender Schüler wurde gelöscht:")
        print(deleted_students)
else:
    print("fertig! {} Datensätze geschrieben und {} Schüler gelöscht".format(counter,len(dif)))
    if len(deleted_students) > 0:
        print("")
        print("Folgende Schüler wurden gelöscht:")
        for x in deleted_students:
            print(x)

#Damit am Ende das Fenster mit der Ausgabe offen bleibt
input("Enter drücken, um das Fenster zu schließen...")