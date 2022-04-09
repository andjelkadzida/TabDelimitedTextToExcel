import os
import sys
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import xlsxwriter

root = tk.Tk()
root.title("Konvertor xyz, qtt i lst ekstenzija u Excel")
root.iconbitmap('C:/Users/andje/Desktop/TabDelimitedTextToExcel/launcher.ico')
root.withdraw()


def fileConversion():
    qttToChange = filedialog.askopenfilename()
    if qttToChange.endswith('.qtt') or (qttToChange.endswith('.xyz') or qttToChange.endswith('.lst')):
        while not os.path.isfile(qttToChange):
            print("Greska: " + qttToChange + " nije fajlidna putanja do fajla. Pokusajte ponovo...")
            qttToChange = filedialog.askopenfilename()
        df = pd.read_csv(qttToChange, sep='\t' or '\s*' or '\s+' or 't', engine='python', encoding='utf-8')
        df.rename(columns={'Unnamed: 0': ''}, inplace=True)
        df.rename(columns={'Unnamed: 33': ''}, inplace=True)
        fileName = os.path.splitext(qttToChange)[0]
        writer = pd.ExcelWriter(fileName + '.xlsx', engine='xlsxwriter')
        df.to_excel(writer, engine='python', encoding='utf-8')
        writer.save()
        # df.to_excel(fileName + '.xlsx', index=False)
        # os.remove(qttToChange)
        convertAgain()

    else:
        tk.messagebox.showerror(title="Greška", message="Izabrali ste fajl koji nije .qtt, .xyz ili .lst")
        convertAgain()


def convertAgain():
    result = tk.messagebox.askquestion(title="Ponovno konvertovanje",
                                       message="Da li želite da konvertujete još neki fajl?",
                                       icon='question')
    if result == 'yes':
        fileConversion()
    else:
        sys.exit()


fileConversion()
