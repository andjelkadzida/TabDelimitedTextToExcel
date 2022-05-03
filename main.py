import os
import sys
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import xlsxwriter

root = tk.Tk()
root.title("Konvertor xyz, qtt i lst ekstenzija u Excel")
root.iconbitmap(os.path.join(os.path.dirname(os.getcwd()), 'launcher.ico'))
root.wm_iconbitmap(os.path.join(os.path.dirname(os.getcwd()), 'launcher.ico'))
root.withdraw()


def fileConversion():
    extensionToChange = filedialog.askopenfilename()
    if extensionToChange.endswith('.qtt') or (extensionToChange.endswith('.xyz') or extensionToChange.endswith('.lst')):
        while not os.path.isfile(extensionToChange):
            print("Greška: " + extensionToChange + " nije validna putanja do fajla. Pokušajte ponovo...")
            extensionToChange = filedialog.askopenfilename()
        df = pd.read_csv(extensionToChange, sep='\t' or '\s*' or '\s+' or 't', engine='python', encoding='cp1252')
        df.rename(columns={'Unnamed: 0': ''}, inplace=True)
        df.rename(columns={'Unnamed: 33': ''}, inplace=True)
        if extensionToChange.endswith('.lst') or (extensionToChange.endswith('qtt')):
            fileName = os.path.splitext(extensionToChange)[0]
            writer = pd.ExcelWriter(fileName + '.xlsx', engine='xlsxwriter')
            df.to_excel(writer, engine='python', encoding='utf-8', index=False)
        else:
            fileName = os.path.splitext(extensionToChange)[0]
            writer = pd.ExcelWriter(fileName + '.xlsx', engine='xlsxwriter')
            df.to_excel(writer, engine='python', encoding='utf-8')
        writer.save()
        # df.to_excel(fileName + '.xlsx', index=False)
        # os.remove(extensionToChange)
        convertAgain()

    else:
        res = tk.messagebox.askquestion(title="Greška",
                                        message="Izabrali ste fajl koji nije .qtt, .xyz ili .lst.\nDa li želite da pokušate ponovo?",
                                        icon='error')
        if res == 'yes':
            convertAgain()
        else:
            sys.exit()


def convertAgain():
    result = tk.messagebox.askquestion(title="Ponovno konvertovanje",
                                       message="Da li želite da konvertujete još neki fajl?",
                                       icon='question')
    if result == 'yes':
        fileConversion()
    else:
        sys.exit()


fileConversion()
