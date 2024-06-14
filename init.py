import tkinter as tk
import os
import re

def getMaster():
    files = [f for f in os.listdir('.') if (os.path.isfile(f) and "xlsx" in f)]
    if len(files) != 1:
        raise ValueError("place exactly one master file in xlsx format at root directory")
    masterFile = files[0]
    return masterFile

def timeFormCheck(time):
    return re.match(r"\d+/\d+/\d+ \d+:\d+:\d+", time)

def main():
    window = tk.Tk()

    lblDate = tk.Label(text="Convocation Date: mm/dd/yyyy")
    entDate = tk.Entry()

    lblStart = tk.Label(text="start time: hh:mm")
    entStart = tk.Entry()
    lblEnd = tk.Label(text="end time: hh:mm")
    entEnd = tk.Entry()


        

    def runScript():
        masterFile = getMaster()
        startTime = f"{entDate.get()} {entStart.get()}:00"
        endTime = f"{entDate.get()} {entEnd.get()}:00"
        if not (timeFormCheck(startTime) and timeFormCheck(endTime)):
            raise ValueError("time form not valid")

        os.system(f"python absentList.py \"{masterFile}\" \"{startTime}\" \"{endTime}\"")
        window.destroy()




    btnConfirm = tk.Button(text="confirm", command=runScript)

    lblDate.pack()
    entDate.pack()
    lblStart.pack()
    entStart.pack()
    lblEnd.pack()
    entEnd.pack()
    btnConfirm.pack()
    window.mainloop()

main()