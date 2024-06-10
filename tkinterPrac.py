import tkinter as tk
import os
import re

files = [f for f in os.listdir('.') if (os.path.isfile(f) and "xlsx" in f)]

if len(files) != 1:
    raise ValueError("place exactly one master file in xlsx format at root directory")

masterFile = files[0]
startTime = ''
endTime = ''

window = tk.Tk()

lblStart = tk.Label(text="start time: mm/dd/yyyy hh:mm")
entStartTime = tk.Entry()
lblEnd = tk.Label(text="end time: mm/dd/yyyy hh:mm")
entEndTime = tk.Entry()

def timeFormCheck(time):
    return re.match(r"\d+/\d+/\d+ \d+:\d+:\d+", time)
    

def runScript():
    global startTime
    global endTime
    startTime = entStartTime.get()
    endTime = entEndTime.get()
    if not timeFormCheck(startTime) and timeFormCheck(endTime):
        raise ValueError("time form not valid")
    os.system(f"python absentList.py {masterFile} \"{startTime}\" \"{endTime}\"")
    print("program completed, result can be found under \"result\" directory")



btnConfirm = tk.Button(text="confirm", command=runScript)

lblStart.pack()
entStartTime.pack()
lblEnd.pack()
entEndTime.pack()
btnConfirm.pack()
window.mainloop()
