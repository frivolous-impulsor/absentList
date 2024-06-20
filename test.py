import tkinter as tk
import os
import re

startTime: str
endTime: str

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
        global startTime
        global endTime
        startTime = f"{entDate.get()} {entStart.get()}:00"
        endTime = f"{entDate.get()} {entEnd.get()}:00"
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
    return (startTime, endTime)

print(main())