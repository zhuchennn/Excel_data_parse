import tkinter as tk

window1 = tk.Tk()
canvas1 = tk.Canvas(window1, width=400, height=300, relief='raised')
canvas1.pack()

title = tk.Label(window1, text='Data Extraction Inputs')
title.config(font=('helvetica', 14))
canvas1.create_window(200, 25, window=title)

title1 = tk.Label(window1, text = 'File Path Input')
title1.config(font=('helvetica', 10))
canvas1.create_window(200, 80, window=title1)
direct = tk.Entry(window1)
canvas1.create_window(200, 115, window=direct)

title2 = tk.Label(window1, text= 'File Name Input')
title2.config(font=('helvetica', 10))
canvas1.create_window(200, 155, window=title2)
filename = tk.Entry(window1)
canvas1.create_window(200, 190, window=filename)

def close_window():
    global direct
    direct = direct.get()
    global filename
    filename = filename.get()
    window1.destroy()

button1 = tk.Button(text= 'confirm', command = close_window)
canvas1.create_window(200, 230, window=button1)

window1.mainloop()

print(direct)
print(filename)

