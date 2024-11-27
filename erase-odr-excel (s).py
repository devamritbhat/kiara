import pandas as pd
from tkinter import Label, Tk, messagebox

exl = "Z:\\MARKETING  TEAM\\Transactions\\EMR Import Excel\\ORDER IMPORT EXCEL (S).xlsx"

root = Tk()
root.iconify()

df = pd.read_excel(exl, sheet_name=0)

with pd.ExcelWriter(exl, engine="openpyxl", mode='a', if_sheet_exists="overlay") as writer:
    pd.DataFrame([[""]*36 for i in range(len(df))]).to_excel(writer, startcol=0, startrow=1, index=False, header=False)

def confirm():
    w = Label(root, text ='Kiara Jewellery', font = "50")  
    w.pack() 
    messagebox.showinfo("Order Raising", "Excel Cleared " + str(len(df)) + " Rows") 

confirm()
