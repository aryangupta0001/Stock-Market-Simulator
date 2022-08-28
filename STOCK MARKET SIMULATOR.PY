import sqlite3
import tkinter as tk
from tkinter import *
from tkinter import messagebox as msb
from tkinter import ttk
from tkinter.font import BOLD
import pyttsx3
from tkcalendar import Calendar, DateEntry
from tkinter import filedialog
import pandas as pd
import requests


class App(Tk):

    def __init__(self):
        super().__init__()

        self.title("PORTFOLIO MANAGER")
        self.geometry("1024x720")

        self.mymenu = Menu(self)

        self.mymenu.add_command(label = "HOME", command = self.main)

        self.m2 = Menu(self.mymenu, tearoff = 0)
        self.mymenu.add_cascade(label = "TRADES", menu = self.m2)
        self.m2.add_command(label = "All Trades", command = self.AllTransactions)
        self.m2.add_command(label = "Search Trades", command = self.SearchTransaction)

        self.m3 = Menu(self.mymenu, tearoff = 0)
        self.mymenu.add_cascade(label = "SHARES", menu = self.m3)
        self.m3.add_command(label = "View All Scrips", command = self.AllScrips)
        self.m3.add_command(label = "Veiw Shares", command = self.ViewShares)
        self.m3.add_command(label = "BUY", command = self.AllTransactions)
        self.m3.add_command(label = "SELL", command = self.SearchTransaction)


        self.config(menu = self.mymenu)

        self.MFrame = Frame(self)
        self.MFrame.grid()

        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Treeview", foreground='white', background = "PURPLE", font = ("Arial Narrow", 12, "bold"))
        style.configure("Treeview.Heading", fieldbackground = "WHITE", foreground = 'black')
        
        self.main()


    def main(self):

        for i in self.MFrame.winfo_children():
            i.destroy()


        Label(self.MFrame, text = "Your Portfolio Manager", font = ("Segoe Print", 20, "bold"), pady = 20, padx = 10).grid(row = 0, column = 3)

        Button(self.MFrame, text = "NEW SCRIP", width = 20, bg = "ORANGE", fg = "BLACK", font = ("Calibri", 12, "bold"),command = self.AddScrip).grid(row = 4, column = 3)

        Button(self.MFrame, text = "ALL SCRIPS", width = 20, bg = "YELLOW", fg = "BLACK", font = ("Calibri", 12, "bold"), command = self.AllScrips).grid(row = 4, column = 4, pady = 5)

        Button(self.MFrame, text = "BUY", width = 20, bg = "GREEN", fg = "WHITE", font = ("Calibri", 12, "bold"), command = self.BuyShares).grid(row = 6, column = 3)

        Button(self.MFrame, text = "SELL", width = 20, bg = "RED", fg = "WHITE", font = ("Calibri", 12, "bold"), command = self.SellShares).grid(row = 6, column = 4)

        Button(self.MFrame, text = "ALL TRADES", width = 20, bg = "ORANGE", fg = "BLACK", font = ("Calibri", 12, "bold"), command = self.AllTransactions).grid(row = 8, column = 3, pady = 5)

        Button(self.MFrame, text = "SEARCH TRADES", width = 20, bg = "BLUE", fg = "WHITE", font = ("Calibri", 12, "bold"), command = self.SearchTransaction).grid(row = 8, column = 4, pady = 5)

        Button(self.MFrame, text = "VIEW SHARES", width = 20, bg = "BLACK", fg = "WHITE", font = ("Calibri", 12, "bold"), command = self.ViewShares).grid(row = 10, column = 3)


    def AddScrip(self):
        for i in self.MFrame.winfo_children():
            i.destroy()

        self.ScripName = StringVar()
        self.Ltp = DoubleVar()


        Label(self.MFrame, text = "Your Portfolio Manager", font = ("Segoe Print", 20, "bold"), pady = 20, padx = 10).grid(row = 0, column = 3)
        Label(self.MFrame, text = "\t\t").grid(row = 0, column = 0)


        Label(self.MFrame, text = "SCRIP NAME : ", font = ("CALIBRI", 14), width = 12, anchor = W).grid(row  = 1, column = 1)
        self.E1 = Entry(self.MFrame, textvariable = self.ScripName, width = 20, relief = SOLID)
        self.E1.grid(row = 1, column = 2)

        Label(self.MFrame, text = "LTP\t     : ", font = ("CALIBRI", 14), width = 12, anchor = W).grid(row  = 2, column = 1)
        self.E2 = Entry(self.MFrame, textvariable = self.Ltp, width = 20, relief = SOLID)
        self.E2.grid(row = 2, column = 2)

        Label(self.MFrame, text = "(Leave blank to fetch from NSE)", font = ("CALIBRI", 11)).grid(row = 2, column = 3, sticky = W)

        Button(self.MFrame, text = "Submit", command = self.CreateScrip, bg = "YELLOW").grid(row = 3, column = 2)

        Button(self.MFrame, text = "Bulk Add New SCRIP", command = self.BulkAddScrips, bg = "black", fg = "white", font = ("ARIAL", 10, BOLD)).grid(row = 5, column = 2)
        self.test = 1



    def BulkAddScrips(self):
        filename = filedialog.askopenfile(initialdir = "/gui/Quick access", title = "Open", filetypes = (("excel files", "*.xls"), ("excel files", "*.xlsx"), ("excel files", "*.csv")))

        cur.execute("SELECT SNAME FROM SCRIP")
        row = cur.fetchall()
        duplicate_scrips = 0

        try : 
            scrip = pd.read_excel(filename.name, sheet_name = "Sheet1", header = 0)

            data = scrip.values.tolist()

            empty = []
            count = 0

            for i in data:
                count+=1

                if "nan" == str(i[0]):
                    empty.append(str(count))
    
                else:
                    valueok = 0

                    for j in row:
                        if i[0] in j: 
                            duplicate_scrips+=1
                            valueok = 1
                            break
                    
                    if valueok == 0:
                        if "nan" == str(i[1]):
                            ltp = self.FetchLTP(str(i[0]), "ltp")
                        else:
                            ltp = i[1]

                        insert_value = "INSERT INTO SCRIP (SNAME, LTP) VALUES (?, ?)"
                        value = (i[0], ltp)

                        cur.execute(insert_value, value)
                        conn.commit()
            
            index = ",".join(empty)

            pyttsx3.speak(f"Succesfully Added {len(data) - duplicate_scrips - len(empty)} New Scrips")

            if duplicate_scrips>0:
                msb.showinfo("", f"{duplicate_scrips} duplicate scrips not merged")

            if len(index)>0:
                msb.showwarning("Empty scrips", "Scrip Name not found at indexes :\n" + index)

            
        except Exception as e:
            if "'NoneType' object has no attribute 'name'" in str(e):
                pass
    
        
    def CreateScrip(self):
        valueok = 0
        sname = self.ScripName.get().upper()

        if len(sname) > 0:
            try:
                ltp = self.Ltp.get()
                print(ltp)
                if ltp == 0.0:
                    ltp = self.FetchLTP(sname, "ltp")
                    print(ltp)

            except Exception as e:
                if "could not convert string to float:" in str(e):
                    ltp = self.FetchLTP(sname, "ltp")

                    if(ltp == "INVALID SCRIP"):
                        msb.showerror("Scrip Error", "Invalid Scrip Entered, Please check the Scrip & try again")
                        valueok = 1
                    else:
                       valueok = 0
                else:
                    msb.showerror("ERROR", str(e))
                    valueok = 1

        cur.execute("SELECT SNAME FROM SCRIP")
        row = cur.fetchall()

        for i in row:
            if sname in i:
                msb.showerror("SCRIP ERROR", "SCRIP ALREADY EXISTS")
                valueok = 1
                break

        if len(sname) == 0:
            msb.showerror("SCRIP ERROR", "INVALID OR NO SCRIP NAME ENTERED")
            valueok = 1

        else:
            if valueok == 0:
                insert_value = "INSERT INTO SCRIP (SNAME, LTP) VALUES (?, ?)"
                value = (sname, ltp)

                cur.execute(insert_value, value)
                conn.commit()

                pyttsx3.speak("DONE")

                msb.showinfo("SCRIP", sname + " added to your Portfolio")

            self.E1.delete(0, 'end')
            self.E2.delete(0, 'end')
    
    def FetchLTP(self, scrip, info):

        url = "https://cnbc.p.rapidapi.com/symbols/translate"

        querystring = {"symbol":scrip}

        headers = {
            'x-rapidapi-host': "cnbc.p.rapidapi.com",
            'x-rapidapi-key': "7bea777958msh51adf9b02b44661p16dc8bjsndcb53ff7c02c"
            }

        response = requests.request("GET", url, headers=headers, params=querystring)

        KEY = response.text

        if str("\"errorMessage\":\"Invalid Symbol\"") in KEY:
            return "INVALID SCRIP"

        else:
            KEY = KEY.split(",")

            KEY = KEY[0].split(":")[1]
            l = len(KEY)
            KEY = KEY[1:l-1]

            url = "https://cnbc.p.rapidapi.com/symbols/get-summary"

            querystring = {"issueIds" : KEY}

            headers = {
                'x-rapidapi-host': "cnbc.p.rapidapi.com",
                'x-rapidapi-key': "7bea777958msh51adf9b02b44661p16dc8bjsndcb53ff7c02c"
                }

            response = requests.request("GET", url, headers=headers, params=querystring)


            if str("\"symbol\":\"NULL\"") in response.text :
                return "INVALID SCRIP"

            else:
                if info == "ltp":
                    Data = response.text.split("last")
                    Data = Data[1].split("\"")
                    Data = Data[2]
                
                elif info == "prev_close":
                    Data = response.text.split("\"previous_day_closing\":\"")
                    Data = Data[1].split("\"")[0]

                if "," in Data:
                    a = []

                    for i in Data:
                        if i != ",":
                            a.append(i)
                    
                    Data = "".join(a)        

                Data = float(Data)
                return Data


    
    def AllScrips(self):

        cur.execute("SELECT * FROM SCRIP")
        row = cur.fetchall()

        if len(row) == 0:
            msb.showinfo("NO SCRIP", "YOU HAVE NOT ADDED ANY SCRIP YET")

        else:
            for i in self.MFrame.winfo_children():
                i.destroy()

            Label(self.MFrame, text = "Your Portfolio Manager", font = ("Segoe Print", 20, "bold"), pady = 20, padx = 10).grid(row = 0, column = 0, columnspan = 10)
            cur.execute("SELECT * FROM SCRIP")

            tree = ttk.Treeview(self.MFrame, column=("c1", "c2", "c3"), show='headings', height = 31) 

            tree.column("#1", width = 60, anchor = CENTER)
            tree.heading("#1", text="SCRIP ID")

            tree.column("#2", width = 110, anchor = CENTER)
            tree.heading("#2", text="SCRIP NAME")

            tree.column("#3", width = 110, anchor = CENTER)
            tree.heading("#3", text="LTP")

            Label(self.MFrame, text = "\t").grid(row = 1, column = 0)

            tree.grid(row = 1, column = 1, columnspan = 5)

            scrlbar = ttk.Scrollbar(self.MFrame, orient = VERTICAL, command = tree.yview)
            scrlbar.grid(row = 1, column = 7, sticky = NS)

            tree['yscrollcommand'] = scrlbar.set

            for i in row:
                tree.insert("", END, values = i)
            
            self.detailedview = BooleanVar()

            Checkbutton(self.MFrame, text = "Detailed View : ", textvariable = self.detailedview, command = self.DetailedScripView).grid(row = 2, column = 1)
        
    def DetailedScripView(self):
        pass

    def BuyShares(self):
        self.Transaction = "BUY" 
        self.AddShares()

    def SellShares(self):
        self.Transaction = "SELL"
        self.AddShares()

    def AddShares(self):

        for i in self.MFrame.winfo_children():
            i.destroy()
        
        Label(self.MFrame, text = "Your Portfolio Manager", font = ("Segoe Print", 20, "bold"), pady = 20, padx = 10).grid(row = 0, column = 1, columnspan = 5)

        if self.Transaction == "BUY":
            clr = "GREEN"
        else:
            clr = "RED"

        Label(self.MFrame, text = "\tCreating " + self.Transaction + " Order", font = ("CALIBRI", 15, "bold"), foreground = clr).grid(row = 1, column = 1)

        self.ScripName = StringVar()
        self.Qty = IntVar()
        self.LTP = DoubleVar()
        self.Cost = DoubleVar()
        self.Date = StringVar()

        cur.execute("SELECT * FROM  SCRIP")
        row = cur.fetchall()

        self.scriplist = []

        for i in row:
            self.scriplist.append(i[1])

        self.ScripList = ttk.Combobox(self.MFrame, textvariable = self.ScripName)
        self.ScripList["values"] = self.scriplist

        Label(self.MFrame, text = "Select SCRIP\t: ", font = ("CALIBRI", 14)).grid(row  = 2, column = 1, sticky = "W")
        self.ScripList.grid(row = 2, column = 2, sticky = W)

        Label(self.MFrame, text = "Qty.\t\t:", font = ("CALIBRI", 14)).grid(row  = 4, column = 1, sticky = "W")
        self.E1 = Entry(self.MFrame, textvariable = self.Qty, relief = GROOVE)
        self.E1.grid(row = 4, column = 2, sticky = W)

        Label(self.MFrame, text = "LTP\t\t: ", font = ("CALIBRI", 14)).grid(row = 5, column  = 1, sticky = "W")
        self.E2 = Entry(self.MFrame, textvariable = self.LTP, relief = GROOVE)
        self.E2.grid(row = 5, column = 2, sticky = W)
        
        Label(self.MFrame, text = "(Leave blank to fetch from NSE)", font = ("CALIBRI", 11)).grid(row = 5, column = 3, sticky = W)

        Label(self.MFrame, text = "COST\t\t: ", font = ("CALIBRI", 14)).grid(row = 6, column  = 1, sticky = "W")
        self.E3 = Entry(self.MFrame, textvariable = self.Cost, relief = GROOVE)
        self.E3.grid(row = 6, column = 2, sticky = W)

        Label(self.MFrame, text = "(Leave blank if same as LTP)", font = ("CALIBRI", 11)).grid(row = 6, column = 3, sticky = W)

        Label(self.MFrame, text = "DATE\t\t: ", font = ("CALIBRI", 14)).grid(row = 7, column  = 1, sticky = "W")
        cal = DateEntry(self.MFrame, textvariable = self.Date, width = 16, background= "magenta3", foreground= "white", bd = 2)
        cal.grid(row = 7, column = 2, sticky=  W)

        Button(self.MFrame, text = "SUBMIT", command = self.RecordTransaction).grid(row = 12, column = 2)


    def RecordTransaction(self):

        valueok = 0

        if valueok == 0:
            ScripName = self.ScripName.get().upper()
            if len(ScripName) == 0:
                msb.showerror("SCRIP ERROR", "NO SCRIP SELECTED")
                valueok = 1

            else:

                if ScripName not in self.scriplist:
                    msb.showerror("SCRIP ERROR", "SCRIP NOT FOUND")
                    valueok = 1

            Transaction = self.Transaction.upper()
        
        if valueok == 0:

            try:
                TQty = int(self.Qty.get())
                print(TQty)
                if TQty == 0:
                    msb.showerror("Value Error", "INVALID or NO Quantity entered")
                    valueok = 1

            except Exception as e:
                if "expected floating-point number but got" in str(e):
                    msb.showerror("Value Error", "INVALID or NO Quantity entered")
                    valueok = 1
                else:
                    msb.showerror("", str(e))
                    valueok = 1

        if valueok == 0:

            try:
                TLTP = self.LTP.get()
                if TLTP == 0:
                    TLTP = self.FetchLTP(ScripName, "ltp")

            except Exception as e:
                if "expected floating-point number but got" in str(e):
                    TLTP = self.FetchLTP(ScripName, "ltp")
                    valueok = 0
                else:
                    msb.showerror("", str(e))
                    valueok = 1

            TCOST = TLTP

        if valueok == 0:
            try:
                TCOST = self.Cost.get()
                if TCOST == 0:
                    TCOST = TLTP

            except Exception as e:
                if "expected floating-point number but got" in str(e):
                    TCOST = TLTP
                    valueok = 0
                else:
                    msb.showerror("", str(e))
                    valueok = 1

        if valueok == 0:
            self.ScripList.set('')
            self.E1.delete(0, "end")
            self.E2.delete(0, "end")
            self.E3.delete(0, "end")

            Date = self.Date.get()
            DATE = []
            DATE = Date.split("/")

            cur.execute("SELECT * FROM TRANSACTIONS WHERE SNAME IS '" + ScripName + "'")
            row = cur.fetchall()

            if len(row) == 0:
                Total_Qty = 0
                Used_Qty = 0 
                Net_Qty = 0
                Current_Investment = 0
                Total_Investment = 0
                BProfit = 0

            else:
                l = len(row)
                current = row[l-1]
                Total_Qty = current[9]
                Used_Qty = current[10]
                Net_Qty = current[11]
                Current_Investment = current[13]
                Total_Investment = current[14]
                AValue = current[15]
                BProfit = current[18]

            DD = str(DATE[1])
            MM = str(DATE[0])
            YYYY = str(DATE[2])
            SNAME = ScripName
            SIDE = Transaction
            QTY = TQty
            LTP = TLTP
            COST = TCOST

            if SIDE == "BUY":
                TOTALQTY = Total_Qty + QTY
                USEDQTY = Used_Qty
                NETQTY = Net_Qty + QTY
                TINVESTMENT = round(QTY * COST, 2)
                CURRENT_INVESTMENT = round(Current_Investment + TINVESTMENT, 2)
                TOTAL_INVESTMENT = round(Total_Investment + TINVESTMENT, 2)
                AVALUE = round(TOTAL_INVESTMENT / TOTALQTY, 2)
                TPROFIT = 0

            else:
                if QTY > Net_Qty:
                    msb.showerror("Transaction Error", "You do not have enough shares")
                    valueok = 1

                else:
                    TOTALQTY = Total_Qty
                    USEDQTY = Used_Qty + QTY
                    NETQTY = Net_Qty - QTY
                    TINVESTMENT = round(-1 * QTY * COST, 2)
                    TOTAL_INVESTMENT = round(Total_Investment, 2)
                    TPROFIT = round(QTY * (COST - AValue), 2)

                    if NETQTY == 0:
                        AVALUE = 0
                        CURRENT_INVESTMENT = 0
                    
                    else:

                        a = NETQTY
                        amount = 0

                        for i in range(len(row), 0, -1):
                            A = row[i-1]

                            if "BUY" in A:
                                if a>=A[6]:
                                    amount+=A[12]
                                    a-=A[6]

                                else:
                                    amount = amount + round((A[12]/A[6]*a), 2)
                                    a = 0

                        AVALUE = round(amount/NETQTY, 2)
                        CURRENT_INVESTMENT = round(amount, 2)
            
            if valueok == 0:
                CURRENT_VALUE = round(NETQTY * LTP, 2)
                BPROFIT = BProfit + TPROFIT
                HPROFIT = round(NETQTY * (LTP - AVALUE), 2)
                TOTAL_PROFIT = BPROFIT + HPROFIT
                PER_PROFIT = round(TOTAL_PROFIT * 100 / TOTAL_INVESTMENT, 2)


                insert_value = f"INSERT INTO  TRANSACTIONS(DD, MM, YYYY, SNAME, TSIDE, TQTY, LTP, COST, TOTALQTY, USEDQTY, NETQTY, TINVESTMENT, CURRENT_INVESTMENT, TOTAL_INVESTMENT, AVALUE, CURRENT_VALUE, TPROFIT, BPROFIT, HPROFIT, TOTAL_PROFIT, PER_PROFIT) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                value = (DD, MM, YYYY, SNAME, SIDE, QTY, LTP, COST, TOTALQTY, USEDQTY, NETQTY, TINVESTMENT, CURRENT_INVESTMENT, TOTAL_INVESTMENT, AVALUE, CURRENT_VALUE, TPROFIT, BPROFIT, HPROFIT, TOTAL_PROFIT, str(PER_PROFIT) + " %")

                cur.execute(insert_value, value)
                conn.commit()

                cur.execute("SELECT MAX(T_ID) FROM TRANSACTIONS")
                a = cur.fetchall()
                a = a[0]
                a = a[0]
                
                
                previous = self.FetchLTP(ScripName, "prev_close")

                pyttsx3.speak("TRANSACTION RECORDED")
                msb.showinfo("Transaction Recorded", SIDE + " Transaction recorded with TID  : " + str(a))

                # UPDATING SCRIP TABLE

                DP_L = round(NETQTY*(LTP - previous), 2)

                update_command = f'''UPDATE SCRIP
                SET LTP = ?,
                TOTALQTY = ?,
                USEDQTY = ?,
                NETQTY = ?,
                AVALUE = ?,
                TOTAL_INVESTMENT = ?,
                CURRENT_INVESTMENT = ?,
                CURRENT_VALUE = ?,
                DP_L = ?,
                HP_L = ?,
                BP_L = ?,
                TP_L = ?,
                PER_P_L = ?
                WHERE SNAME = ?'''

                update_data = (LTP, TOTALQTY, USEDQTY, NETQTY, AVALUE, TOTAL_INVESTMENT, CURRENT_INVESTMENT, CURRENT_VALUE, DP_L, HPROFIT, BPROFIT, TOTAL_PROFIT, PER_PROFIT, SNAME)

                cur.execute(update_command, update_data)
                conn.commit()

        else:
            self.AddShares()


    def AllTransactions(self):

        for i in self.MFrame.winfo_children():
            i.destroy()

        Label(self.MFrame, text = "\t\t").grid(row = 0, column = 0)

        Label(self.MFrame, text = "Your Portfolio Manager", font = ("Segoe Print", 20, "bold"), pady = 20, padx = 10).grid(row = 0, column = 0, columnspan = 20)

        cur.execute("SELECT * FROM TRANSACTIONS")

        A = cur.fetchall()

        B = []

        for a in A:
            b = []
            b.append(a[0])
            b.append(str(a[1]) + "-" + str(a[2]) + "-" + str(a[3]))

            for i in range(4, len(a)):
                b.append(a[i])

            B.append(b)

        self.ShowTransaction(B)

    def SearchTransaction(self):

        self.DAY = IntVar()
        self.MONTH = IntVar()
        self.YEAR = IntVar()
        self.SCRIP = StringVar()
        self.BUYSELL = StringVar()

        day = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31)
        month = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)

        cur.execute("SELECT SNAME FROM SCRIP")
        row = cur.fetchall()

        self.scriplist = []

        for i in row:
            self.scriplist.append(i[0])

        b = ("BUY", "SELL")

        for i in self.MFrame.winfo_children():
            i.destroy()

        Label(self.MFrame, text = "Your Portfolio Manager", font = ("Segoe Print", 20, "bold"), pady = 20, padx = 10).grid(row = 0, column = 0, columnspan = 10)

        Label(self.MFrame, text = "\t\t\t").grid(row = 1, column = 0)
        Label(self.MFrame, text = "Search Filters : ", font = ("CALIBRI", 14)).grid(row = 2, column = 1)  

        Label(self.MFrame, text = "DD\t\t:").grid(row = 3, column = 1)
        ttk.Combobox(self.MFrame, textvariable = self.DAY, values = day, width = 10).grid(row = 3, column = 2, sticky = W)

        Label(self.MFrame, text = "MM\t\t:").grid(row = 4, column = 1)
        ttk.Combobox(self.MFrame, textvariable = self.MONTH, values =month, width = 10).grid(row = 4, column = 2, sticky = W)

        Label(self.MFrame, text = "YYYY\t\t:").grid(row = 5, column = 1)
        Entry(self.MFrame, textvariable = self.YEAR, width = 13, relief = SUNKEN).grid(row = 5, column = 2, sticky = W)

        Label(self.MFrame, text = "SCRIP\t\t:").grid(row = 6, column = 1)
        ttk.Combobox(self.MFrame, textvariable = self.SCRIP, values = self.scriplist).grid(row = 6, column = 2, sticky = W)

        Label(self.MFrame, text = "BUY / SELL\t:").grid(row = 7, column = 1)
        ttk.Combobox(self.MFrame, textvariable = self.BUYSELL, values = b, width = 10).grid(row = 7, column = 2, sticky = W)

        Button(self.MFrame, text = "VIEW", command = self.View).grid(row = 9, column = 1, sticky = E, pady = 5)


    def View(self):

        search = ["SELECT * FROM TRANSACTIONS "]

        i = 0

        try:
            day = self.DAY.get()
            if day>0:
                search.append(" WHERE DD IS '" + str(day) + "'")
                i = 1

        except Exception as e:
                if "expected floating-point number but got" in str(e):
                    pass
                else:
                    msb.showerror("", str(e))
        

        try:
            month = self.MONTH.get()
            if month>0:
                
                if i == 0:
                    search.append(" WHERE MM IS '" + str(month) + "'")
                else:
                    search.append(" AND MM IS '" + str(month) + "'")
                
                i = 1

        except Exception as e:
                if "expected floating-point number but got" in str(e):
                    pass
                else:
                    msb.showerror("", str(e))

        try:
            year = self.YEAR.get()
            if year>0:
                
                if i == 0:
                    search.append(" WHERE YYYY IS '" + str(year) + "'")
                else:
                    search.append(" AND YYYY IS '" + str(year) + "'")
                
                i = 1

        except Exception as e:
                if "expected floating-point number but got" in str(e):
                    pass
                else:
                    msb.showerror("", str(e))

        scrip = self.SCRIP.get().upper()

        if len(scrip)>0:
            if scrip not in self.scriplist:
                msb.showerror("SCRIP ERROR", "INVALID OR NO SCRIP ENTERED")

            else:
                if i == 0:
                    search.append(" WHERE SNAME IS '" + str(scrip) + "'")
                else:
                    search.append(" AND SNAME IS '" + str(scrip) + "'")

        buysell = self.BUYSELL.get().upper()

        if len(buysell)>0:
            if buysell not in ["BUY", "SELL"]:
                msb.showerror("SCRIP ERROR", "INVALID TRANSACTION SIDE ENTERED")

            else:
                if i == 0:
                    search.append(" WHERE TSIDE IS '" + str(buysell) + "'")
                else:
                    search.append(" AND TSIDE IS '" + str(buysell) + "'")

        cur.execute(" ".join(search))

        A = cur.fetchall()

        if (len(A)>0):
            
            B = []

            for a in A:
                b = []
                b.append(a[0])
                b.append(str(a[1]) + "-" + str(a[2]) + "-" + str(a[3]))

                for i in range(4, len(a)):
                    b.append(a[i])

                B.append(b)
            
            for i in self.MFrame.winfo_children():
                i.destroy()

            Label(self.MFrame, text = "Your Portfolio Manager", font = ("Segoe Print", 20, "bold"), pady = 20, padx = 10).grid(row = 0, column = 0, columnspan = 10)

            self.ShowTransaction(B)
        
        else:
            msb.showinfo("", "No Record Found")
            self.DAY.set('')
            self.MONTH.set('')
            self.YEAR.set('')
            self.SCRIP.set('')
            self.BUYSELL.set('')

    def ShowTransaction(self, lst):
        
        blank=  [""]

        tree = ttk.Treeview(self.MFrame, column=("c1", "c2", "c3", "c4","c5", "c6", "c7", "c8", "c9", "c10", "c1", "c12", "c13", "c14", "c15", "c16" ,"c17", "c18", "c19", "c20"), show='headings', height = 30)

        tree.column("#1", width = 40, anchor = CENTER)
        tree.heading("#1", text="ID")

        tree.column("#2", width = 55, anchor = CENTER)
        tree.heading("#2", text="DATE")

        tree.column("#3", width = 100, anchor = CENTER)
        tree.heading("#3", text="SCRIP")

        tree.column("#4", width = 60, anchor = CENTER)
        tree.heading("#4", text="BUY/SELL")

        tree.column("#5", width = 50, anchor = CENTER)
        tree.heading("#5", text="QTY.")

        tree.column("#6", width = 70, anchor = CENTER)
        tree.heading("#6", text="LTP")

        tree.column("#7", width = 60, anchor = CENTER)
        tree.heading("#7", text="PRICE")

        tree.column("#8", width = 60, anchor = CENTER)
        tree.heading("#8", text="TOTAL QTY.")

        tree.column("#9", width = 60, anchor = CENTER)
        tree.heading("#9", text="USED QTY.")

        tree.column("#10", width = 80, anchor = CENTER)
        tree.heading("#10", text="NET QTY.")

        tree.column("#11", width = 100, anchor = CENTER)
        tree.heading("#11", text="TINVESTMENT")

        tree.column("#12", width = 80, anchor = CENTER)
        tree.heading("#12", text="CURRENT INVT.")

        tree.column("#13", width = 100, anchor = CENTER)
        tree.heading("#13", text="TOTAL INVT.")

        tree.column("#14", width = 110, anchor = CENTER)
        tree.heading("#14", text="AVG. VALUE")

        tree.column("#15", width = 90, anchor = CENTER)
        tree.heading("#15", text="CURRENT VALUE")

        tree.column("#16", width = 80, anchor = CENTER)
        tree.heading("#16", text="TPROFIT")

        tree.column("#17", width = 80, anchor = CENTER)
        tree.heading("#17", text="BPROFIT")

        tree.column("#18", width = 80, anchor = CENTER)
        tree.heading("#18", text="HPROFIT")

        tree.column("#19", width = 80, anchor = CENTER)
        tree.heading("#19", text="TOTAL PROFIT")

        tree.column("#20", width = 80, anchor = CENTER)
        tree.heading("#20", text="% PROFIT")

        tree.grid(row = 1, column = 0)

        scrlbar = ttk.Scrollbar(self.MFrame, orient = VERTICAL, command = tree.yview)
        scrlbar.grid(row = 1, column = 1, sticky = NS)

        tree['yscrollcommand'] = scrlbar.set

        for i in lst:
            tree.insert("", END, values = i)
            tree.insert("", END, values = blank[0])

    def ViewShares(self):

        for i in self.MFrame.winfo_children():
            i.destroy()

        Label(self.MFrame, text = "Your Portfolio Manager", font = ("Segoe Print", 20, "bold"), pady = 20, padx = 10).grid(row = 0, column = 0, columnspan = 5)


        self.ScripName = StringVar()

        cur.execute("SELECT * FROM  SCRIP")
        row = cur.fetchall()

        a = []

        for i in row:
            a.append(i[1])

        ScripList = ttk.Combobox(self.MFrame, textvariable = self.ScripName)
        ScripList["values"] = a

        Label(self.MFrame, text = "Scrip : ", font = "CALIBRI").grid(row = 1, column = 1)
        ScripList.grid(row = 1, column = 2)
        Button(self.MFrame, text = "VIEW", relief = SOLID, command = self.Display).grid(row = 2,  column = 2)

    def Display(self):

        scrip = self.ScripName.get().upper()

        cur.execute("SELECT SNAME FROM SCRIP")

        row = cur.fetchall()

        scriplist = []

        for i in row:
            scriplist.append(i[0])

        if scrip not in scriplist:

            msb.showerror("SCRIP ERROR", "SCRIP NOT FOUND")
            self.ViewShares()

        else:
            cur.execute("SELECT * FROM SCRIP WHERE SNAME IS '" + scrip + "'")

            self.ShareInfo = cur.fetchall()[0]

            if str(self.ShareInfo[5]) == "None" or self.ShareInfo[5] == 0:
                msb.showinfo("", "NO SHARES OF " + scrip + " EXISTS IN YOUR PORTFOLIO !!!")
                self.ViewShares()

            else:

                for i in self.MFrame.winfo_children():
                    i.destroy()

                Label(self.MFrame, text = "Your Portfolio Manager", font = ("Segoe Print", 20, "bold"), pady = 20, padx = 10).grid(row = 0, column = 0, columnspan = 5)

                Label(self.MFrame, text = "        ", font = "CALIBRI").grid(row = 1, column = 1)
                Label(self.MFrame, text = "        ", font = "CALIBRI").grid(row = 1, column = 2)


                b = "₹ " + str(self.ShareInfo[2])
                c = self.ShareInfo[5]
                d = self.ShareInfo[11]
                e = self.ShareInfo[8]
                f = self.ShareInfo[9]
                g = round(d/e*100, 2)
                h = self.ShareInfo[6]
                i = self.ShareInfo[3]
                j = self.ShareInfo[4]
                k = self.ShareInfo[10]

                if float(d)>0:
                    txtclr = "GREEN"
                elif float(d) == 0:
                    txtclr = "GREY"
                else:
                    txtclr = "RED"

                Label(self.MFrame, text = scrip, font = ("ARIAL", 16, BOLD), width = 22, anchor = W).grid(row = 4, column = 3, sticky = W)

                Label(self.MFrame, text = "LTP ", font = ("ARIAL", 14), width = 18, anchor = E).grid(row = 4, column = 4, sticky = W)
                Label(self.MFrame, text = b, font = ("ARIAL", 15)).grid(row = 4, column = 4, sticky = E)

                Button(self.MFrame, text = "Update", command = self.UpdateLTP).grid(row  = 4, column = 5, sticky = W)

                Label(self.MFrame, text = "  NET QTY", font = ("ARIAL", 14)).grid(row = 5, column = 3, sticky = W)
                Label(self.MFrame, text = str(c) + "  ", font = ("ARIAL", 14, "bold"), anchor = E, width = 16).grid(row = 5, column = 3, sticky = E)

                Label(self.MFrame, text = "  P&L  ", font = ("ARIAL", 14)).grid(row = 5, column = 4, sticky = W)
                Label(self.MFrame, text = "₹  " + str(d) + " (" + str(g) +" %) ", foreground = txtclr, font = ("ARIAL", 14, "bold")).grid(row = 5, column = 4, sticky = E)

                Label(self.MFrame, text =  "  Average Price", font = ("ARIAL", 14)).grid(row = 9, column = 3, sticky = W)
                Label(self.MFrame, text = str(h) + "  ", font = ("ARIAL", 14, BOLD), anchor = E).grid(row = 9, column = 3, sticky = E)

                Label(self.MFrame, text =  "  Day P&L", font = ("ARIAL", 14)).grid(row = 9, column = 4, sticky = W)
                Label(self.MFrame, text = str(k) + " ", font = ("ARIAL", 14, BOLD), anchor = E).grid(row = 9, column = 4, sticky = E)

                Label(self.MFrame, text =  "  Total Qty", font = ("ARIAL", 14)).grid(row = 10, column = 3, sticky = W)
                Label(self.MFrame, text = str(i) + "  ", font = ("ARIAL", 14, BOLD), anchor = E).grid(row = 10, column = 3, sticky = E)

                Label(self.MFrame, text =  "  Used Qty", font = ("ARIAL", 14)).grid(row = 10, column = 4, sticky = W)
                Label(self.MFrame, text = str(j) + "  ", font = ("ARIAL", 14, BOLD), anchor = E).grid(row = 10, column = 4, sticky = E)

                Label(self.MFrame, text =  "  Current Invt.", font = ("ARIAL", 14)).grid(row = 11, column = 3, sticky = W)
                Label(self.MFrame, text = str(e) + "  ", font = ("ARIAL", 14, BOLD), anchor = E).grid(row = 11, column = 3, sticky = E)

                Label(self.MFrame, text =  "  Current Value", font = ("ARIAL", 14)).grid(row = 11, column = 4, sticky = W)
                Label(self.MFrame, text = str(f) + "  ", font = ("ARIAL", 14, BOLD), anchor = E).grid(row = 11, column = 4, sticky = E)

                Label(self.MFrame).grid(row = 12)
                Label(self.MFrame, font = ("ARIAL", 16, BOLD), width = 22).grid(row = 12, column = 4)

                Button(self.MFrame, text = "BUY", fg = "WHITE", bg = "GREEN", font = ("ARIAL", 16, BOLD), width = 20, command = self.BuyShares).grid(row = 13, column = 3)
                Button(self.MFrame, text = "SELL", fg = "WHITE", bg = "RED", font = ("ARIAL", 16, BOLD), width = 20, command = self.AddShares).grid(row = 13, column = 4)

    def UpdateLTP(self):

        scrip = self.ScripName.get().upper()

        ltp = self.FetchLTP(scrip, "ltp")
        previous = self.FetchLTP(scrip, "prev_close")

        current_value = round(self.ShareInfo[5]*ltp, 2)
        hp_l = round(current_value - self.ShareInfo[8], 2)
        tp_l = round(hp_l + self.ShareInfo[11], 2)
        per_p_l = round(tp_l * 100 / self.ShareInfo[7], 2)

        dp_l = round(self.ShareInfo[5]*(ltp - previous), 2)

        update_command = f'''UPDATE SCRIP SET
            LTP = ?,
            CURRENT_VALUE = ?,
            DP_L = ?,
            HP_L = ?,
            TP_L = ?,
            PER_P_L = ?
            WHERE SNAME = ?'''

        update_data = (ltp, current_value, dp_l, hp_l, tp_l, per_p_l, scrip)

        cur.execute(update_command, update_data)
        conn.commit()

        Label(self.MFrame, text = "Your Portfolio Manager", font = ("Segoe Print", 20, "bold"), pady = 20, padx = 10).grid(row = 0, column = 0, columnspan = 5)

        self.Display()

        pyttsx3.speak("UPDATED")


    def updateltp(self):
        scrip = self.ScripName.get().upper()

        try:
            ltp = float(self.newltp.get())
        
        except Exception as e:
            if "could not convert string to float:" in str(e):
                msb.showerror("Value Error", "INVALID or NO  LTP entered")

        else:
            cur.execute("UPDATE SCRIP SET LTP = '" + str(ltp) + "' where SNAME IS '" + scrip + "'")
            pyttsx3.speak("UPDATED")

def PORTFOLIO():
    app = App()
    app.mainloop()


if __name__ == "__main__":

    conn = sqlite3.connect("PORTFOLIO.db")

    cur = conn.cursor()

    create_table_scrip = '''CREATE TABLE SCRIP
        (S_ID               INTEGER     PRIMARY KEY     AUTOINCREMENT,
        SNAME               TEXT,
        LTP                 REAL,
        TOTALQTY            INTEGER,
        USEDQTY             INTEGER,
        NETQTY              INTEGER, 
        AVALUE              REAL,
        TOTAL_INVESTMENT    REAL,
        CURRENT_INVESTMENT  REAL,
        CURRENT_VALUE       REAL,
        DP_L                REAL,
        HP_L                REAL,
        BP_L                REAL,
        TP_L                REAL,
        PER_P_L             REAL
        )'''

    try:
        cur.execute(create_table_scrip)
        conn.commit()

    except Exception as e:
        if str(e) == "table SCRIP already exists":
            pass

        else:
            msb.showwarning("Connection", "Error !!! Database Connection Failed\n" + str(e))

    create_table_transaction = '''CREATE TABLE TRANSACTIONS
        (T_ID   INTEGER     PRIMARY KEY     AUTOINCREMENT,
        DD                  TEXT,
        MM                  TEXT,
        YYYY                TEXT,
        SNAME               TEXT,
        TSIDE               TEXT,
        TQTY                INTEGER,
        LTP                 REAL,
        COST                REAL,
        TOTALQTY            INTEGER,
        USEDQTY             INTEGER,
        NETQTY              INTEGER,
        TINVESTMENT         REAL,
        CURRENT_INVESTMENT  REAL,   
        TOTAL_INVESTMENT    REAL,
        AVALUE              REAL,
        CURRENT_VALUE       REAL,
        TPROFIT             REAL,
        BPROFIT             REAL,
        HPROFIT             REAL,
        TOTAL_PROFIT        REAL,
        PER_PROFIT          TEXT    );'''

    try:
        cur.execute(create_table_transaction)
        conn.commit()

    except Exception as e:
        if str(e) == "table TRANSACTIONS already exists":
            pass

        else:
            msb.showwarning("Connection Error", "Error !!! Database Connection Failed, Can't fetch Transactions\n" + str(e))

    PORTFOLIO()