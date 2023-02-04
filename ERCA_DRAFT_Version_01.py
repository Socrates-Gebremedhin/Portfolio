# Name: ERCA data retrieving tool
# Designed by Socrates G/medhin
# period: 29/June/2021

#-------import important classes---------------
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from tkinter.ttk import *
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
import tkinter.simpledialog
import math
import numpy as np
import pandas as pd
from functools import partial
import random
import time
from linker_class import linker
#-----------------------------------------------

windows = Tk()
window1 = Tk()
windows.title("ERCA Data Retrieving Tool")
#-----------------------------------------------

class ERCA_tool:
    def __init__(self, window, window2):#initialize all important variables and widgets
        self.window = window
        self.window2 = window2
        self.window2.overrideredirect('true')
        window2.withdraw()
        self.progress = Progressbar(self.window2)#, orient = "horizontal", length = 100, mode = 'determinate')
        self.progress.grid(row =1 , column =1)
        self.window2.attributes("-topmost", True)
        menubar=Menu(window)
        window.config(menu=menubar)
        High_level= Menu(menubar,tearoff=0)
        menubar.add_cascade(label="File", menu=High_level)
        High_level.add_command(label="Load files", command=self.load_files)
        High_level.add_command(label="Clean files", command=self.clean)
        Hs_label=Label(window, text = "HS code" )
        #Important variables
        self.hs_codes = list()
        self.columns = ["Year", "Month", "Day", "HS Code", "HS Description", "Country (Origin)", "Country (Consignment)",
                        "Gross Weight (Kg)", "Net Weight (Kg)", "CIF/FOB Value (ETB)", "CIF/FOB Value (USD)"] # Default columns, for now only columns to analyze
        self.df_list = list()
        self.gc_years = list()
        self.use_gc = 0
        self.use_ec = 0
        self.d_type = ["int", "int", "int", "str", "str", "str", "str", "float", "float", "float", "float"]
        self.df_main = pd.DataFrame()
        self.df_main_ec = pd.DataFrame()
        self.hs_code_file_path = ''
        hs_description = StringVar()
        self.hs_desc_label = Label(window, text = hs_description)
        self.hs_items = list()
        self.hs_descr = list()
        self.hs_items_desc=list()
        self.hs_filtered = ""
        self.all_anlys_cols = ["Gross Weight (Kg)", "Net Weight (Kg)", "CIF/FOB Value (ETB)", "CIF/FOB Value (USD)"]
        self.anlys_cols = list()
        #-----------------
        wid_b = 15
        wid_L = 35
        wid_E = 33
        wid_C = 30
        wid_La =15
        self.hs_extract_button=Button(window,text="Extract HS Code",command=self.extract_hs_code,width= wid_b )
        self.hs_combobox=ttk.Combobox(window,width= wid_C)
        self.hs_combobox["values"] = self.hs_codes
        self.hs_combobox.bind("<<ComboboxSelected>>", self.fil_item)
        self.hs_combobox.bind("<Key>", self.fil_item)
        self.period_combobox=ttk.Combobox(window,values=["Year","Month"],width= wid_C)
        self.period_combobox.bind("<<ComboboxSelected>>",self.change_trend_scale)
        self.period_label=Label(window,text="Yearly/Monthly",width= wid_La)
        self.gc_ec_combobox=ttk.Combobox(window,state='readonly',width= wid_C)
        self.gc_ec_combobox['values']=["G.C.","E.C."]
        self.gc_ec_combobox.bind("<<ComboboxSelected>>",self.change_calander_type)
        self.gc_ec_label=Label(window,text="Calander Type",width= wid_La)
        self.analyze_button=Button(window,text="Manage",command=self.manage,width= wid_b)
        self.pivot_label = Label(window, text = "Pivot By",width= wid_La)
        self.pivot_by_combobox=ttk.Combobox(window,state='readonly',width= wid_C)
        self.pivot_by_combobox['values']=["Country (Origin)", "Country (Consignment)"]
        self.pivot_by_combobox.bind("<<ComboboxSelected>>",self.pivot_by)
        self.Hs_listbox = Listbox(window,width= wid_L)
        self.choose_combo = ttk.Combobox(window,state='readonly',width= wid_C)
        self.choose_label = Label(window,text="Values",width= wid_La)
        self.choose_combo['values']=self.all_anlys_cols
        self.pivot_by_combobox.bind("<<ComboboxSelected>>",self.choose_pivot)
        self.choose_list_label = Label(window,text="chosen values",width= wid_La)
        self.hs_list_label = Label(window,text="Chosen HS Codes",width= wid_La)
        self.Choose_value = Listbox(window, width= wid_L)
        self.choose_combo.bind("<<ComboboxSelected>>",self.choose)
        self.Choose_value.bind('<Double-1>',self.drop_value)
        self.scroll_hs = Scrollbar(window)
        self.Hs_listbox.config(yscrollcommand=self.scroll_hs.set)
        self.Hs_listbox.bind('<Double-1>',self.drop)
        self.hs_filter = StringVar()
        self.Hs_filter = Entry(window, textvariable = self.hs_filter, width= wid_E)
        self.Hs_filter_label = Label(window, text = "Filter HS starting with:",width= wid_La)
        self.Hs_filter_get = Button(window, text = "Get HS Code", command = self.get_hs,width= wid_b)
        self.scroll_hs.config(command = self.Hs_listbox.yview)
        self.scroll_value = Scrollbar(window)
        self.Choose_value.config(yscrollcommand=self.scroll_value.set)
        self.scroll_value.config(command = self.Choose_value.yview)
        #grids
        space = 4
        self.scroll_hs.grid(row=4,column=3,sticky="nse")
        Hs_label.grid(row=2,column=1,sticky = "w", padx = space, pady = space)
        self.Hs_filter.grid(row=3,column=2,sticky = "w", padx = space, pady = space)
        self.Hs_filter_label.grid(row=3,column=1,sticky = "w", padx = space, pady = space)
        self.Hs_filter_get.grid(row=3,column=3, sticky = "w", padx = space, pady = space)
        self.hs_combobox.grid(row=2,column=2,sticky = "w", padx = space, pady = space)
        self.Hs_listbox.grid(row=4,column=1, columnspan=3, sticky = "ew", padx = space, pady = space)
        #self.hs_list_label.grid(row=1,column=3)
        #self.hs_desc_label.grid(row = 3 , column = 1, columnspan = 4)
        self.hs_extract_button.grid(row=2,column=3,sticky = "w", padx = space, pady = space)
        self.period_label.grid(row=5,column=1,sticky = "w", padx = space, pady = space)
        self.period_combobox.grid(row=5,column=2,sticky = "w", padx = space, pady = space)
        self.gc_ec_combobox.grid(row=6,column=2,sticky = "w", padx = space, pady = space)
        self.gc_ec_label.grid(row=6,column=1,sticky = "w", padx = space, pady = space)
        self.pivot_by_combobox.grid(row=7,column=2,sticky = "w", padx = space, pady = space)
        self.pivot_label.grid(row=7,column=1,sticky = "w", padx = space, pady = space)
        self.choose_combo.grid(row=8,column=2,sticky = "w", padx = space, pady = space)
        self.choose_label.grid(row=8,column=1,sticky = "w", padx = space, pady = space)
        self.Choose_value.grid(row=9,column=2,sticky = "w", padx = space, pady = space)
        self.scroll_value.grid(row=9,column=2,sticky="nse", padx = space, pady = space)
        #self.choose_list_label.grid(row=6,column=3)
        self.analyze_button.grid(row=8,column = 3, sticky = "w", padx = space, pady = space)
        #temporary variables
        self.val = ''
    def drop(self,event):
        cs = self.Hs_listbox.curselection()
        self.hs_items.pop(cs[0])
        self.hs_items_desc.pop(cs[0])
        self.Hs_listbox.delete(cs[0])
        print(self.Hs_listbox.get(cs))
    def get_hs(self):
        print(self.hs_filter.get())
        for j in range(len(self.hs_codes)):
            if str(self.hs_codes[j]).startswith(self.hs_filter.get()) and self.hs_filter.get()!="" and (str(self.hs_codes[j]) not in self.hs_items):
                print("BINGOOOOOO")
                self.hs_items.append(str(self.hs_codes[j]))
                self.hs_descr = self.df_main.loc[self.df_main["HS Code"]==self.hs_codes[j],"HS Description"].unique()
                self.hs_items_desc.append(str(self.hs_descr[0]))
                self.Hs_listbox.insert(len(self.hs_items),self.hs_items[-1]+ " : " + self.hs_items_desc[-1])
    def drop_value(self,event):
        cs = self.Choose_value.curselection()
        self.anlys_cols.pop(cs[0])
        self.Choose_value.delete(cs[0])
        print(self.Choose_value.get(cs))
    def load_files(self):
        w = 100 # width for the Tk root
        h = 50 # height for the Tk root
        ws = self.window.winfo_x() # width of the screen
        hs = self.window.winfo_y()
        wr= self.window.winfo_width()
        hr = self.window.winfo_height()
        x = (ws) + (wr/2) - w/2
        y = (hs) + (hr/2) - h/2 
        self.window2.geometry('%dx%d+%d+%d' % (w, h, x, y))
        file_path = tuple()
        while True:
            try:
                file_path = askopenfilename(multiple = True)
                df_col_list = list()
                amended_columns_list = list()
                key = ["year", "month", "day", "code", "description", "origin", "consignment", "gross", "net", "value (etb)", "value (usd)"]
                amended_columns_list = [self.columns]
                amended_df_list = list()
                self.window2.deiconify()
                self.progress.pack()
                self.progress["value"] = 0
                for j in range(10):
                    print(self.progress["value"])
                    self.progress["value"] = self.progress["value"] + 5
                    self.window2.update()
                    time.sleep(0.1)
                for i in range(len(file_path)):
                    temp_dataframe = pd.read_csv(file_path[i],index_col=0)
                    amended_df_list.append(pd.DataFrame(columns = self.columns))
                    df_col_list.append(temp_dataframe.columns)
                    self.df_list.append(temp_dataframe)
                for i in range(len(df_col_list)):
                    for j in range(len(key)):
                        tester = self.df_list[i].columns
                        tester = [test.strip().lower() for test in tester]
                        if any([key[j] in test for test in tester]):
                            for k in range(len(tester)):
                                if key[j] in tester[k]:
                                    amended_df_list[i][self.columns[j]]= self.df_list[i][self.df_list[i].columns[k]]
                    amended_df_list[i].dropna(axis = "columns", how = "all", inplace = True)
                    amended_columns_list.append(amended_df_list[i].columns)
                    if len(amended_columns_list[i+1])>len(amended_columns_list[i]):
                        for b in range(len(amended_columns_list[i+1])):
                            if amended_columns_list[i+1][b] not in amended_columns_list[i]:
                                amended_df_list[i].drop(columns = amended_columns_list[i+1][b] , inplace = True)
                self.columns = amended_df_list[0].columns
                self.df_main =pd.DataFrame(columns = self.columns)
                self.progress['value'] = 60
                for i in range(len(amended_df_list)):
                    self.df_main = pd.concat([self.df_main,amended_df_list[i]])
                self.df_main.index=range(len(self.df_main[self.df_main.columns[0]]))
                self.df_main = self.df_main.astype({"HS Code":"string"})
                self.df_main["HS Code"] = self.df_main["HS Code"].apply(lambda x: x[:x.find(".")] if x.find(".")!=-1 else x).astype({"HS Code":"string"})
                self.df_main["HS Description"] = self.df_main["HS Description"].apply(lambda x: "MISSING" if x in ["","#NAME?", "#NULL!", "#NUM!", "#REF!","#VALUE!", "#DIV/0", "#N/A", "nan", "NAN"] else x.lstrip("-_= "))
                self.df_main["HS Description"] = self.df_main["HS Description"].apply(lambda x: x.upper())
                print(self.df_main["Year"])           
                self.gc_years = self.df_main[self.columns[0]].unique()
                self.progress['value'] = 100
                self.window2.withdraw()
                messagebox.showinfo("Information", "Completed Loading " + str(len(file_path)) + " file, please extract HS Code" if len(file_path)==1 else "Completed Loading " + str(len(file_path)) + " files, Please extract HS Code")
                break
            except:
                self.window2.withdraw()
                messagebox.showinfo("Information","Please select a .csv file")
        
            pass
    def clean(self):
        pass
    def fil_item(self,event):
        if len(event.keysym) !=1:
            self.val = self.val[:len(self.val)-1]
        else:
            self.val = self.val + event.keysym
        print(self.val)
        list_items = [items for items in self.hs_codes]
        self.hs_combobox.config(value=list(filter(lambda x: self.val in x, list_items )))
        if (self.hs_combobox.get() in self.hs_codes) and (self.hs_combobox.get() not in self.hs_items):
            self.hs_items.append(self.hs_combobox.get())
            self.hs_descr =self.df_main.loc[self.df_main["HS Code"]==self.hs_combobox.get(),"HS Description"].unique()
            self.hs_items_desc.append(self.hs_descr[0])
            self.Hs_listbox.insert(len(self.hs_items),self.hs_items[-1]+ " : "+self.hs_items_desc[-1])
        print(self.hs_items)
        print(self.hs_items_desc)
    def extract_hs_code(self):
        self.hs_codes = list(self.df_main.loc[:,"HS Code"].unique())
        print("HS codes: ",len(self.hs_codes))
        self.hs_combobox.config(value = self.hs_codes)
        messagebox.showinfo("Information", "Extracted HS Codes")
        pass
    def change_trend_scale(self,event):
        pass
    def change_calander_type(self,event):
        print(self.gc_ec_combobox.get())
        if self.gc_ec_combobox.get() == "G.C.":
            pass
        else:
            if "Day" in self.df_main.columns and len(self.df_main_ec)==0:
                self.window2.deiconify()
                self.progress.pack()
                self.progress["value"] = 0
                self.window2.update()
                start = time.time()
                calander_store = pd.DataFrame()
                GC_Year = self.df_main["Year"]
                GC_Month = self.df_main["Month"]
                GC_Day = self.df_main["Day"]
                EC_year = [0 for i in range(len(GC_Year))]
                EC_month = [0 for i in range(len(GC_Year))]
                EC_day = [0 for i in range(len(GC_Year))]
                def is_GC_leap(year):
                    if year%4==0:
                        if year%100==0 and year%400!=0:
                            return False
                        else:
                            return True
                    else:
                        return False
                def is_EC_leap(year):
                    if (year+1)%4==0:
                        return True
                    else:
                        return False
                self.progress["value"] = 10
                self.window2.update()
                end_1 = time.time()
                calander_store["is_GC_leap"] = self.df_main["Year"].apply(is_GC_leap)
                temp_df = self.df_main["Year"].apply(lambda x: x+1)
                calander_store["is_GC_leap+1"] = temp_df.apply(is_GC_leap)
                calander_store["is_EC_leap"] = self.df_main["Year"].apply(is_EC_leap)
                calander_store["end_year"] = calander_store["is_EC_leap"].apply(lambda x: 366 if x else 365)
                calander_store["sep_day"] = calander_store["is_EC_leap"].apply(lambda x: 12 if x else 11)
                calander_store["pag_day"] = calander_store["is_EC_leap"].apply(lambda x: 6 if x else 5)
                calander_store["feb_day"] = calander_store["is_GC_leap+1"].apply(lambda x: 29 if x else 28)
                calander_store["EC_year"] = GC_Year
                calander_store["EC_Month"] = GC_Month
                calander_store["dist_mon_sep"] = GC_Month
                calander_store.loc[calander_store["EC_Month"]<9,"EC_year"] = calander_store.loc[calander_store["EC_Month"]<9,"EC_year"].apply(lambda x: x-8)
                calander_store.loc[calander_store["EC_Month"]<9,"dist_mon_sep"] = calander_store.loc[calander_store["EC_Month"]<9,"dist_mon_sep"].apply(lambda x: x+3)
                calander_store.loc[calander_store["EC_Month"]>9,"EC_year"] = calander_store.loc[calander_store["EC_Month"]>9,"EC_year"].apply(lambda x: x-7)
                calander_store.loc[calander_store["EC_Month"]>9,"dist_mon_sep"] = calander_store.loc[calander_store["EC_Month"]>9,"dist_mon_sep"].apply(lambda x: x-9)
                calander_store["sep_less_gc_day"] = calander_store["sep_day"]<GC_Day
                calander_store["sep_great_gc_day"] = calander_store["sep_day"]>=GC_Day
                calander_store.loc[(calander_store["EC_Month"]==9) & (calander_store["sep_less_gc_day"]==True),"EC_year"] = calander_store.loc[(calander_store["EC_Month"]==9) & (calander_store["sep_less_gc_day"]==True),"EC_year"].apply(lambda x: x-8)
                calander_store.loc[calander_store["EC_Month"]==9,"dist_mon_sep"] = calander_store.loc[calander_store["EC_Month"]==9,"dist_mon_sep"].apply(lambda x: 12)
                calander_store.loc[(calander_store["EC_Month"]==9) & (calander_store["sep_great_gc_day"]==True),"EC_year"] = calander_store.loc[(calander_store["EC_Month"]==9) & (calander_store["sep_great_gc_day"]==True),"EC_year"].apply(lambda x: x-7)
                Month_days_G_b=[30,31,30,31,31]
                self.progress["value"] = 50
                self.window2.update()
                Month_days_G_a = [31,30,31,30,31,31]
                calander_store["GC_day"] = GC_Day
                calander_store["num_days"] = calander_store.apply(lambda x: sum(Month_days_G_b[:x["dist_mon_sep"]])+sum(Month_days_G_a[:0 if x["dist_mon_sep"]-6<=0 else x["dist_mon_sep"]-6])+x["feb_day"]*(0 if x["dist_mon_sep"]<=5 else 1) - x["sep_day"]+x["GC_day"], axis = 1)
                calander_store["EC_day"] = calander_store.apply(lambda x: x["num_days"]-x["end_year"]+1 if x["num_days"]>=x["end_year"] else x["num_days"]%30+1, axis = 1)
                calander_store["EC_month"] = calander_store.apply(lambda x: 1 if x["num_days"]>=x["end_year"] else math.floor(x["num_days"]/30)+1, axis = 1)
                end_2 = time.time()
                test_date = 12548
                print("end_1: ", end_1-start)
                print("end_2: ", end_2-start)
                print("Year_EC: ",calander_store.loc[test_date,"EC_year"], "Month_EC: ",calander_store.loc[test_date,"EC_month"], "Day_EC: ",calander_store.loc[test_date,"EC_day"])
                print("Year_GC: ",self.df_main.loc[test_date,"Year"], "Month_GC: ",self.df_main.loc[test_date,"Month"], "Day_GC: ",self.df_main.loc[test_date,"Day"])
                print("num_days: ",calander_store.loc[test_date,"num_days"],"end_year: ",calander_store.loc[test_date,"end_year"] )
                print("dist_mon_sep: ",calander_store.loc[test_date,"dist_mon_sep"],"feb_day: ",calander_store.loc[test_date,"feb_day"] )
                self.df_main_ec = self.df_main
                self.df_main_ec["Year"] = calander_store["EC_year"]
                self.df_main_ec["Month"] = calander_store["EC_month"]
                self.df_main_ec["Day"] = calander_store["EC_day"]
                self.window2.update()
                self.progress["value"] = 100
                self.window2.withdraw()
                messagebox.showinfo("Information", "Date Converted to Ethiopian Calander system")
            elif "Day" in self.df_main.columns and len(self.df_main_ec)!=0:
                pass
            else:
                messagebox.showinfo("Information","You can not use this mode without DAY value being absent from the database")
        pass
    def pivot_by(self,event):
        pass
    def choose(self,event):
        self.anlys_cols.append(self.choose_combo.get())
        print(self.anlys_cols)
        self.Choose_value.insert(len(self.anlys_cols),self.anlys_cols[-1])
        pass
    def choose_pivot(self,event):
        pass
    def manage(self):
        self.window2.deiconify()
        self.progress["value"] = 0
        for j in range(10):
            self.progress["value"] = self.progress["value"] + 2.5
            self.window2.update()
            time.sleep(0.1)
        #anlys_cols = ["Gross Weight (Kg)", "Net Weight (Kg)", "CIF/FOB Value (ETB)", "CIF/FOB Value (USD)"]
        col_sheet_names = [item.replace("(","_") for item in self.anlys_cols]
        col_sheet_names = [item.replace(")","_") for item in self.anlys_cols]
        col_sheet_names = [item.replace("/","_") for item in self.anlys_cols]
        sheet_names = col_sheet_names + ["H"+ item for item in col_sheet_names]
        print(self.hs_combobox.get())
        print(self.period_combobox.get())
        print(self.pivot_by_combobox.get())
        print(self.gc_ec_combobox.get())
        if self.gc_ec_combobox.get() == "G.C." and self.period_combobox.get() in ["Year","Month"] and len(self.hs_items)!=0 and self.pivot_by_combobox.get() in ["Country (Origin)", "Country (Consignment)"] and self.choose_combo.get() in self.all_anlys_cols:
            pivot_i = list()
            print(self.df_main.loc[5,"HS Code"])
            piv = self.df_main.loc[(self.df_main["HS Code"].isin(self.hs_items)),:]
            print("piv: ", piv.info())
            print("HHHHHHHHHHHHHHH: ", self.df_main.info())
            self.progress["value"] = 40
            self.window2.update()
            for i in range(len(self.anlys_cols)):
                pivot_i.append(pd.pivot_table(self.df_main.loc[(self.df_main["HS Code"].isin(self.hs_items)),:],index=self.pivot_by_combobox.get(), columns=self.period_combobox.get(), values=self.anlys_cols[i],aggfunc="sum"))
                print(pivot_i[i].info())
            self.progress["value"] = 60
            self.window2.update()
            for i in range(len(self.anlys_cols)):
                pivot_i.append(pd.pivot_table(self.df_main.loc[(self.df_main["HS Code"].isin(self.hs_items)),:],index="HS Code", columns=self.period_combobox.get(), values=self.anlys_cols[i],aggfunc="sum"))
                print("HS Description: ", self.hs_items_desc)
                print(pivot_i[-1].info())
                pivot_i[-1].loc[:,"HS Description"] = self.hs_items_desc
                pivot_i[-1]=pivot_i[-1][["HS Description"] + list(pivot_i[-1].columns[:len(pivot_i[-1].columns)-1])]
            self.progress["value"] = 100
            self.window2.update()
            self.window2.withdraw()
            save_file = asksaveasfilename(defaultextension=".xlsx")
            with pd.ExcelWriter(save_file) as writer:
                for i in range (len(sheet_names)):
                    pivot_i[i].to_excel(writer, sheet_name=sheet_names[i])
            messagebox.showinfo("Information", "Data Extraction successful!")
        elif self.gc_ec_combobox.get() == "E.C." and self.period_combobox.get() in ["Year","Month"] and len(self.hs_items)!=0 and self.pivot_by_combobox.get() in ["Country (Origin)", "Country (Consignment)"] and self.choose_combo.get() in self.all_anlys_cols:
            pivot_i = list()
            print(self.df_main_ec.loc[5,"HS Code"])
            piv = self.df_main_ec.loc[(self.df_main_ec["HS Code"].isin(self.hs_items)),:]
            print("piv: ", piv.info())
            print("HHHHHHHHHHHHHHH: ", self.df_main_ec.info())
            for i in range(len(self.anlys_cols)):
                pivot_i.append(pd.pivot_table(self.df_main_ec.loc[(self.df_main_ec["HS Code"].isin(self.hs_items)),:],index=self.pivot_by_combobox.get(), columns=self.period_combobox.get(), values=self.anlys_cols[i],aggfunc="sum"))
                print(pivot_i[i].info())
            for i in range(len(self.anlys_cols)):
                pivot_i.append(pd.pivot_table(self.df_main_ec.loc[(self.df_main_ec["HS Code"].isin(self.hs_items)),:],index="HS Code", columns=self.period_combobox.get(), values=self.anlys_cols[i],aggfunc="sum"))
                print("HS Description: ", self.hs_items_desc)
                print(pivot_i[-1].info())
                pivot_i[-1].loc[:,"HS Description"] = self.hs_items_desc
                pivot_i[-1]=pivot_i[-1][["HS Description"] + list(pivot_i[-1].columns[:len(pivot_i[-1].columns)-1])]
            save_file = asksaveasfilename(defaultextension=".xlsx")
            with pd.ExcelWriter(save_file) as writer:
                for i in range (len(sheet_names)):
                    pivot_i[i].to_excel(writer, sheet_name=sheet_names[i])
            messagebox.showinfo("Information", "Data Extraction successful!")
            
        else:
            self.window2.withdraw()
            messagebox.showinfo("Information","Please fill in all values")
                
        pass
    def cleaner(self, dataframe):
        df = dataframe
        Header = df.columns
        df_new = pd.DataFrame(columns = Header)
        for i in range(len(dataframe.columns)):
            if self.d_type[i]=="int":
                out = df.loc[:,Header[i]].astype(str)
                out = out.apply(lambda x: x.strip(" "))
                out = out.str.replace(",","")
                out = out.str.replace(" ","")
                out = out.str.replace("-","0")
                out = pd.to_numeric(out,downcast = "integer", errors = "coerce")
                df_new.loc[:,Header[i]] = out.astype(int, errors = "ignore")
                df_new = df_new.fillna({Header[i]:0})
            elif self.d_type[i]=="float":
                out = df.loc[:,Header[i]].astype(str)
                out = out.apply(lambda x: x.strip(" "))
                out = out.str.replace(",","")
                out = out.str.replace(" ","")
                out = out.str.replace("-","0")
                out = pd.to_numeric(out,downcast = "float", errors = "coerce")
                df_new.loc[:,Header[i]] = out
                df_new = df_new.fillna({Header[i]:0})
            else:
                out = df.loc[:,Header[i]].astype(str)
                out = out.apply(lambda x: "MISSING" if x in ["#NAME?", "#NULL!", "#NUM!", "#REF!","#VALUE!", "#DIV/0", "#N/A", "nan", "NAN"] else x.lstrip("-_= "))
                out = out.apply(lambda x: x[:x.find(".")] if x.find(".")!=1 else x)
                out = out.apply(lambda x: x.upper())
                df_new.loc[:,Header[i]] = out.astype("string")
                df_new = df_new.fillna({Header[i]:"Missing"})
        df_new.drop_duplicates(keep = "first", inplace = True)
        df_new.drop(df_new.columns[df.columns.str.contains("Unnamed", case = False)] , axis = 1, inplace = True)
        df_new = df_new.loc[(df_new[Header[0]]!=0) | (df_new[Header[0]]!=0) != "" ,:]
        return df_new

ERCA_tool(windows, window1)

windows.mainloop()
window1.mainloop()

