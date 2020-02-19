
from tkinter import*
from tkinter import messagebox
from openpyxl import load_workbook
import xlrd
import pandas as pd
import tkinter as tk
from pandas import DataFrame
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg




root=Tk()                               #Main window 
f=Frame(root)
frame1=Frame(root)
frame2=Frame(root)
frame3=Frame(root)
root.title("Simple Student Record System")
root.geometry("720x230")
root.configure(background="yellow")


scrollbar=Scrollbar(root)
scrollbar.pack(side=RIGHT, fill=Y)

firstname=StringVar()                    #Declaration of all variables
lastname=StringVar()
id=StringVar()
dept=StringVar()
designation=StringVar()
remove_firstname=StringVar()
remove_lastname=StringVar()
searchfirstname=StringVar()
searchlastname=StringVar()
sheet_data=[]
row_data=[]

def emp_dict(*args):                   #To add a new entry and check if entry already exist in excel sheet
    #print("done")
    workbook_name="sample.xlsx"
    workbook=xlrd.open_workbook(workbook_name)
    worksheet=workbook.sheet_by_index(0)
    
    wb=load_workbook(workbook_name)
    page=wb.active
    
    p=0
    for i in range(worksheet.nrows):
        for j in range(worksheet.ncols):
            cellvalue=worksheet.cell_value(i,j)
            print(cellvalue)   
            sheet_data.append([])
            sheet_data[p]=cellvalue
            p+=1
    print(sheet_data)
    fl=firstname.get()
    fsl=fl.lower()
    ll=lastname.get()
    lsl=ll.lower()
    if (fsl and lsl) in sheet_data:
        print("found")
        messagebox.showerror("Error","This student  exist")
    else:
        print("not found")
        for info in args:
            page.append(info)
        messagebox.showinfo("Done","Successfully added the student record")

    wb.save(filename=workbook_name)
    
def add_entries():                       #to append all data and add entries on click the button
    a=" "
    f=firstname.get()
    f1=f.lower()
    l=lastname.get()
    l1=l.lower()
    d=dept.get()
    d1=d.lower()
    de=designation.get()
    de1=de.lower()
    list1=list(a)
    list1.append(f1)
    list1.append(l1)
    list1.append(d1)
    list1.append(de1)
    emp_dict(list1)


def add_info():                                           #for taking user input to add the enteries
    frame2.pack_forget()
    frame3.pack_forget()
    
    emp_first_name=Label(frame1,text="Enter first name of the student: ",bg="red",fg="white")
    emp_first_name.grid(row=1,column=1,padx=10)
    e1=Entry(frame1,textvariable=firstname)
    e1.grid(row=1,column=2,padx=10)
    e1.focus()
    
    emp_last_name=Label(frame1,text="Enter last name of the student: ",bg="red",fg="white")
    emp_last_name.grid(row=2,column=1,padx=10)
    e2=Entry(frame1,textvariable=lastname)
    e2.grid(row=2,column=2,padx=10)
    
    emp_dept=Label(frame1,text="Select department of student: ",bg="red",fg="white")
    emp_dept.grid(row=3,column=1,padx=10)
    dept.set("Select Option")
    e4=OptionMenu(frame1,dept,"Select Option","IT","International managment","Aviational managment")
    e4.grid(row=3,column=2,padx=10)
    
    emp_desig=Label(frame1,text="Select nationality of student: ",bg="red",fg="white")
    emp_desig.grid(row=4,column=1,padx=10)
    designation.set("Select Option")
    e5=OptionMenu(frame1,designation,"Select Option","Kyrgyz","Kazakh","Russian","Polish","Ukrainan", 
                  "USA","Indian","Belarus","German")
    e5.grid(row=4,column=2,padx=10)
    
    button4=Button(frame1,text="Add to list",command=add_entries)
    button4.grid(row=5,column=2,pady=10)
    
    frame1.configure(background="Red")
    frame1.pack(pady=10)
    
def clear_all():             #for clearing the entry widgets
    frame1.pack_forget()
    frame2.pack_forget()
    frame3.pack_forget()

    
def remove_emp():                #for taking user input to remove enteries
    clear_all()
    emp_first_name=Label(frame2,text="Enter first name of the student",bg="red",fg="white")
    emp_first_name.grid(row=1,column=1,padx=10)
    e6=Entry(frame2,textvariable=remove_firstname)
    e6.grid(row=1,column=2,padx=10)
    e6.focus()
    emp_last_name=Label(frame2,text="Enter last name of the student",bg="red",fg="white")
    emp_last_name.grid(row=2,column=1,padx=10)
    e7=Entry(frame2,textvariable=remove_lastname)
    e7.grid(row=2,column=2,padx=10)
    remove_button=Button(frame2,text="Click to remove",command=remove_entry)
    remove_button.grid(row=3,column=2,pady=10)
    frame2.configure(background="Red")
    frame2.pack(pady=10)

def remove_entry():  #to remove entry from excel sheet
    rsf=remove_firstname.get()
    rsf1=rsf.lower()
    print(rsf1)
    rsl=remove_lastname.get()
    rsl1=rsl.lower()
    print(rsl1)
    workbook_name="sample.xlsx"
    path="sample.xlsx"
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(0)

    for row_num in range(sheet.nrows):
        row_value = sheet.row_values(row_num)
        if (row_value[1]==rsf1 and row_value[2]==rsl1):
            print(row_value)
            print("found")
            file="sample.xlsx"
            x=pd.ExcelFile(file)
            dfs=x.parse(x.sheet_names[0])
            dfs=dfs[dfs['First Name']!=rsf]
            dfs.to_excel("sample.xlsx",sheet_name='student',index=False)
            messagebox.showinfo("Done","Successfully removed the student record")
    clear_all()

def search_emp():     #can implement search by 1st name,last name,emp id, designation
    clear_all()
    emp_first_name=Label(frame3,text="Enter first name of the student",bg="red",fg="white")   #to take user input to seach
    emp_first_name.grid(row=1,column=1,padx=10)
    e8=Entry(frame3,textvariable=searchfirstname)
    e8.grid(row=1,column=2,padx=10)
    e8.focus()
    emp_last_name=Label(frame3,text="Enter last name of the student",bg="red",fg="white")
    emp_last_name.grid(row=2,column=1,padx=10)
    e9=Entry(frame3,textvariable=searchlastname)
    e9.grid(row=2,column=2,padx=10)
    search_button=Button(frame3,text="Click to search",command=search_entry)
    search_button.grid(row=3,column=2,pady=10)
    
    frame3.configure(background="Red")
    frame3.pack(pady=10)

    
def search_entry():
    sf=searchfirstname.get()
    ssf1=sf.lower()
    print(ssf1)
    sl=searchlastname.get()
    ssl1=sl.lower()
    print(ssl1)
    path="sample.xlsx"
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(0)

    for row_num in range(sheet.nrows):
        row_value = sheet.row_values(row_num)
        if (row_value[1]==ssf1 and row_value[2]==ssl1):
            print(row_value)
            print("found")
            messagebox.showinfo("Done","Searched student Exist")
            clear_all()
    #else:
    if(row_value[1]!=ssf1 and row_value[2]!=ssl1):
        print("Not found")
        messagebox.showerror("Sorry","student record does not Exist")
        clear_all()

        
#Main window buttons and labels
        
label1=Label(root,text="SIMPLE STUDENT RECORD SYSTEM")
label1.config(font=('Italic',16,'bold'), justify=CENTER, background="Blue",fg="Yellow", anchor="center")
label1.pack(fill=X)

label2=Label(f,text="Select an action: ",font=('bold',12), background="Black", fg="White")
label2.pack(side=LEFT,pady=10)
button1=Button(f,text="Add", background="brown", fg="Black", command=add_info, width=12)
button1.pack(side=LEFT,ipadx=20,pady=10)
button2=Button(f,text="Remove", background="Brown", fg="Black", command=remove_emp, width=12)
button2.pack(side=LEFT,ipadx=20,pady=10)
button3=Button(f,text="Search", background="Brown", fg="Black", command=search_emp, width=12)
button3.pack(side=LEFT,ipadx=20,pady=10)
button6=Button(f,text="Close", background="Brown", fg="Black", width=12, command=root.destroy)
button6.pack(side=LEFT,ipadx=20,pady=10)
f.configure(background="Black")
f.pack()



data1 = {'Country': ['PL','KG','KZ','UKR','CH'],
         'num_of_st': [5234,345,1432,2587,56]
        }
df1 = DataFrame(data1,columns=['Country','num_of_st'])


data2 = {'Year': [2020,2019,2018,2017,2016,2015,2014,2013,2012,2011],
         'appl_st': [221,543,134,322,143,432,97,54,34,32]
        }
df2 = DataFrame(data2,columns=['Year','appl_st'])


root= tk.Tk() 
  
figure1 = plt.Figure(figsize=(6,5), dpi=100)
ax1 = figure1.add_subplot(111)
bar1 = FigureCanvasTkAgg(figure1, root)
bar1.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH)
df1 = df1[['Country','num_of_st']].groupby('Country').sum()
df1.plot(kind='bar', legend=True, ax=ax1)
ax1.set_title('Country Vs num of st')

figure2 = plt.Figure(figsize=(5,4), dpi=100)
ax2 = figure2.add_subplot(111)
line2 = FigureCanvasTkAgg(figure2, root)
line2.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH)
df2 = df2[['Year','appl_st']].groupby('Year').sum()
df2.plot(kind='line', legend=True, ax=ax2, color='r',marker='o', fontsize=10)
ax2.set_title('Year Vs. appl st')



root.mainloop()
