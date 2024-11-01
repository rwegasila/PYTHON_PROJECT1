
import  tkinter
from tkinter import*
from tkinter import ttk
from datetime import date
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageTk
import os
import openpyxl
from openpyxl import workbook
import pathlib


#Ritrieving Data from the widgets
def enter_data():
      cond_check=check_var.get()
      if cond_check == "Accepted" :
       
        Firstname =first_name_frame.get()
        Lastname = last_name_frame.get()
        if Firstname and Lastname:
           
          Title= title_combobox.get()
          Age=age_spinbox.get()
          Nationality = nationality_combobox.get()
          Complited_course =numcourse_spinbox.get()
          Number_of_semister=numsemister_spinbox.get()
          registration_checkbutton = reg_status_var.get()
          Birthdate=date_lable_Entry.get()
          gender =   gender_check.get()
          course_lable = course_lable_combobox.get()
          department = department_combobox.get()
          year_study = year_study_combobox.get()
  

          print("FirstName:",Firstname, "LastName:", Lastname, "Title:",Title,)
       
          print("Age:",Age, "Nationality:", Nationality)
          print("Birthdate", Birthdate)
          print("Gender", gender)
          print("course Name",course_lable )
          print("department", department)
          print("Year of Study", year_study)
          print("............................................")
          print( "Complited_course:", Complited_course, "Number_of_semister:", Number_of_semister)
          print("Registration_Status", registration_checkbutton)

        #Insert row into Excell Sheet
       
          path = "C:\\Users\\DEVIS\\Desktop\\REGISTRATION FORM DATA.xlsx"
          if not os.path.exists(path):
              workbook = openpyxl.Workbook()
              sheet = workbook.active
              heading=["Firstname", "Lastname", "Title", "Age", "Nationality", "Birthdate","gender", "course_lable", "department", "year_study" "Complited_course", "Number_of_semister", "registration_checkbutton"]
              sheet.append(heading)
              workbook.save(path)

          workbook = openpyxl.load_workbook(path)
          sheet = workbook.active
          heading=[Firstname, Lastname, Title, Age, Nationality, Complited_course, Number_of_semister, registration_checkbutton]
          workbook.save(path)

        else : 
          tkinter.messagebox.showwarning(title="Error", message="Firstname and Lastname are required")
      else:   
          tkinter.messagebox.showwarning(title="Error", message="you not accept the terms and condition")

#Create Root window it is also called Parent window
window = tkinter.Tk()

#Creating window Title
window.title("Data Entry Form")

#Creating window heading
Label(window, text="Email: devisrwegasila6@gmail.com", width=10, height=3, bg="orange", anchor='e').pack(side=TOP, fill=X)
Label(window, text="STUDENT REGISTRATION FORM", width=10, height=2, bg="#c36464",fg="brown" , font='aerialbold').pack(side=TOP, fill=X)

#Searching  box
search= StringVar()
Entry(window, textvariable=search, width=20,bd=2,  ).place(x=920,y=60)
search = Button(window, text="search", compound=LEFT,  width=12,  font="aerial 13 bold")
search.place(x=1060, y=58)

update_button= Button(window, text="update", )
update_button.place(x=122, y=60)


#Creating Frame Inside the Window
frame = tkinter.Frame(window)

#To show the created Frame Inside the window
frame.pack()

#Creating the Lable Frame inside the Frame and saving the user information
user_info_frame = tkinter.LabelFrame(frame, text="Student information")

#Position of that frame
user_info_frame.grid(row=0, column=0, padx=40, pady=40)

#Creating the small lable frame inside the user info fram frame
first_name_frame = tkinter.Label(user_info_frame, text="First Name")
last_name_frame = tkinter.Label(user_info_frame, text="Last Name")

#Position of that small Lable inside the user info frame Frame 
first_name_frame.grid(row=0, column=0)
last_name_frame.grid(row=0, column=1)


#To create Entry for the first and Last name
first_name_frame = tkinter.Entry(user_info_frame)
last_name_frame = tkinter.Entry(user_info_frame)

#Grid in order to place them on the screen
first_name_frame.grid(row=1, column=0)
last_name_frame.grid(row=1, column=1)

#Creating the user title and title combobox that help user for selection ,you can get the combobox by importing another module called ttk
title_lable =tkinter.Label(user_info_frame, text="Title")
title_combobox = ttk.Combobox(user_info_frame, values=["", "Mr.", "Ms."])
title_lable.grid(row=0,column=2)
title_combobox.grid(row=1, column=2)

#Creating the spinbox for Age input
age_lable =tkinter.Label(user_info_frame, text="Age")
age_spinbox = tkinter.Spinbox(user_info_frame, from_=18, to=25)
age_lable.grid(row=2, column=0)
age_spinbox.grid(row=3, column=0)

#Nationality
nationality_lable = tkinter.Label(user_info_frame, text="Nationality")
nationality_combobox = ttk.Combobox(user_info_frame, values=["KENYA","RWANDA", "BURUNDI", "TANZANIA", "AUSTRALIA", "NIGERIA", "UGANDA", "CALPHONIA"])
nationality_lable.grid(row=2, column=1)
nationality_combobox.grid(row=3,  column=1)

#Registration number
RegNo_lable = tkinter.Label(user_info_frame, text="Registration No.")
RegNo_lable_spinbox = tkinter.Spinbox(user_info_frame,from_=1, to="infinity")
RegNo_lable.grid(row=2,column=2)
RegNo_lable_spinbox.grid(row=3,column=2)

#Date setting
date_lable= tkinter.Label(user_info_frame, text="Birth Date")
date_lable_Entry= tkinter.Entry(user_info_frame)
date_lable_Entry.insert(0,"DD/MM/YY")
date_lable_Entry.bind("<FocusIn>",lambda e: date_lable_Entry.delete('0', 'end'))
date_lable.grid(row=4, column=0)
date_lable_Entry.grid(row=5, column=0)

#Course name
course_lable = tkinter.Label(user_info_frame, text="Course Name")
course_lable_combobox = ttk.Combobox(user_info_frame, values=["Data science Engineering", "Computer science Engineering", "Electrical Engineering", "Electronic Engineering", "Software Engineering", "Labscience Engineering", "Physics and Mathematics Education","Biomedical Engineering"])
course_lable.grid(row=4, column=2)
course_lable_combobox.grid(row=5, column=2)

#Department Name 
department = tkinter.Label(user_info_frame, text="Department")
department_combobox = ttk.Combobox(user_info_frame, values=["CoICT", "CoSTE", "CoACT", "CoAST"])
department.grid(row=6, column=0)
department_combobox.grid(row=7, column=0)

#Gender 
gender= tkinter.Label(user_info_frame, text="Gender")
gender.grid(row=4, column=1)
gender_check= ttk.Combobox(user_info_frame,values=["Male", "Female"]  )
gender_check.grid(row=5, column=1)

#Year of study
year_study = tkinter.Label(user_info_frame, text="Year of study")
year_study_combobox = ttk.Combobox(user_info_frame, values=["First Year", "Second Year", "Third Year", "Fourth Year"])
year_study.grid(row=6, column=1)
year_study_combobox.grid(row=7, column=1)



#Creating space in your widget
for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

#Saving course information by creating the new frame inside the window
corse_info_frame = tkinter.LabelFrame(frame, text="Regstration status")
corse_info_frame.grid(row=1, column=0, sticky="news", padx=40, pady=40 )

#Used to check if the student was registed or not
reg_status_var = tkinter.StringVar(value="Not Registed")
registration_checkbutton = tkinter.Checkbutton(corse_info_frame, text="current Registered", variable=reg_status_var, 
                                               onvalue="Registered", offvalue="Not Registered")


registration_checkbutton.grid(row=1, column=0)

#Number of complited course
numcourse_lable = tkinter.Label(corse_info_frame, text="#Complited Course")
numcourse_spinbox = tkinter.Spinbox(corse_info_frame, from_=0, to="infinity")
numcourse_lable.grid(row=0, column=1)
numcourse_spinbox.grid(row=1, column=1)

#Number of semisters
numsemister_lable = tkinter.Label(corse_info_frame, text="Number of semister")
numsemister_spinbox = tkinter.Spinbox(corse_info_frame, from_=1, to="infinity")
numsemister_lable.grid(row=0,column=2)
numsemister_spinbox.grid(row=1, column=2)

#Creating space in your widget
for widget in corse_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

#Creating the New frame Terms and condition
terms_info_frame = tkinter.LabelFrame(frame, text="Terms & Condition")
terms_info_frame.grid(row=2, column=0, sticky="news", padx=40, pady=40)

check_var = tkinter.StringVar(value="Not Registered")
cond_check = tkinter.Checkbutton(terms_info_frame,text="I agree this terms and Condition" , variable=check_var,
                                 onvalue="Accepted", offvalue="Not Registered" )
cond_check.grid(row=1, column=0)

#Creating space in your widget
for widget in terms_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

    #Button
button = tkinter.Button(frame, text="Enter Data", command=enter_data)
button.grid(row=3, column=0,sticky="news", padx=40, pady=40)



#Window Roop for excution
window.mainloop()