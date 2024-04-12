#importing Tkinter library
import tkinter as tk
#importing ttk which allows to style / theme our windows or app
from tkinter import ttk
#importing openpyxl which is a library that allows us to work with excel files
import openpyxl

#
##
###
####
#####
#*************** CREATES THE INSERT DATA INFORMATION TO TABLE ********************

def insert_row():
    #this is getting all the information that was inputted by the user into the fields
    name = name_entry.get()
    age = int(age_spinbox.get())
    country = status_combobox.get()
    employment_status = "Employed" if a.get() else "Unemployed"

    print(name,age,country,employment_status)

    #This will insert the new data from the data entry into the active sheet in the excel sheet
    path = r"C:\Users\Lenovo\PycharmProjects\TravelDataEntryProject\travelbudget.xlsx"
    workbook = openpyxl.load_workbook(path)
    # you're getting the ACTIVE sheets that's CURRENTLY OPEN despite having multiple tabs in one excel file
    sheet = workbook.active
    #this will save the data entry into a list below
    row_values = [name, age, country, employment_status]
    #this will append or insert the data into the excel sheet
    #***** Make sure that the excel sheet is CLOSED. It will not work if its open.
    sheet.append(row_values)
    workbook.save(path)

    #insert the new data into the table in the windows view. So that we can see it as well
    treeview.insert('', tk.END, values=row_values)

    #clearing the entry that was inserted before and returning the default value in the field
    name_entry.delete(0, 'end')
    name_entry.insert(0, 'Name')
    age_spinbox.delete(0, 'end')
    age_spinbox.insert(0, 'Age')
    status_combobox.set(combo_list[0])
    checkbutton.state(["!selected"])



#*************** CREATES THE TOGGLE MODE FUNCTION ********************
#this is a FUNCTION that will help user to change from Light to Dark mode
# by clicking the MODE Switch
# 'select' means by the ACTION that user executed by clicking the switch
def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

#*************** CREATES THE LOADING OF DATA FUNCTION ********************

#this function helps us to load the data into the excel sheet specified
def load_data():
    path = r"C:\Users\Lenovo\PycharmProjects\TravelDataEntryProject\travelbudget.xlsx"
    workbook = openpyxl.load_workbook(path)
    # you're getting the ACTIVE sheets that's CURRENTLY OPEN despite having multiple tabs in one excel file
    sheet = workbook.active

    list_values = list(sheet.values)
    #print(list_values) #this will only show what's inside the excel sheet. It will not show up inthe window itself.
    for col_name in list_values[0]: #which is the column names from the excel sheet (sub-zero)
        # meaning putting the column name under the headings. eg. if col name 'Name'
        # is will put it as 'Name' as well on the window
        treeview.heading(col_name, text=col_name)

        # '1' and onwards are the list of user details (sub-one and onwards) this will
        #input the details inside the window
        for value_tuple in list_values[1:]:
            treeview.insert('', tk.END, values=value_tuple)

#root widget/window (OUR PARENT WINDOW)
root = tk.Tk()

#for apply styling or theme in our application
style = ttk.Style(root)

#calling our choice of theme
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")

#the theme that we choose to work with
style.theme_use("forest-dark")

#Combo_list of countries
combo_list = ['Singapore', 'Philippines', 'Indonesia', 'Malaysia',
              'Vietnam', 'Thailand', 'Japan']

#this creates a frame INSIDE the 'root' (PARENT WINDOW)
frame = ttk.Frame(root)
frame.pack() #this helps the frame more responsive and resizable

#this creates a 'label frame' a sub-frame inside a 'frame' window
widgets_frame = ttk.LabelFrame(frame, text="Insert Travel Budget")

#you must put '.pack' OR you can put '.grid' so that it will appear in the window
# .pack will also make the entire frame to be centered in the middle of the
# root window, no matter how you resize the root window
# '.grid' will create geometry manager or creating grids inside your window
# .pack or .grid will always be the last line to wrap it all up
# 'padx' gives 20 spaces on both sides. 'pady' gives 10 spaces on top & bottom
widgets_frame.grid(row=0, column=0, padx=20, pady=10)

#uset input fields and this will be inside 'widgets_frame' inside the root window
name_entry = ttk.Entry(widgets_frame)

#
##
###
####
#####
#*************** LEFT SIDE OF THE WINDOW ********************
#*************************************************************

#***************CREATION OF 'NAME' FIELD*********************

#this will help to put 'Name' inside the field for users to fill in
#However, the user have to manually delete the 'Name' from inside the
#field box. Which is a bad user experience...
name_entry.insert(0, 'Name')

#this will help the word 'Name' inside the field disappear when
#users click inside the field
#'.bind' helps to bind certain events together by using FocusIn a feature
#that will help when user click inside a field you are able to type in it???
#'.delete' helps to delete the 'Name' inside the field from the start '0'
#until the 'end'. And lambda e, helps to call this.
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end'))

#this will create another rows of grid inside the widget_frame with 1 column only
# stick = 'ew' meaning it will expand it east - west to fill up the size of the
# 'label frame'
# if you see 'pady=(0, 5)' this means top has 0 spaces and bottom has 5 spaces
name_entry.grid(row=0, column=0, padx=5, pady=(0, 5), sticky='ew')

#***************CREATION OF 'AGE' FIELD*********************

#this will create an up and down button for age field starting form 18 to 100 yrs old
age_spinbox = ttk.Spinbox(widgets_frame, from_=18, to=100)
age_spinbox.grid(row=1, column=0, padx=0, sticky='ew')
age_spinbox.insert(0, 'Age')
age_spinbox.grid(row=1, column=0, padx=5, pady=5, sticky='ew')

#***************CREATION OF 'COUNTRY' FIELD*********************

#this will create a COMBO BOX field or Drop down menu inside the 'label frame'
#inside the 'values' this will be the list of options to choose from by users
status_combobox = ttk.Combobox(widgets_frame, values=combo_list)

#this will help to put the FIRST '0' position from the list as the default
#in the field. Eg. 'Singapore'
status_combobox.current(0)
status_combobox.grid(row=2, column=0, padx=5, pady=5, sticky='ew')

#***************CREATION OF 'EMPLOYED' FIELD*********************

# 'a' variable will help us later to save the information entered
# (if ticked or not) to be saved inside the excel sheet
a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widgets_frame, text="Employed", variable = a)
checkbutton.grid(row=3, column=0, sticky='nsew')

#***************CREATION OF 'SUBMIT' BUTTON*********************

#The command function will connect the button to the 'insert_row' function whereby, it will insert the data
#into the table
button = ttk.Button(widgets_frame, text="Submit", command=insert_row)

# stick='nsew' helps to centre the button?
button.grid(row=4, column=0, padx=5, pady=5, sticky='nsew')

#***************CREATION A SEPERATOR LINE *********************

seperator = ttk.Separator(widgets_frame)
seperator.grid(row=5, column=0, padx=(20,10), pady=5, sticky='ew')

#***************CREATION DARK & LIGHT MODE SWICH *********************

#This will appear as a CHECKBOX, but we want it to be a SWITCH
# Hence, we add in 'style = "Switch" to show a switch icon
# 'command=toggle_mode' is a command that will execute the Function
#that we will manually create called 'toggle_mode' which helps to change mode
#from Light to Dark mode window
mode_switch = ttk.Checkbutton(widgets_frame, text = "Mode", style='Switch',
                              command=toggle_mode)
mode_switch.grid(row=6, column=0, padx=5, pady=5, sticky='nsew')

#
##
###
####
#####
#*************** RIGHT SIDE OF THE WINDOW ********************
#*************************************************************

#********* VIEWABLE TABLE ON THE RIGHT SIDE WINDOW **********

#This is the table on the right hand side of the window
#The placement of this portion is inside 'frame' which is the PARENT WINDOW (root)
treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)

#*************** SCROLL BAR & MOVING WITH THE TABLE CONTENT ********************

#this will create a scrollable bar on the rigt hand side of the window
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side='right', fill='y')

#*************** ADDING THE HEADER NAMES ********************

#this will create the HEADER of each COLUMN of the table
#'yscrollcommand=treeScroll.set' helps to ensure that the table will move when the scroll bar was moved
cols = ("Name", "Age", "Country", "Employment")
treeview = ttk.Treeview(treeFrame, show="headings",
                        yscrollcommand=treeScroll.set, columns=cols, height=13)
#these are the columns with the width of the table
#we can't see them yet in the preview window
treeview.column("Name", width=100)
treeview.column("Age", width=100)
treeview.column("Country", width=100)
treeview.column("Employment", width=100)
treeview.pack() #***** When to know when to use .PACK or .GRID?

#this will help to move the Treeview (table) if the user scrolls the bar up and down
treeScroll.config(command=treeview.yview)

#Calling this function will load the data
load_data()


#event loop that will help to launch our application
root.mainloop()
