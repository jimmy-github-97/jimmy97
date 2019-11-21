from tkinter import *
from PIL import Image,ImageTk
from tkinter import ttk
from tkinter import filedialog
import cv2 as cv
import numpy as np
#from matplotlib import pyplot as plt
from numpy import array
import xlsxwriter

window=Tk()
window.geometry("2048x600")  #size of window
window.title("FYP") #title of window
#variable to store the value of first and last point
fm = StringVar()
fy = StringVar()
fr = StringVar()
fr_unit = StringVar()
sm = StringVar()
sy = StringVar()
sr = StringVar()
sr_unit = StringVar()

def print1():
    first_month = int(fm.get())
    first_year = int(fy.get())
    first_rate = float(fr.get())
    first_unit = fr_unit.get()
    sec_month = int(sm.get())
    sec_year = int(sy.get())
    sec_rate = float(sr.get())
    sec_unit = sr_unit.get()
    print(f"First point is is ({first_month},{first_year},{first_rate}{first_unit}).")
    print(f"Last point is ({sec_month},{sec_year},{sec_rate}{sec_unit}).")

def exit1():
    exit()

#Get image from user
def get_image():
    filename=filedialog.askopenfilename(initialdir="/",title="Select A File",filetype=(("jpeg","*.jpg"),("png","*.png"),("All Files","*.*")))
    label_dir=Label(text=filename)
    label_dir.place(x=200,y=500)
    load=Image.open(filename)
    render=ImageTk.PhotoImage(load)
    img=Label(window,image=render)
    img.image = render
    img.place(x=450,y=50)
    input_img = cv.imread(filename)
    cv.imshow('Original',input_img)
    img2 = cv.GaussianBlur(input_img, (5, 5), 0)
    cv.imshow('blur filter', img2)
    gray = cv.cvtColor(img2, cv.COLOR_BGR2GRAY)
    cv.imshow('gray', gray)
    th2 = cv.adaptiveThreshold(gray, 255, cv.ADAPTIVE_THRESH_GAUSSIAN_C, cv.THRESH_BINARY, 5, 2)
    cv.imshow('th2', th2)
    hsv = cv.cvtColor(img2, cv.COLOR_BGR2HSV)
    lower_red = np.array([0, 88, 153])
    upper_red = np.array([76, 255, 255])
    mask2 = cv.inRange(hsv, lower_red, upper_red)
    contours, hierarchy = cv.findContours(mask2, cv.RETR_TREE, cv.CHAIN_APPROX_SIMPLE)
    cv.drawContours(img, contours, -1, (0, 255, 0), 1)
    contours = array(contours)
    array1 = contours[0].reshape(-1, 2)
    array2 = contours[1].reshape(-1, 2)
    array3 = contours[2].reshape(-1, 2)
    array4 = contours[3].reshape(-1, 2)
    array5 = contours[4].reshape(-1, 2)
    combined = np.concatenate((array1, array2, array3, array4, array5))
    outWorkbook = xlsxwriter.Workbook("imgpixel.xlsx")
    outSheet = outWorkbook.add_worksheet()
    # write headers
    outSheet.write("A1", "X")
    outSheet.write("B1", "Y")
    sort_combined = sorted(combined, key=lambda x: x[0])
    pixel_location = []
    # take out the similar point
    for i in range(len(sort_combined) - 1):
        if sort_combined[i + 1][0] != sort_combined[i][0]:
            pixel_location.append(sort_combined[i])

    for row in range(len(pixel_location)):
        outSheet.write(row + 1, 0, pixel_location[row][0])
        outSheet.write(row + 1, 1, pixel_location[row][1])
    y_interval = ((sec_rate - firs_rate)/(pixel_location[row][1]-pixel_location[1][1]))


    outWorkbook.close()
    cv.waitKey()
    cv.destroyAllWindows()

def compute():
    exit()

def graphical_UI():
    label0=Label(window,text="Determination of rolling return",fg="blue",bg="yellow",relief="solid",font=("arial",16,"bold"))
    label0.pack(fill="x")
    label1=Label(window,text="Please insert your graph image",fg="blue",bg="yellow",relief="solid",font=("arial",16,"bold"))
    label1.place(x=0,y=400)

    #user insert the first and last point
    label2=Label(window,text="Insert Date of First Point:",fg="blue",bg="yellow",relief="solid",font=("arial",16,"bold"))
    label2.place(x=0,y=100)
    list_month = ['1','2','3','4','5','6','7','8','9','10','11','12']
    droplist_fm = OptionMenu(window,fm,*list_month)
    fm.set("Select Month")
    droplist_fm.config(width=15)
    droplist_fm.place(x=300,y=100)
    list_year = ['2009','2010','2011','2012','2013','2014','2015','2016','2017','2018','2019']
    droplist_fy = OptionMenu(window,fy,*list_year)
    fy.set("Select Year")
    droplist_fy.config(width=15)
    droplist_fy.place(x=450,y=100)

    label3 = Label(window, text="Insert rate of First Point:", fg="blue", bg="yellow", relief="solid",
                   font=("arial", 16, "bold"))
    label3.place(x=0, y=150)
    entry_fr = Entry(window, textvar=fr)
    entry_fr.place(x=300, y=150)
    list_unit = ['RM',['%']]
    droplist_unit = OptionMenu(window,fr_unit,*list_unit)
    fr_unit.set("Select Unit")
    droplist_unit.config(width=10)
    droplist_unit.place(x=450,y=150)

    label4 = Label(window, text="Insert Date of second Point:", fg="blue", bg="yellow", relief="solid",
                   font=("arial", 16, "bold"))
    label4.place(x=0, y=200)
    droplist_sm = OptionMenu(window, sm, *list_month)
    sm.set("Select Month")
    droplist_sm.config(width=15)
    droplist_sm.place(x=300, y=200)
    droplist_sy = OptionMenu(window,sy,*list_year)
    sy.set("Select Year")
    droplist_sy.config(width=15)
    droplist_sy.place(x=450,y=200)

    label5 = Label(window, text="Insert rate of Second Point:", fg="blue", bg="yellow", relief="solid",
                   font=("arial", 16, "bold"))
    label5.place(x=0, y=250)
    entry_sr = Entry(window, textvar=sr)
    entry_sr.place(x=300, y=250)
    list_unit = ['RM',['%']]
    droplist_unit2 = OptionMenu(window,sr_unit,*list_unit)
    sr_unit.set("Select Unit")
    droplist_unit2.config(width=10)
    droplist_unit2.place(x=450,y=250)

graphical_UI();
button1=Button(window,text="Browse",width=12,fg="black",bg="white",command=get_image)
button1.place(x=50,y=450)
button3=Button(window,text="Enter",width=12,fg="black",bg="white",command=print1)
button3.place(x=50,y=300)
button2=Button(window,text="Exit",fg="black",bg="white",command=exit1)
button2.place(x=200,y=450)
button4=Button(window,text="Compute",width=12,fg="black",bg="white",command=compute)
button4.place(x=50,y=550)
window.mainloop()