from datetime import datetime
from datetime import date
# from email.policy import strict
# from hashlib import new
# from operator import imod
import tkinter
from tkinter.font import BOLD
import tkinter.messagebox
import pickle
import numpy as np
from numpy import array, pad
import os
import xlsxwriter

root = tkinter.Tk()
root.title("Task Manager")
logo = tkinter.PhotoImage(file='Images/logo.png')
root.iconphoto(False, logo)

#Functions
def addTasks():
    task = entryTask.get()  
    if task != "":
        # datetime object containing current date and time
        today = date.today()
        dt = today.strftime("%b-%d-%Y")
        taskPrev = task +"                          Created - " + dt
        tasks = listboxTask.get(0, listboxTask.size())
        if taskPrev not in tasks:
            listboxTask.insert(tkinter.END, task+"                          Created - " + dt)
            entryTask.delete(0, tkinter.END)
        else:
            tkinter.messagebox.showwarning(title="Warning!", message="Task already exists.")
    else:
        tkinter.messagebox.showwarning(title="Warning!", message="You must enter a task.")

def deleteTasks():
    try:
        task = listboxTask.curselection()[0]
        taskName = listboxTask.get(listboxTask.curselection()[0])
        listboxTask.delete(task)
        try:
            fileRemove = f"./saves/{taskName}.dat"
            os.remove(fileRemove)
        except:
            print('No file created already')
        tkinter.messagebox.showinfo(title="Success!", message="Successfully deleted!")
    except:
        tkinter.messagebox.showwarning(title="Warning!", message="You must select a task.")

def loadTasks():
    cnt = 0
    try:
        tasks = pickle.load(open("./saves/tasks.dat", "rb"))
        for task in tasks:
            listboxTask.insert(tkinter.END, task)
            cnt += 1
        if cnt != 1:
            tkinter.messagebox.showinfo(title="Success!", message=f"{cnt} Tasks Loaded!")
        else:
            tkinter.messagebox.showinfo(title="Success!", message=f"{cnt} Task Loaded!")
    except:
        tkinter.messagebox.showwarning(title="Warning!", message="No tasks on file")

def saveTasks():
    tasks = listboxTask.get(0, listboxTask.size())
    if listboxTask.size() != 0:
        pickle.dump(tasks,open("./saves/tasks.dat","wb"))
        tkinter.messagebox.showinfo(title="Save!", message="Tasks Saved!")
    else:
        tkinter.messagebox.showwarning(title="Warning!", message="You must have tasks.")

def updateTasks():
    newTask=entryTask.get()
    today = date.today()
    dt = today.strftime("%b-%d-%Y")
    taskPrev = newTask +"                          Created - " + dt
    names = listboxTask.get(0, listboxTask.size())
    if newTask != "" and taskPrev not in names:
        taskIndex = listboxTask.curselection()
        taskName = listboxTask.get(listboxTask.curselection()[0])
        try: 
            newTask = newTask +"                          Created - " + dt
            listboxTask.delete(taskIndex)
            listboxTask.insert(taskIndex,newTask)
            tkinter.messagebox.showinfo(title="Success!", message="Successfully Updated!")
            try:
                old_name = f"./saves/{taskName}.dat"
                new_name = f"./saves/{newTask}.dat"
                os.rename(old_name, new_name)
            except:
                print("n há ficheiro")
        except: 
            tkinter.messagebox.showwarning(title="Warning!", message="Error updating task")
    elif taskPrev in names:
        tkinter.messagebox.showwarning(title="Warning!", message="Task already exists.")
    else:
        tkinter.messagebox.showwarning(title="Warning!", message="   Enter a new task description\n\t     OR\n            Select a Task")

def clearListbox():
    listboxTask.delete(0, listboxTask.size())

def sortAsc():

    allTasks = listboxTask.get(0, listboxTask.size())
    arrTasks = np.array(allTasks)
    arrTasks.sort()
    clearListbox()

    for task in arrTasks:
        listboxTask.insert(tkinter.END, task)

def sortDesc():

    allTasks = listboxTask.get(0, listboxTask.size())
    arrTasks = np.array(allTasks)
    arrTasks[::-1].sort()
    clearListbox()

    for task in arrTasks:
        listboxTask.insert(tkinter.END, task)

def closeApp():
    tasks = listboxTask.get(0, listboxTask.size())
    answer = tkinter.messagebox.askokcancel("Quit", "Do you want to quit?")
    if answer:
        pickle.dump(tasks,open("./saves/tasks.dat","wb"))
        workbook = xlsxwriter.Workbook('tasks.xlsx')
        worksheet = workbook.add_worksheet()
        i = 0
        for task in tasks:
            taskInfo = task.split("                          Created - ")
            worksheet.write(i, 0, taskInfo[0])
            worksheet.write(i, 1, taskInfo[1])
            i = i+1
        workbook.close()
        root.destroy()
        
# def openNewWindow():
#     newWindow = tkinter.Toplevel(root)
#     newWindow.title("New Window")
#     newWindow.geometry("300x300")
#     tkinter.Label(newWindow,text ="This is a new window").pack()

def pop_window(dummy_event):
    taskName = listboxTask.get(listboxTask.curselection()[0])
    top = tkinter.Toplevel(root)
    top.title(f'{taskName}')
    top.geometry("400x400")
    textBoxTop = tkinter.Text(top, height=20, width=50)
    textBoxTop.pack()
    top.resizable(0,0)
    #textBoxTop.insert(tkinter.END, f"{taskName}")
    buttonSaveInTask = tkinter.Button(top, text="Save", command=lambda: saveInTask(textBoxTop.get(1.0,"end"),top))
    buttonSaveInTask.pack(pady=5)
    try:
        path = f"./saves/{top.title()}.dat"
        description = pickle.load(open(path, "rb"))
        textBoxTop.insert(tkinter.END, f"{description}")
    except:
        print("No data on file!")

def saveInTask(textBox,window):
    description = textBox
    #if len(description) != 0:
    path = f"./saves/{window.title()}.dat"
    print(path)
    pickle.dump(description,open(path,"wb"))
    tkinter.messagebox.showinfo(title="Save!", message="Tasks Saved!")
    window.destroy()
    
#Create GUI

frameTasks = tkinter.Frame(root)
frameTasks.pack()

listboxTask = tkinter.Listbox(frameTasks, height=20,width=100)
listboxTask.bind("<Double-Button-1>", pop_window)
listboxTask.pack(side=tkinter.LEFT, pady=5)

scrollBarTasks = tkinter.Scrollbar(frameTasks)
scrollBarTasks.pack(side=tkinter.RIGHT, fill=tkinter.Y)

listboxTask.config(yscrollcommand=scrollBarTasks.set)
scrollBarTasks.config(command = listboxTask.yview)

entryTask = tkinter.Entry(root, width=100)
entryTask.pack(pady=5)

btnAdd = tkinter.PhotoImage(file='Images/btnAdd.png')
buttonAddTask = tkinter.Button(root, text="Add Task", command=addTasks,image=btnAdd,borderwidth = 0)
buttonAddTask.pack(pady=1)

btnDelete = tkinter.PhotoImage(file='Images/btnDelete.png')
buttonDeleteTask = tkinter.Button(root, text="Delete Task", command=deleteTasks, image=btnDelete, borderwidth = 0)
buttonDeleteTask.pack(pady=1)

btnUpdate = tkinter.PhotoImage(file='Images/btnUpdate.png')
buttonUpdateTask = tkinter.Button(root, text="Update Task", command=updateTasks,image=btnUpdate, borderwidth = 0)
buttonUpdateTask.pack(pady=1)

# btnLoad = tkinter.PhotoImage(file='Images/btnLoad.png')
# buttonLoadTask = tkinter.Button(root, text="Load Tasks", command=loadTasks, image=btnLoad, borderwidth = 0)
# buttonLoadTask.pack(pady=1)

# btnSave = tkinter.PhotoImage(file='Images/btnSave.png')
# buttonSaveTask = tkinter.Button(root, text="Save Tasks", command=saveTasks, image=btnSave, borderwidth = 0)
# buttonSaveTask.pack(pady=1)

btnAsc = tkinter.PhotoImage(file='Images/btnAsc.png')
buttonSortAsc = tkinter.Button(root, text="Sort Asc Task", command=sortAsc, image=btnAsc, borderwidth = 0)
buttonSortAsc.pack(side=tkinter.LEFT, pady=1)

btnDesc = tkinter.PhotoImage(file='Images/btnDesc.png')
buttonSortDesc = tkinter.Button(root, text="Sort Desc Task", command=sortDesc, image=btnDesc, borderwidth = 0)
buttonSortDesc.pack(side=tkinter.RIGHT, pady=1)

root.protocol("WM_DELETE_WINDOW", closeApp)
root.resizable(0,0)
root.after(100, loadTasks)
root.mainloop()