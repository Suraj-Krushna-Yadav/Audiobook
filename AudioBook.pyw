# importing all requires modules, classes, functions etc.
import tkinter as tk
from tkinter import Tk
from tkinter import Button
from tkinter import Label
from tkinter import Radiobutton
from tkinter import messagebox
from tkinter import filedialog

import shutil

import os

import PyPDF2
import pyttsx3
import pygame
import easygui



# counter to check whether current pdf is first or second, initially it is 0
global coun
coun = 0



# INitialization of pygame mixer and pyttsx3 engine
pygame.mixer.init()
engine = pyttsx3.init()



# creating main window or root window
root = Tk()
root.title("AudioBook by 79,76,60")
root.geometry("700x500")
root.minsize(700,500)

# background img setting
txco="#003233"              #....... COLOR OF TEXT
bgco="#00F0CC"              #....... COLOR OF BACKGROUND
buttco="#00fa99"            #....... COLOR OF BACKGROUND OF BUTTON

root.configure(background=bgco)


# -------------------->>>>>>>>> defininng all functions

def open_window():
    
    # getting file path from user
    global pdf_path 
    pdf_path=easygui.fileopenbox()
    
    if validation_pdf():
        
        if pdf_info():
            
            if extraction_text():
                pass
            else:

                if messagebox.askokcancel("\aExtraction Error !","Extraction of PDF's text is unsuccessful !\nPlease, select another PDF"):
                    open_file()
        else :

            if messagebox.askokcancel("\aDecryption Error !","Decryption of PDF is unsuccessful !\nPlease, select another PDF"):
                open_file()         
    else:

        if messagebox.askretrycancel("\aNo PDF","The selected file is not a pdf !"):
            open_file()



def open_file():
    open_window()


def validation_pdf():
    return ".pdf" in pdf_path


def get_file_name():
    global name_of_file
    name_of_file = pdf_path.rfind("\\")
    name_of_file = pdf_path[name_of_file+1:-4]


# getting doc info, pg no and name and set this to label
def pdf_info():
    try:
        
        # opening of file or creating file object
        global pdf_file
        pdf_file = open(pdf_path,"rb")
        
        # reading the file
        global reader
        reader = PyPDF2.PdfFileReader(pdf_file)
        
        global no_pg
        no_pg = reader.getNumPages()
        
        xxxx = reader.getDocumentInfo()

        get_file_name()
        
        global info_label
        info_label ='File Name : '+name_of_file+"\nTotal Pages : "+str(no_pg)+"\n"
        
        for x,y in xxxx.items():
            if x=="/CreationDate" or x=="/ModDate":
                continue
            else:
                info_label += '\n' + x.replace('/','') + " : " + y.replace('/','')
        
        global time_need
        time_need = "\n\nAfter clicking on play, some time require for processing depending on PDF's size & processor version"
        info_label+=time_need
        
        # setting text to PDF_info_bar_label_2
        PDF_info_bar_label_2.config(text = info_label)
        
        return True
        
    except Exception:
        
        pdf_file.close()
        
        return False

    

# validation for extraction of text from pdf
def extraction_text():
    try:
        
        global pdf_text
        pdf_text=""
        
        if validation_pg_no():
            
            ini_pg_no = user_pg_no.get()
            
            if ini_pg_no != 0:
                ini_pg_no = ini_pg_no-1
                
            for no in range(ini_pg_no,no_pg):
                page = reader.getPage(no)
                pdf_text += page.extractText()
            
            
        return True
    
    except Exception:
        
        pdf_file.close()
        
        return False



def validation_pg_no():
    try:
        tr = user_pg_no.get()>=0 and user_pg_no.get()<=no_pg
        if tr:
            return tr
        else:
            return False
            messagebox.showwarning("\aError","Please provide correct pg no !")

    except:
        return False
        messagebox.showwarning("\aError","Please provide correct pg no !")
        


def Creating_AudioBook():
    global coun
    global outfile
    
    # ------ Deleting previous temp file if availaiable
    coun+=1
    
    if coun == 2:
        
        coun = 1
        
        if os.path.exists(outfile):
            stop()
            os.remove(outfile)
            
        else:
            pass
    
    
    # creating temp file name
    outfile = str(name_of_file)+".wav"
    
    
    #collecting voices from speaker engine
    voices = engine.getProperty('voices')
    li=[]
    for v in voices:
        li.append(v)

    
    # setting voice according to user input
    engine.setProperty('voice', li[user_voice.get()].id)




    # extracting text from pdf
    pdf_text=""
    
    if validation_pg_no():
        
        ini_pg_no = user_pg_no.get()
        
        if ini_pg_no != 0:
            ini_pg_no = ini_pg_no-1
            
        for no in range(ini_pg_no,no_pg):
            page = reader.getPage(no)
            pdf_text += page.extractText()
        
        
        # saving temporary .wav file
        engine.save_to_file(pdf_text , outfile)
        engine.runAndWait()
        
        return True
    
    else:
        return False
   
    
   
def play():
    try:
        if Creating_AudioBook():

            # removing non required text from label
            ttt = info_label.replace(time_need,"")
            PDF_info_bar_label_2.config(text=ttt)
            
            # setting high volume for out music
            pygame.mixer.music.set_volume(1)
            # loading music
            pygame.mixer.music.load(outfile)
            # playing
            pygame.mixer.music.play()
            
        else:
            messagebox.showwarning("\aError","Please provide correct pg no !")
   
    except Exception:
        
        if messagebox.askokcancel("\aBad PDF","PDF may contains image or unreadable data\nYou may select another PDF"):
            open_file()
        
            
        
def pause():
    pygame.mixer.music.pause()


def resume():
    pygame.mixer.music.unpause()


def stop():
    pygame.mixer.music.unload()
    pygame.mixer.music.stop()
  
    
def save():
    copy_file()
   
    
def copy_file():
    # we actually copy the temp file to user specific folder
    source1 = outfile
    destination1=filedialog.askdirectory()
    try:
        shutil.copy(source1,destination1)
        messagebox.showinfo('\aConfirmation', "File Saved !")

    except:
        if messagebox.askretrycancel("Restriction","Please choose another folder"):
            save()
        else:
            pass
        


def on_closing():
    if messagebox.askokcancel("\aQuit", "Do you want to quit?"):
        stop()
        try:
            # deleting temp file if present
            if os.path.exists(outfile):
                os.remove(outfile)
           
            else:
                pass
            root.destroy()
        except:
            root.destroy()
root.protocol("WM_DELETE_WINDOW", on_closing)




# ------------------------  All Widgets & positions  

select_pdf_button = Button(root, text = "Select PDF", foreground=txco, background=buttco, activebackground="yellow", command= open_file)
select_pdf_button.place(x=310, y=40)

Pg_no_entry_label = Label(root, text="Enter Initial Pg No", width=20, background=bgco, foreground=txco)
Pg_no_entry_label.place(x=20, y=130)

user_pg_no = tk.IntVar()
Pg_no_user_entry = tk.Entry(root, textvariable = user_pg_no, width = 10, background="light yellow")
Pg_no_user_entry.place(x=200, y=130)

Voice_entry_label = Label(root, text="Select Voice Type Of Reader", width=20, background=bgco, foreground=txco)
Voice_entry_label.place(x=20, y=220)

user_voice = tk.IntVar()
R1_male = Radiobutton(root, text = "Male Voice", variable = user_voice, value = 0, foreground=txco, background=bgco, activebackground="red")
R1_male.place(x=200, y=220)

R1_female = Radiobutton(root, text = "Female Voice", variable = user_voice, value = 1, foreground=txco, background=bgco, activebackground="pink")
R1_female.place(x=300, y=220)

PDF_info_bar_label_1 = Label(root, text="AudioBook Info Window", foreground=txco, bg = buttco, width = 32)
PDF_info_bar_label_1.place(x=450, y=105)

PDF_info_bar_label_2 = Label(root, bg="light yellow", width = 32, height = 20, wraplength=220)
PDF_info_bar_label_2.place(x=450, y=125)


play_button = Button(root, text = "Play", command=play, background=buttco, activebackground="green", foreground=txco, font="bold", )
play_button.place(x=70, y=320)

pause_button = Button(root, text = "Pause", command=pause, background=buttco, activebackground="pink", foreground=txco, font="bold", width = 8)
pause_button.place(x=200, y=320)

resume_button = Button(root, text = "Resume", command=resume, background=buttco, activebackground="orange", foreground=txco, font="bold", width = 8)
resume_button.place(x=282, y=320)

stop_button = Button(root, text = "Stop", command=stop, background=buttco, activebackground="red", foreground=txco, font="bold")
stop_button.place(x=70, y=380)

save_button = Button(root, text = "Save To Device", command=save, background=buttco, activebackground="orange", foreground=txco, font="bold", width = 17)
save_button.place(x=200, y=380)


root.mainloop()