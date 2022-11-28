from tkinter import *
from tkinter import filedialog
import speech_recognition
import pyttsx3
import PyPDF2
from tkinter import *
import PyPDF2
import aspose.words as aw
from fpdf import FPDF
from tkinter import filedialog
import tkinter.messagebox as mbox
import moviepy.editor as mp
from PIL import Image, ImageTk
import moviepy
import moviepy.editor
import os
from tkinter import ttk
Window = Tk()
Window.geometry('700x700')






def resize_image(event):
    new_width = event.width
    new_height = event.height
    image = copy_of_image.resize((new_width, new_height))
    photo = ImageTk.PhotoImage(image)
    label.config(image = photo)
    label.image = photo #avoid garbage collection

image = Image.open('file.png')
copy_of_image = image.copy()
photo = ImageTk.PhotoImage(image)
label = ttk.Label(Window, image = photo)
label.bind('<Configure>', resize_image)
label.pack(fill=BOTH, expand = YES)

#Window.config(bg="light blue")
Window.title("Convertor app")
Window.config(bg="light blue")
# Add image file
page1= Label(Window, text="Convertor app",font=("Arial", 20), bg = "light blue", fg = "blue", borderwidth=3, relief="raised")
startingpagenumber = Entry(Window)
page1.place(relx=0.43,rely=0.1)

#startingpagenumber.place(relx=0.6,rely=0.1)


def file():
    path = filedialog.askopenfilename()
    book = open(path, 'rb')
    pdfreader = PyPDF2.PdfFileReader(book)
    pages = pdfreader.numPages
    speaker = pyttsx3.init()
    
    for i in range(int(startingpagenumber.get()), pages):
        page = pdfreader.getPage(i) 
        txt = page.extractText()
        speaker.say(txt)
        speaker.runAndWait()
page1= Label(Window, text=" Enter starting page number",font=("Arial", 13), bg = "light blue", fg = "red")
startingpagenumber = Entry(Window)
page1.place(relx=0.43,rely=0.17)
startingpagenumber.place(relx=0.7,rely=0.18)

B=Button(Window, text="Pdf to voice", command=file,font=("Arial", 13), bg = "light blue", fg = "red")
B.place(relx=0.43,rely=0.22)

UserVoiceRecognizer = speech_recognition.Recognizer()
def file1():
    while(1):
        try:
            with speech_recognition.Microphone() as UserVoiceInputSource:
                            UserVoiceRecognizer.adjust_for_ambient_noise(UserVoiceInputSource, duration=0.5)
 
            # The Program listens to the user voice input.
                            UserVoiceInput = UserVoiceRecognizer.listen(UserVoiceInputSource)
 
                            UserVoiceInput_converted_to_Text = UserVoiceRecognizer.recognize_google(UserVoiceInput)
                            UserVoiceInput_converted_to_Text = UserVoiceInput_converted_to_Text.lower()
                            print(UserVoiceInput_converted_to_Text)
    
        except KeyboardInterrupt:
                    print('A KeyboardInterrupt encountered; Terminating the Program !!!')
                    exit(0)
        except speech_recognition.UnknownValueError:
                            print("No User Voice detected OR unintelligible noises detected OR the recognized audio cannot be matched to text !!!")
                            Window.destroy()

B1=Button(Window, text="Speech to  text", command=file1,font=("Arial", 13), bg = "light blue", fg = "red")
B1.place(relx=0.43,rely=0.27)


#FILE_PATH = 'paper1.pdf'

def file2():
    path = filedialog.askopenfilename()
    try:
         with open(path, mode='rb') as f:
        
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc)
            
            
            
            reader = PyPDF2.PdfFileReader(f)

            page = reader.getPage(1)

            print(page.extractText())
            builder.write(page.extractText())
            doc.save("pdftoword.docx")
    except:
        mbox.showinfo("Failed","\n Input only pdf")
B1=Button(Window, text="Convert pdf to word(docx)", command=file2,font=("Arial", 13), bg = "light blue", fg = "red")
B1.place(relx=0.43,rely=0.32)


from docx2pdf import convert

    
 
# Converting docx present in the same folder
# as the python file


def file3():
    
    path = filedialog.askopenfilename()
    book = open(path, 'rb')
    doc = aw.Document(book)
    doc.save("wordtopdf.pdf")
        
B1=Button(Window, text="Convert word to pdf", command=file3,font=("Arial", 13), bg = "light blue", fg = "red")
B1.place(relx=0.43,rely=0.37)





def mp4_choose():
    global filename,onlyfilename
    filename = filedialog.askopenfilename(initialdir="/", title="Choose MP4",filetypes=(("Text files", "*.Mp4*"), ("all files", "*.*")))
    # Change label contents
    # label_file_explorer.configure(text="File : " + filename)
    onlyfilename = os.path.basename(filename)


# created a choose button , to choose the image from the local system
chooseb = Button(Window, text='Select file for the following:', command=mp4_choose,font=("Arial", 13), bg = "light blue", fg = "red")
chooseb.place(relx=0.43,rely=0.42)

# Function for convert Mp4 to Mp3
def convert():
    
    video = moviepy.editor.VideoFileClip(filename)
    # Convert video to audio
    audio=video.audio

    aud_fname = ""
    for i in onlyfilename:
        if i == '.':
            break
        else:
            aud_fname = aud_fname + i
    print(aud_fname)
    audio.write_audiofile(f'{aud_fname}.mp3')
    mbox.showinfo("Success", "Video converted to Audio.\n\nAudio Saved Successfully")

# created a choose button , to choose the image from the local system
convertb = Button(Window, text='Convert MP4  to MP3',command=convert,font=("Arial", 13), bg = "light blue", fg = "red")
convertb.place(relx=0.43,rely=.47)
#B1=Button(Window, text="Convert mp4 to mp3", command=file4)
#B1.place(relx=0.35,rely=0.8)








from moviepy.editor import *
def file4():

    

    clip = (VideoFileClip(filename))
    clip.write_gif("output11.gif")

b2 = Button(Window, text="convert video to gif", command=file4,font=("Arial", 13), bg = "light blue", fg = "red")
b2.place(relx=0.43,rely=0.52)


def file5():

    clip = mp.VideoFileClip(filename)
    clip.write_videofile("myvideo11.mp4")



b2 = Button(Window, text="convert gif to video", command=file5,font=("Arial", 13), bg = "light blue", fg = "red")
b2.place(relx=0.43,rely=0.57)


def jpg_to_png():
	global im1

	# import the image from the folder
	import_filename = filedialog.askopenfilename()
	if import_filename.endswith(".jpg"):

		im1 = Image.open(import_filename)

		# after converting the image save to desired
		# location with the Extersion .png
		export_filename = filedialog.asksaveasfilename(defaultextension=".png")
		im1.save(export_filename)

		# displaying the Messaging box with the Success
		mbox.showinfo("success ", "your Image converted to Png")
	else:

		# if Image select is not with the Format of .jpg
		# then display the Error
		
		Label_2.place(x=80, y=280)
		mbox.showerror("Fail!!", "Something Went Wrong...")


def png_to_jpg():
    global im1
    import_filename = filedialog.askopenfilename()
    if import_filename.endswith(".png"):
        im1 = Image.open(import_filename)
        if im1.mode!="RGB":
            im1=im1.convert("RGB")
        export_filename = filedialog.asksaveasfilename(defaultextension=".jpg")
        im1.save(export_filename)
        mbox.showinfo("success ", "your Image converted to jpg ")
    else:
        

        mbox.showerror("Fail!!", "Something Went Wrong...")


button1 = Button(Window, text="JPG_to_PNG", command=jpg_to_png,font=("Arial", 13), bg = "light blue", fg = "red")

button1.place(relx=0.43,rely=0.62)

button2 = Button(Window, text="PNG_to_JPEG",  command=png_to_jpg,font=("Arial", 13), bg = "light blue", fg = "red")

button2.place(relx=0.43,rely=0.67)

import os
from PIL import Image, ImageOps
def conv():
    filename = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("jpeg files","*.jpg"),("png files","*.png"),("webp files","*.webp")))
    print(filename)
    if filename:
      try:
        basename = os.path.basename(filename) 
        Window.filename = os.path.splitext(basename)[0] + ".pdf"
        print (Window.filename)

#Use the Pillow library to convert the image to a PDF file


        im = Image.open(filename)

        im2 = ImageOps.fit(im, (612, 792), Image.ANTIALIAS)

        im2.save(Window.filename)

        mbox.showinfo("Success!", "File saved successfully as PDF")

      except: 
       mbox.showerror("Error!", "There was an error converting the file")



but1=Button(text=" Convert JPEG to PDF ", command=conv,font=("Arial", 13), bg = "light blue", fg = "red")
but1.place(relx=0.43,rely=0.72)



def exit_win():
    if mbox.askokcancel("Exit", "Do you want to exit?"):
        Window.destroy()

# creating an exit button
exitB = Button(Window, text='EXIT', command=exit_win,font=("Arial", 15), bg = "light blue", fg = "red")
exitB.place(relx=0.45,rely=.8)
Window.mainloop()