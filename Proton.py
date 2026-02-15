import pyttsx3
import speech_recognition as sr
from datetime import date
import time
import webbrowser
import datetime
from pynput.keyboard import Key, Controller
import pyautogui
import sys
import os
from os import listdir
from os.path import isfile, join
import smtplib
import wikipedia
import Gesture_Controller
#import Gesture_Controller_Gloved as Gesture_Controller
import app
from threading import Thread
import psutil
import cv2
import numpy as np




# -------------Object Initialization---------------
today = date.today()
r = sr.Recognizer()
keyboard = Controller()
engine = pyttsx3.init('sapi5')
engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

# ----------------Variables------------------------
file_exp_status = False
files =[]
path = ''
is_awake = True  #Bot status

# ------------------Functions----------------------
def reply(audio):
    app.ChatBot.addAppMsg(audio)

    print(audio)
    engine.say(audio)
    engine.runAndWait()


def wish():
    hour = int(datetime.datetime.now().hour)

    if hour>=0 and hour<12:
        reply("Good Morning!")
    elif hour>=12 and hour<18:
        reply("Good Afternoon!")   
    else:
        reply("Good Evening!")  
        
    reply("I am Proton, how may I help you?")

# Set Microphone parameters
with sr.Microphone() as source:
        r.energy_threshold = 500 
        r.dynamic_energy_threshold = False

# Audio to String
def record_audio():
    with sr.Microphone() as source:
        r.pause_threshold = 0.8
        voice_data = ''
        audio = r.listen(source, phrase_time_limit=5)

        try:
            voice_data = r.recognize_google(audio)
        except sr.RequestError:
            reply('Sorry my Service is down. Plz check your Internet connection')
        except sr.UnknownValueError:
            print('cant recognize')
            pass
        return voice_data.lower()

    import webbrowser

# Function to take a screenshot
def take_screenshot():
    screenshot = pyautogui.screenshot()
    screenshot.save("screenshot.png")
    reply("Screenshot taken and saved as screenshot.png")

# Function to start and stop screen recording
# Function to start and stop screen recording with error handling
def start_screen_recording():
    try:
        reply("Screen recording started. Say 'stop recording' to end.")
        screen_size = pyautogui.size()
        fourcc = cv2.VideoWriter_fourcc(*"XVID")
        out = cv2.VideoWriter("screen_recording.avi", fourcc, 10.0, screen_size)
        
        while True:
            img = pyautogui.screenshot()
            frame = np.array(img)
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            out.write(frame)

            # Check if 'stop recording' command is given
            if 'stop recording' in record_audio():
                reply("Stopping the screen recording.")
                break

    except Exception as e:
        reply(f"An error occurred during screen recording: {e}")

    finally:
        out.release()
        reply("Screen recording saved as screen_recording.avi")


# Function to search for a video on YouTube
    def play_youtube_video(query):
        youtube_search_url = f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"
        webbrowser.open(youtube_search_url)
        return f"Searching for {query} on YouTube and playing the first video."

# Executes Commands (input: string)
def respond(voice_data):
    global file_exp_status, files, is_awake, path
    print(voice_data)
    voice_data.replace('proton','')
    app.eel.addUserMsg(voice_data)

    if is_awake==False:
        if 'wake up' in voice_data:
            is_awake = True
            wish()

    # STATIC CONTROLS
    elif 'hello' in voice_data:
        wish()

    elif 'what is your name' in voice_data:
        reply('My name is Proton!')

    elif 'date' in voice_data:
        reply(today.strftime("%B %d, %Y"))

    elif 'time' in voice_data:
        reply(str(datetime.datetime.now()).split(" ")[1].split('.')[0])

    elif 'search' in voice_data:
        reply('Searching for ' + voice_data.split('search')[1])
        url = 'https://google.com/search?q=' + voice_data.split('search')[1]
        try:
            webbrowser.get().open(url)
            reply('This is what I found Sir')
        except:
            reply('Please check your Internet')

    elif 'location' in voice_data:
        reply('Which place are you looking for ?')
        temp_audio = record_audio()
        app.eel.addUserMsg(temp_audio)
        reply('Locating...')
        url = 'https://google.nl/maps/place/' + temp_audio + '/&amp;'
        try:
            webbrowser.get().open(url)
            reply('This is what I found Sir')
        except:
            reply('Please check your Internet')

    elif ('bye' in voice_data) or ('by' in voice_data):
        reply("Good bye Sir! Have a nice day.")
        is_awake = False

    elif ('exit' in voice_data) or ('terminate' in voice_data):
        if Gesture_Controller.GestureController.gc_mode:
            Gesture_Controller.GestureController.gc_mode = 0
        app.ChatBot.close()
        #sys.exit() always raises SystemExit, Handle it in main loop
        sys.exit()
        
    
    # DYNAMIC CONTROLS
    elif 'launch gesture recognition' in voice_data:
        if Gesture_Controller.GestureController.gc_mode:
            reply('Gesture recognition is already active')
        else:
            gc = Gesture_Controller.GestureController()
            t = Thread(target = gc.start)
            t.start()
            reply('Launched Successfully')

    elif ('stop gesture recognition' in voice_data) or ('top gesture recognition' in voice_data):
        if Gesture_Controller.GestureController.gc_mode:
            Gesture_Controller.GestureController.gc_mode = 0
            reply('Gesture recognition stopped')
        else:
            reply('Gesture recognition is already inactive')
        
    elif 'copy' in voice_data:
        with keyboard.pressed(Key.ctrl):
            keyboard.press('c')
            keyboard.release('c')
        reply('Copied')
          
    elif 'page' in voice_data or 'pest'  in voice_data or 'paste' in voice_data:
        with keyboard.pressed(Key.ctrl):
            keyboard.press('v')
            keyboard.release('v')
        reply('Pasted')
    
    # MICROSOFT APPLICATIONS
    elif 'open word' in voice_data:
        reply('Opening Microsoft Word')
        os.system(r'"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"')

    elif 'open excel' in voice_data:
        reply('Opening Microsoft Excel')
        os.system(r'"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"')

    elif 'open powerpoint' in voice_data:
        reply('Opening Microsoft PowerPoint')
        os.system(r'"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"')

    elif 'open outlook' in voice_data:
        reply('Opening Microsoft Outlook')
        os.system(r'"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"')

    elif 'open onenote' in voice_data:
        reply('Opening Microsoft OneNote')
        os.system(r'"C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE"')

    elif 'shutdown' in voice_data or 'shut down' in voice_data:
        reply("Shutting down the computer. Goodbye!")
        os.system('shutdown /s /t 1')

    elif 'hibernate' in voice_data:
        reply("Hibernating the computer. Goodbye!")
        os.system('shutdown /h')


    elif 'restart' in voice_data:
        reply("Restarting the computer.")
        os.system('shutdown /r /t 1')
        
    elif 'log off' in voice_data or 'sign out' in voice_data:
        reply("Logging off.")
        os.system('shutdown /l')

# Open Camera
    elif 'open camera' in voice_data:
        reply('Opening Camera')
        os.system('start microsoft.windows.camera:')

# Open Clock
    elif 'open clock' in voice_data:
        reply('Opening Clock')
        os.system('start ms-clock:')

# Open Flipkart Application
    elif 'open flipkart' in voice_data:
        reply('Opening Flipkart')
        os.system(r'"C:\Path\To\FlipkartApp.exe"')  # Replace with the correct path to your Flipkart app

# Open Microsoft Store
    elif 'open microsoft store' in voice_data:
        reply('Opening Microsoft Store')
        os.system('start ms-windows-store:')

    elif 'open whatsapp' in voice_data:
        reply('Opening WhatsApp')
        os.system('start whatsapp:')

# Open Power BI
    elif 'open power bi' in voice_data:
        reply('Opening Power BI')
        os.system(r'"C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe"')


    elif 'battery' in voice_data or 'battery percentage' in voice_data or 'battery status' in voice_data:
        battery = psutil.sensors_battery()
        if battery:
            percentage = battery.percent
            reply(f"The current battery percentage is {percentage} percent.")
        if battery.power_plugged:
            reply("Your system is currently plugged in.")
        else:
            reply("Your system is not plugged in.")  

    # Handle other commands...


    # File Navigation (Default Folder set to C://)
    elif 'list' in voice_data:
        counter = 0
        path = 'C://'
        files = listdir(path)
        filestr = ""
        for f in files:
            counter+=1
            print(str(counter) + ':  ' + f)
            filestr += str(counter) + ':  ' + f + '<br>'
        file_exp_status = True
        reply('These are the files in your root directory')
        app.ChatBot.addAppMsg(filestr)
        
    elif file_exp_status == True:
        counter = 0   
        if 'open' in voice_data:
            if isfile(join(path,files[int(voice_data.split(' ')[-1])-1])):
                os.startfile(path + files[int(voice_data.split(' ')[-1])-1])
                file_exp_status = False
            else:
                try:
                    path = path + files[int(voice_data.split(' ')[-1])-1] + '//'
                    files = listdir(path)
                    filestr = ""
                    for f in files:
                        counter+=1
                        filestr += str(counter) + ':  ' + f + '<br>'
                        print(str(counter) + ':  ' + f)
                    reply('Opened Successfully')
                    app.ChatBot.addAppMsg(filestr)
                    
                except:
                    reply('You do not have permission to access this folder')
                                    
        if 'back' in voice_data:
            filestr = ""
            if path == 'C://':
                reply('Sorry, this is the root directory')
            else:
                a = path.split('//')[:-2]
                path = '//'.join(a)
                path += '//'
                files = listdir(path)
                for f in files:
                    counter+=1
                    filestr += str(counter) + ':  ' + f + '<br>'
                    print(str(counter) + ':  ' + f)
                reply('ok')
                app.ChatBot.addAppMsg(filestr)
                   
    

# Integrating with voice command in your Proton app
    elif 'play' in voice_data:
        song_or_video = voice_data.replace('play', '').strip()
        reply(f"Opening YouTube and searching for {song_or_video}.")
        play_youtube_video(song_or_video)

    elif 'screenshot' in voice_data:
        take_screenshot()

    elif 'start recording' in voice_data:
        start_screen_recording()

    
    else: 
        reply('This feature will be soon in Proton')
        reply('Stay tuned...')


# ------------------Driver Code--------------------

t1 = Thread(target = app.ChatBot.start)
t1.start()

# Lock main thread until Chatbot has started
while not app.ChatBot.started:
    time.sleep(0.5)

wish()
voice_data = None
while True:
    if app.ChatBot.isUserInput():
        #take input from GUI
        voice_data = app.ChatBot.popUserInput()
    else:
        #take input from Voice
        voice_data = record_audio()

    #process voice_data
    if 'proton' in voice_data:
        try:
            #Handle sys.exit()
            respond(voice_data)
        except SystemExit:
            reply("Thank you for choosing Proton")
            reply("Have a nice day ahead")
            break
        except:
            #some other exception got raised
            print("EXCEPTION raised while closing.") 
            break
        


