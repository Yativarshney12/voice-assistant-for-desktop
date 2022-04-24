import pyttsx3
import speech_recognition as sr
import datetime
import pyautogui
import os
import time
import wikipedia
import webbrowser
import sys
import pyjokes
import ctypes
import pywhatkit as kit
import winshell
import requests
from requests import get
from urllib.request import urlopen
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import QTimer, QTime, Qt
from PyQt5.QtGui import QMovie
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from assistantUi import Ui_assistantUi


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

#text to speech
def speak(audio):
    engine.say(audio)
    print(audio)
    engine.runAndWait()


#to wish
def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")   

    else:
        speak("Good Evening!") 
    
    assname = ("fellow")
    speak("how can i help you,Sir")

class MainThread(QThread):
    def __init__(self):
        super(MainThread,self).__init__()

    def run(self):
        self.TaskExecution()


    #to convert voice into text
    def takeCommand(self):

        r = sr.Recognizer()
        with sr.Microphone() as source:
            print("Listening...")
            r.pause_threshold = 1
            audio = r.listen(source)

        try:
            print("Recognizing...")    
            self.query = r.recognize_google(audio, language='en-in')
            print(f"Me: {self.query}\n")

        except Exception as e:    
            speak("Sorry Sir Say that again please...")  
            return "None"
        return self.query

    def TaskExecution(self):
        wishMe()
        while True:
        #if 1:

            self.query = self.takeCommand().lower()

            #to start the program of this program...
            if 'hello' in self.query:
                speak("hello Sir what can i help you")
                ord = self.takeCommand().lower()

            #to open the wikipedia.........
            elif 'wikipedia' in self.query:
                speak('Searching Wikipedia...')
                self.query = self.query.replace("wikipedia", "")#to correct in iet project....
                results = wikipedia.summary(self.query, sentences=2)
                speak("According to Wikipedia")
                print(results)
                speak(results)

            elif 'open google' in self.query:
               speak("Sir,What should i search")
               srch = self.takeCommand().lower()
               webbrowser.open(f"{srch}")


            #to change the voice of jarvis....
            elif 'change your voice' in self.query:#to correct in iet project....
                if 'male' in self.query:
                    engine.setProperty('voice', voices[0].id)
                else:
                    engine.setProperty('voice', voices[1].id)
                    speak("Hello Sir, I have switched my voice. How is it?") 
        
            #to increase or decrease the volume....
            elif 'volume up' in self.query:
                speak("volume increasing")
                pyautogui.press("volumeup")

            elif 'volume down' in self.query:
                speak("volume decreasing")
                pyautogui.press("volumedown")

            elif 'volume mute' in self.query:
                speak("volume mute")
                pyautogui.press("volumemute")


            #to ask the battery remaining in system...
            elif 'how much power left' in self.query or 'how much power we have' in self.query or 'battery' in self.query:
                import psutil
                battery = psutil.sensors_battery()
                percentage = battery.percent
                speak(f"sir our system have {percentage} percent battery") 

            #to switch the window from one tab to another...
            elif "switch the window" in self.query or "switch window" in self.query:
                speak("Okay sir, Switching the window")
                pyautogui.keyDown("alt")
                pyautogui.press("tab")
                time.sleep(1)
                pyautogui.keyUp("alt")


            #to take screenshots of display.....
            elif 'screenshot' in self.query or 'take screenshot' in self.query:
                speak("By what name do you want to save the screenshot?")
                name = self.takeCommand()
                image = pyautogui.screenshot()
                name = f"{name}.png"
                image.save(name)
                speak('Screenshot taken.')

                '''elif "show me the picture" in self.query:
                    try:
                        img = file.open('D:\\mini project\\assistant\\' + name)
                        img.show(img)
                        speak("Here it is sir")
                        time.sleep(2)

                    except IOError:
                        speak("Sorry sir, I am unable to display the screenshot")'''

            
            #to play the music of the computer....
            elif 'play music' in self.query:
                music_dir = 'D:\\music'
                songs = os.listdir(music_dir)
                print(songs)    
                os.startfile(os.path.join(music_dir, songs[0]))


            #to tell the current time....
            elif 'the time' in self.query:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")    
                speak(f"Sir, the current time is {strTime}")
                
                
            #to play music from directory..
            elif 'play music' in self.query:
                music_dir = 'D:\\music'
                songs = os.listdir(music_dir)
                print(songs)    
                os.startfile(os.path.join(music_dir, songs[0]))


            #to open n close the notepad.....
            elif 'open notepad' in self.query:
                ntpd = "C:\\Windows\\System32\\notepad.exe"
                os.startfile(ntpd)

            elif 'close notepad' in self.query:
                speak("okay sir ... closing notepad")
                os.system("taskkill /f /im notepad.exe")


            #to clear the recycle bin 
            elif 'empty recycle bin' in self.query:
                winshell.recycle_bin().empty(confirm=True, show_progress=True)
                speak("clearing recycle bin")
                speak("recycle bin cleared successfully")

            #to tell the location of mine.... 
            elif "where is" in self.query:
               self.query = self.query.replace("where is", "")
               location = self.query
               speak("User asked to Locate")
               speak("location")
               webbrowser.open("https://www.google.nl/maps/place/" + location + "")

            #to open some apps of the system....
            elif 'open powerpoint' in self.query:
                speak("opening Power Point presentation")
                power = "C:\\Program Files (x86)\\Microsoft Office\\root\Office16\\POWERPNT.EXE"
                os.startfile(power)
            
            
            elif 'close powerpoint' in self.query:
                speak("okay sir ... closing powerpoint")
                os.system("taskkill /f /im POWERPNT.exe")

            elif ' open word' in self.query or 'MS Word' in self.query:
                speak("opening MS Word")
                os.system("start MS Word")
                
            elif 'close word' in self.query:
                speak("okay sir ... closing word")
                os.system("taskkill /f /im word.exe")

            elif ' open control panel' in self.query:
                speak("opening control panel")
                os.system("start Control Panel")
                
            
                           
            elif 'open chrome' in self.query:
                speak("opening chrome")
                chrm = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
                os.startfile(chrm)

            elif 'open command prompt' in self.query or 'open cmd' in self.query:
                speak("opening command promt")
                os.system("start cmd")


            elif 'open flipkart' in self.query:
                speak("opening flipkart")
                webbrowser.open("flipkart.com")

            elif 'open facebook' in self.query:
                speak("opening facebook")
                webbrowser.open("facebook.com")
        
            elif 'open chrome' in self.query:
                speak("opening chrome")
                chrm = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
                os.startfile(chrm)
            
            elif 'open instagram' in self.query:
                speak("opening instagram")
                webbrowser.open("instagram.com")
            

            #to play songs on youtube...
            elif 'play songs on youtube' in self.query:
              speak("Which song would you like to listen")
              sng = self.takeCommand().lower()
              kit.playonyt(f"{sng}")
              
            
            
            
            
            #to give a name first cll this...
            elif "change your name to" in self.query:
                self.query = self.query.replace("change my name to", "")
                assname = self.query
 
            #then this....
            elif "what's your name" in self.query or "What is your name" in self.query:
                speak("My friends call me")
                speak(assname)
                print("My friends call me", assname)
            
            #then at last this....
            elif "change name" in self.query:
                speak("What would you like to call me, Sir ")
                assname = self.takeCommand()
                speak("Thanks for naming me")

            #to hide files and folders
            elif 'hide all files' in self.query or 'hide this folder' in self.query or 'visible for everyone' in self.query:
                speak("sir please tell me you want to hide this or make it visible for everyone")
                condition = self.takeCommand().lower()
                if 'hide' in condition:
                   os.system("attrib +h /s /d") #os module
                   speak("sir,all the files in this folder are now hidden")

                elif 'visible' in condition:
                   os.system("attrib -h /s /d")
                   speak("sir,all the files in this folder are now visible to everyone. i wish you are taking")

                elif 'leave it' in condition or 'leave for now' in condition:
                   speak("ok sir")

            
            #to keep waiting the program for some tym....
            elif "don't listen" in self.query or "stop listening" in self.query:
                speak("for how much seconds you want to stop jarvis from listening commands")
                a = int(self.takeCommand())
                time.sleep(a)
                print(a)


            #to remember a note in it's memory...
            elif "write a note" in self.query:
                speak("What should i write, sir")
                note = self.takeCommand()
                file = open('assistant.txt', 'w')
                speak("Sir, Should i include date and time")
                snfm = self.takeCommand()
                if 'yes' in snfm or 'sure' in snfm:
                    strTime = datetime.datetime.now().strftime("%H:%M:%S")
                    file.write(strTime)
                    file.write(" :- ")
                    file.write(note)
                else:
                    file.write(note)

            #to show the note which saved in it's memory....
            elif "show note" in self.query:
                speak("Showing Notes")
                file = open("assistant.txt", "r")
                print(file.read())
                
            
            #to tell a random joke....
            elif 'tell me a joke' in self.query:
                joke = pyjokes.get_joke()
                speak(joke)


            #to operate the power of the system....
            elif 'lock window' in self.query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()

            elif 'shutdown the system' in self.query:
                os.system("shutdown /s /t 5")

            elif 'restart the system' in self.query:
                os.system("shutdown /r /t 5")


            elif 'thank you' in self.query or 'thanks' in self.query:
                speak("It's my pleasure sir.")


            #to terminate from the program......
            elif  "goodbye" in self.query or "you can sleep" in self.query or "sleep" in self.query:
                speak("thanks for using me sir . Have a good day")
                sys.exit()




startExecution = MainThread()
            
class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_assistantUi()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.startTask)
        self.ui.pushButton_2.clicked.connect(self.close)
        

    def startTask(self):
        self.ui.movie = QtGui.QMovie("assistant_img/212508.gif")
        self.ui.label.setMovie(self.ui.movie)
        self.ui.movie.start()
        self.ui.movie = QtGui.QMovie("assistant_img/3Hb_eh.gif")
        self.ui.label_2.setMovie(self.ui.movie)
        self.ui.movie.start()
        self.ui.movie = QtGui.QMovie("assistant_img/SpanishFirstAmmonite-size_restricted.gif")
        self.ui.label_3.setMovie(self.ui.movie)
        self.ui.movie.start()
        self.ui.movie = QtGui.QMovie("assistant_img/hud_2.gif")
        self.ui.label_7.setMovie(self.ui.movie)
        self.ui.movie.start()
        self.ui.movie = QtGui.QMovie("assistant_img/c572a4f5-9930-44d3-9818-1d28ab0232b7.gif")
        self.ui.label_8.setMovie(self.ui.movie)
        self.ui.movie.start()
        timer = QTimer(self)
        timer.timeout.connect(self.showTime)
        timer.start(1000)
        startExecution.start()
    
    def showTime(self):
        current_time = QTime.currentTime()
        current_date = QDate.currentDate()
        label_time = current_time.toString('hh:mm:ss')
        label_date = current_date.toString(Qt.ISODate)
        self.ui.textBrowser.setText(label_date)
        self.ui.textBrowser_2.setText(label_time)




app = QApplication(sys.argv)
assistant = Main()
assistant.show()
exit(app.exec_())

   

        
