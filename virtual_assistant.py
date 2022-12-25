import pyttsx3 as pyt
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import smtplib
import pywhatkit
from dotenv import load_dotenv
import requests
import subprocess as sp
from PyQt5 import QtWidgets,QtCore,QtGui
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType

load_dotenv(dotenv_path=r'C:\Users\dell\Desktop\jitu programming\python programming\data.env')
name=os.getenv('NAME')
newskey=os.getenv('NEWS_API_KEY')
weatherkey=os.getenv('WEATHER_API_KEY')

global paths
paths={
    ' ms word':'C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE',
    ' ppt':'C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE',
    ' excel':'C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE',
    ' chrome':'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    ' command prompt':'C:\Windows\system32\cmd.exe'
}

engine=pyt.init("sapi5")
voice=engine.getProperty("voices")
engine.setProperty("voice",voice[0].id)
greetings=['good morning','good night','good evening','hello','hey']



####################################################################################################

#This function will not work properly due to invalid input pertaining to security concerns#
def sendmail(to,context):
    server=smtplib.SMTP_SSL("smtp.gmail.com",465)
    server.ehlo()
    server.starttls()
    server.login("my-email-address","my-password")
    server.sendmail("my-email-address",to,context)
    server.close()
    speak("Email has been successfully sent")

########################################################################################################

def speak(text):
    engine.say(text)
    engine.runAndWait()

def wishme():
    time=int(datetime.datetime.now().hour)
    if time>=0 and time<12:
        speak(f'Good morning boss, I am {name}, how can i help you?')
    elif time>=12 and time<18:
        speak(f'Good afternoon boss, I am {name}, how can i help you?')
    else:
        speak(f'Good evening boss, I am {name}, how can i help you?')

def takecmmnd():
    r=sr.Recognizer()

    with sr.Microphone() as source:
        print("Listening\n")
        r.pause_threshold=2
        audio=r.listen(source)
    try:
        print("Recognizing\n")
        qry=r.recognize_google(audio,language="en-in")
        print("User said:",qry)
    except Exception as e:
        print("Some error occured, please say that again")
        return "none"
    return qry



def ask():
    ques=takecmmnd()
    return ques



def execute():
    wishme()
    while True:
        qry=takecmmnd().lower()
        if "wikipedia" in qry:
            speak("Searching wikipedia....")
            qry=qry.replace("wikipedia","")
            result=wikipedia.summary(qry,sentences=2)
            speak("According to wikipedia")
            print(result)
            speak(result)

            
        elif "describe yourself" in qry:
            print("My name is John, I am a virtual assistant developed with python by master Ajit Kumar Mishra, the legend ")
            speak("My name is John, I am a virtual assistant developed with python by master Ajit Kumar Mishra, the legend ")


        elif "open youtube" in qry:
            webbrowser.open("youtube.com")
        elif "open google" in qry:
            webbrowser.open("google.com")
        elif "song" in qry:
            speak("searching internet")
            qry=qry.replace("song","")
            speak(f'playing {qry}')
            pywhatkit.playonyt(qry)
        elif "the time" in qry:
            sttime=datetime.datetime.now().strftime("%I:%M:%S %p")
            speak(f'The time is:{sttime}')

        elif "open" in qry:
            qry=qry.replace("open","").lower()
            print(f'Opening {qry}')
            speak(f'Opening {qry}')
            if qry in paths.keys():
                os.startfile(paths[qry])
            else:
                print("Can't open the app")
                speak("Can't open the app")
        elif "open camera" in qry:
            print("Opening camera")
            speak("Opening camera")
            sp.run('start microsoft.windows.camera:',shell=True)
        
        elif "system shutdown" in qry:
            speak("shutting down,")
            os.system('shutdown /s /t 1')
        elif "system restart" in qry:
            speak("restarting,")
            os.system('shutdown /r /t 1')
        elif "system log out" in qry:
            speak("logging out,")
            os.system('shutdown -1')
        elif  "exit" in qry:
            speak("Pleasure serving you boss")
            exit()
        elif "what" in qry:
            try:
                speak("Searching internet")
                result1=pywhatkit.search(qry)
                print(result1)
                speak(result1)
            except:
                speak("Couldn't find anything")


        elif "send mail" in qry:
            try:
                speak("Please specify reciever email address")
                to=takecmmnd()
                speak("What should i send")
                context=takecmmnd()
                speak("Sending mail")
                sendmail(to,context)
            except:
                speak("Sorry,an error occured,can't send the mail")


        elif "weather" in qry:
            speak("For which city should i show the weather status for?")
            city=ask()
            speak("Gathering weather data")
            res1=requests.get(f'http://api.openweathermap.org/data/2.5/weather?q={city}&appid={weatherkey}&units=metric').json()
            weather = res1["weather"][0]["main"]
            temperature = res1["main"]["temp"]
            feels_like = res1["main"]["feels_like"]
            print(f'Showing weather status of{city}')
            speak(f'Showing weather status of{city}')
            print(f"The weather outside is {weather},with temperature's of {temperature} but it feels like {feels_like}")
            speak(f"The weather outside is {weather},with temperature's of {temperature} but it feels like {feels_like}")


        elif "news" in qry:
            speak("Today's headlines are")
            newheadlines=[]
            res=requests.get(f'https://newsapi.org/v2/top-headlines?country=in&apiKey={newskey}&category=general').json()
            articles=res['articles']
            for article in articles:
                newheadlines.append(article['title'])
            print(newheadlines[:5])
            speak(newheadlines[:5])

                    
        elif "good evening" in qry:
            speak("good evening boss")
        

        elif "jokes" in qry:
            headers={
                'Accept':'application/json'
            }
            res2=requests.get("https://icanhazdadjoke.com/",headers=headers).json()
            print(res2["joke"])
            speak(res2["joke"])
            
        elif "advice" in qry:
            res3=requests.get("https://api.adviceslip.com/advice").json()
            print(res3['slip']['advice'])
            speak(res3['slip']['advice'])
    
        else:
            print("How else can i help you?")
            speak("How else can i help you?")
            
if __name__=='__main__':
    execute()