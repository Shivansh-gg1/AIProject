import subprocess
import sys
import wolframalpha
import pyttsx3
import tkinter
import json 
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import pywhatkit as kit
import socket

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
print(voices[0].id)
engine.setProperty('voices', voices[0].id)

#TEXT TO SPEECH
def speak(audio):
    engine.say(audio)
    print(audio)
    engine.runAndWait()

# To convert voice to text
def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 5
        audio = r.listen(source, timeout=5, phrase_time_limit=5)


    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}")
    
    except Exception as e:
        speak("Say that again please!")
        return "none"
    return query

# To Wish
def wish():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<=12:
        speak("Good Morning!")
    elif hour>12 and hour<18:   
        speak("Good Afternoon Sir!")
    else:
        speak("Good Evening Sir!")
    speak("I am Max! How can I help you?") 

# To Send Email
##def sendEmail(to, content):
   # server = smtplib.SMTP('smtp.gmail.com', 25)
    #server.connect('smtp.gmail.com', 587)
    #server.ehlo
    #server.starttls
    #server.login('shivanshsharma.6122004@gmail.com', 'Shivansh@123')
    #server.sendmail('shivanshsharma.6122004@gmail.com', to, content)
    #server.close()

if __name__ == "__main__":
    wish()
    while True:

        query = takecommand().lower()

        #logic building for tasks
        
        if "open notepad" in query:
            apath = "C:\\WINDOWS\\system32\\notepad.exe"
            os.startfile(apath)

        elif "open chrome" in query:
            bpath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
            os.startfile(bpath)

        elif "wikipedia" in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=5)
            speak("According to Wikipedia")
            print(results)
            speak(results)
        
        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")
 
        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")

        elif "who are you" in query:
            speak("I am your virtual assistant created by Shivaansh")
                
        elif "open edge" in query:
            cpath = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
            os.startfile(cpath)

        elif "open vs code" in query:
            dpath = "C:\\Users\\shiva\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(dpath)

        elif "open whatsapp" in query:
            epath = "C:\\Users\\Shivansh Sharma\\AppData\\Local\\WhatsApp\\WhatsApp.exe"
            os.startfile(epath)

        elif "open reader" in query:
            fpath = "C:\\Program Files (x86)\\Adobe\\Acrobat Reader DC\\Reader\\AcroRd32.exe"
            os.startfile(fpath)

        elif "open spotify" in query:
            gpath = "C:\\Users\\shiva\\AppData\\Roaming\\Spotify\\Spotify.exe"
            os.startfile(gpath)

        elif "open microsoft teams" in query:
            hpath = "C:\\Users\\shiva\\AppData\\Local\\Microsoft\\Teams\\Update.exe"
            os.startfile(hpath)

        elif "open command prompt" in query:
            os.system("start cmd")

        elif "ip address" in query:
            #ip = get('https://api.ipify.org').text
            #speak(f"your ip address is {ip}")
            
            hostname = socket.gethostname()   
            IPAddr = socket.gethostbyname(hostname)   
            speak("Your Computer Name is:" + hostname) 
            speak("Your Computer IP Address is:" + IPAddr)
        
        elif "open youtube" in query:
            webbrowser.open("www.youtube.com")

        elif "open facebook" in query:
            webbrowser.open("www.facebook.com")

        elif "open twitter" in query:
            webbrowser.open("www.twitter.com")

        elif "open stack overflow" in query:
            webbrowser.open("www.stackoverflow.com")
        
        elif "open browser" in query:
            speak("Sir! What should I search??")
            cm = takecommand().lower()
            webbrowser.open(f"{cm}")                     

        elif "send message" in query:
            kit.sendwhatmsg("+919311901559", "Hello Di!!",19,26)

        elif "play my favourite song" in query:
            kit.playonyt("Wake Up In the Sky")

        elif "the time" in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir the time is {strTime}")

        #elif "open camera" in query:
         #   def open_camera():
          #      sp.run('start microsoft.windows.camera:', shell=True)

        elif "hello max" in query:
            speak("Hello Sir! What can I do for you today?")
            
        elif "who made you" in query or "who created you" in query:
            speak("I have been created by Shivaansh")
        elif "calculate" in query:
             
            app_id = "Wolframalpha api id"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            print("The answer is " + answer)
            speak("The answer is " + answer)
            
        elif 'search' in query or 'play' in query:
             
            query = query.replace("search", "")
            query = query.replace("play", "")         
            webbrowser.open(query)
        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("Recycle Bin Recycled")
        elif 'weather' in query:
             
            # Google Open weather website
            # to get API of Open weather
            api_key = "Api key"
            base_url = "http://api.openweathermap.org / data / 2.5 / weather?"
            city_name = 'Delhi'
            complete_url = base_url + "appid =" + api_key + "&q =" + city_name
            response = requests.get(complete_url)
            x = response.json()
             
            if x["cod"] != "404":
                y = x["main"]
                current_temperature = y["temp"]
                current_pressure = y["pressure"]
                current_humidiy = y["humidity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description))
             
            else:
                speak(" City Not Found ")
            
            
        #elif "open calculator" in query:
            #def open_calculator():
               # sp.Popen(paths['calculator'])

        #elif "send WhatsApp message" in query:
           # def send_whatsapp_message(number, message):
                #kit.sendwhatmsg_instantly(f"+919311901559", "hello di")

        #elif "news" in query:
           # NEWS_API_KEY = config("NEWS_API_KEY")


            ##def get_latest_news():
              #  news_headlines = []
            ##res = requests.get(
               # f"https://newsapi.org/v2/top-headlines?country=in&apiKey={NEWS_API_KEY}&category=general").json()
            #articles = res["articles"]
            #for article in articles:
             #   news_headlines.append(article["title"])
            #return news_headlines[:5]

        

        #elif "send me an email" in query:
         #   try:

          #      speak("What should I say?")
           #     content = takecommand().lower()
            #    to = "shivanshsharma.ait@gmail.com"
             #   sendEmail(to,content)
              #  speak("Email Sent")

            #except Exception as e:
             #   print(e)
              #  speak("sorry sir, I am not able to send this email")

        elif "bye" in query:
            speak("Ok Sir! GoodBye")
            sys.exit()
#to find a joke
        elif "tell me a joke" in query:
            #joke = pyjokes.get_joke()
            #speak(joke)
            My_joke = pyjokes.get_joke(language="en", category="neutral")
            speak(My_joke)
        elif "another joke" in query:
            My_joke = pyjokes.get_joke(language="en", category="neutral")
            speak(My_joke)
        elif "that was not funny" in query:
            speak("I don't remember seeing your stand-up routine. Oh, that's right. You don't have one either. ")
        elif "not funny at all" in query:
                
            speak("looks like I'm gonna have to fire my writers again. They keep giving me lousy material to work with.")
            
        #elif "tell me a random joke" in query:
       #     def get_random_joke():
        #        headers = {
         #           'Accept': 'application/json'
          #      }
           # res = requests.get("https://icanhazdadjoke.com/", headers=headers).json()
            #return res["joke"]
            
        elif 'news' in query:
             
            try:
                jsonObj = urlopen('''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\''')
                data = json.load(jsonObj)
                i = 1
                 
                speak('here are some top news from the times of india')
                print('''=============== TIMES OF INDIA ============'''+ '\n')
                 
                for item in data['articles']:
                     
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:
                 
                print(str(e))

        elif  "shutdown the system" in query:
            os.system("shutdown /s /t 5")

        elif "restart the system" in query:
            os.system("shutdown /r /t 5")

        elif "sleep the system" in query:
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
