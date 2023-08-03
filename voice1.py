import smtplib
import wolframalpha
import ecapture
import pyjokes
import subprocess
import json
import random
import winshell
import ctypes
import requests
import shutil
from twilio.rest import Client
import time
import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import win32com.client
import openai
from config import apikey
from config import gmailPass
from config import acc_sid
from config import auth_tok
from config import ph_no
from config import my_ph_no

# openai.api_key = apikey

# Pyjokes:- Pyjokes is used for collection Python Jokes over the Internet
# Twilio:- Twilio is used for making call and messages.
# Requests: Requests is used for making GET and POST requests.
# BeautifulSoup: Beautiful Soup is a library that makes it easy to scrape information from web pages.

assistant_name = "Jarvis"
engine = pyttsx3.init('sapi5')
voice = engine.getProperty('voices')


chatStr = ""


def chat(userQuery):
    openai.api_key = apikey
    global chatStr
    print(chatStr)
    # text = f"OpenAI response for Prompt: {userQuery} \n***************************\n\n"
    chatStr += f"Hardi: {userQuery} Jarvis: "
    response = openai.Completion.create(
        model="gpt-3.5-turbo",
        messages=[
            {
                "role": "system",
                "content": "Write a letter to my boss asking for 3 days leave."
            },
            {
                "role": "user",
                "content": ""
            }
        ],
        prompt=chatStr,
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    # print(response["choices"][0]["content"])
    chatStr += response["choices"][0]["content"]
    return response["choices"][0]["content"]


def ai(prompt):
    openai.api_key = apikey
    text = f"OpenAI response for Prompt: {prompt} \n***************************\n\n"

    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[
            {
                "role": "system",
                "content": "Write a letter to my boss asking for 3 days leave."
            },
            {
                "role": "user",
                "content": ""
            }
        ],
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    print(response["choices"][0]["content"])
    text += response["choices"][0]["content"]

    if not os.path.exists("OpenAI"):
        os.mkdir("OpenAI")

    with open(f"OpenAI/{''.join(prompt.split('intelligence')[1:]).strip()}.txt", "w") as f:
        f.write(text)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        speak("Good Morning Hardi!!")
    elif 12 < hour < 18:
        speak("Good Afternoon Hardi!!")
    else:
        speak("Good Evening Hardi!!")

    speak("I am Jarvis. How can I help you?")


def username():
    speak("What should I call you?")
    uname = takeCommand()
    speak(f"Welcome {uname}")
    columns = shutil.get_terminal_size().columns
    print("###################################".center(columns))
    print(f"Welcome {uname}".center(columns))
    print("###################################".center(columns))
    speak("I am Jarvis. How can I help you?")


def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    #     encrypt your password
    server.login('joisarharditemp@gmail.com', gmailPass)
    server.sendmail('hardijoisar152@gmail.com', to, content)
    server.close()


def takeCommand():
    """Takes command from the user through microphone."""
    print("Speak something")
    r = sr.Recognizer()
    with sr.Microphone() as source:

        r.adjust_for_ambient_noise(source)
        # r.pause_threshold = 1
        print("Listening!!")
        audio = r.listen(source)

        try:
            print("Recognizing !!")
            text = r.recognize_google(audio, language='en-in')
            print(f"User said : {text}\n")

        except Exception as e:
            print(e)
            print("Sorry, could not understand!!")
            return "An error occurred !!"

    return text


if __name__ == '__main__':
    # speak("Hardi is coding Jarvis")
    username()
    wishMe()
    while True:
        query = takeCommand().lower()
        # query=input("Enter your command: ")

        if 'wikipedia' in query:
            speak("Searching Wikipedia")
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=1)
            speak("According to wikipedia, ")
            speak(results)

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:
            webbrowser.open("stackoverflow.com")

        elif "play music" in query:
            m_dir = "E:\\Hardi\\songs"
            songs = os.listdir(m_dir)
            print(songs)
            # random songs
            os.startfile(os.path.join(m_dir, songs[0]))

        elif 'the time' in query:
            timeNow = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"The time is : {timeNow}")

        # send mail
        # email using dictionary
        elif 'send me email' in query:
            try:
                speak("What should I mail?")
                mailContent = takeCommand()
                to = "hardijoisar152@gmail.com"
                sendEmail(to, mailContent)
                speak("Email has been sent successfully")
            except Exception as e:
                print(e)
                speak("Sorry, couldn't send email at the moment")

        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")

        elif "your name" in query or "your name" in query:
            speak("My friends call me")
            speak(assistant_name)
            print("My friends call me", assistant_name)

        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit()

        elif "who made you" in query or "who created you" in query:
            speak("I have been created by Hardi.")

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        elif 'search' in query or 'play' in query:
            query = query.replace("search", "")
            query = query.replace("play", "")
            webbrowser.open(query)

        elif "who am i" in query:
            speak("I think you are a human.")

        elif "why you came to world" in query:
            speak("It's a secret")

        elif 'powerpoint presentation' in query:
            speak("opening Power Point presentation")
            ppt_path = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint"
            os.startfile(ppt_path)

        # elif 'change background' in query:
        #     ctypes.windll.user32.SystemParametersInfoW(20, 0, "Location of wallpaper", 0)
        #     speak("Background changed successfully")

        elif 'news' in query:
            news = webbrowser.open_new_tab("https://timesofindia.indiatimes.com/home/headlines")
            speak('Here are some headlines from the Times of India,Happy reading')

        elif 'lock window' in query:
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in query:
            speak("Hold On a Sec ! Your system is on its way to shut down")
            subprocess.call('shutdown / p /f')

        elif 'empty recycle bin' in query:
            speak("As you wish ma'am..")
            winshell.recycle_bin().empty(confirm=False, sound=False, show_progress=True)
            speak("Recycle bin is empty now!!")

        elif "don't listen" in query or "stop listening" in query:
            speak("for how many minutes you want to stop jarvis from listening commands")
            a = float(input("for how many minutes you want to stop jarvis from listening commands: "))
            print(f"Okay!!  Jarvis is going to sleep for {a} minutes")
            time.sleep(a * 60)

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.nl / maps / place/" + location + "")

        # elif "camera" in query or "take a photo" in query:
        #     ec.capture(0, "Jarvis Camera ", "img.png")

        elif "restart" in query:
            subprocess.call(['shutdown', '/r'])

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            subprocess.call(["shutdown", "/l", "/t", "10"])

        elif "write a note" in query:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('jarvis.txt', 'w')
            speak("Sir, Should i include date and time")
            dtt = takeCommand()
            if 'yes' in dtt or 'sure' in dtt:
                strTime = datetime.datetime.now().strftime("% H:% M:% S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)

        elif "show note" in query:
            speak("Showing Notes")
            file = open("jarvis.txt", "r")
            print(file.read())
            speak(file.read(6))

        elif 'send message' in query:

            # Your Account SID from twilio
            account_sid = acc_sid
            # Your Auth Token from twilio
            auth_token = auth_tok

            client = Client(account_sid, auth_token)
            try:
                message = client.messages.create(
                    to=ph_no,
                    from_=my_ph_no,
                    body="Hello from Python!")
                print("Message sent")
                print(message.sid)
            except Exception as e:
                speak("Sorry!! Couldn't Process request at the moment.")

        elif "Good Morning" in query:
            speak("A warm Good Morning to you too!!")

        elif "will you be my girlfriend" in query or "will you be my boyfriend" in query:
            speak("I'm not sure about, may be you should give me some time")

        elif "using artificial intelligence" in query:
            ai(query)

        elif "clear chat" in query:
            chatStr = ""
            print("Cleared chat successfully")

        else:
            chat(query)
