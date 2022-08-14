from jarvis_gui import Ui_MainWindow
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from requests import get
from bs4 import BeautifulSoup
from pywikihow import search_wikihow
import PyPDF2
import instaloader
import pyttsx3
import speech_recognition as sr
import datetime
import os
import cv2
import random
import requests
import wikipedia
import webbrowser
import pywhatkit as kit
import smtplib
import sys
import time
import pyjokes
import pyautogui
import operator
import psutil
import speedtest

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)
engine.setProperty('rate', 200)


def speak(audio):
    engine.say(audio)
    print(audio)
    engine.runAndWait()


def take_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source, timeout=5, phrase_time_limit=5)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"You said: {query}")

    except Exception as e:
        speak("Say that again please...")
        return "none"
    return query


def wish():
    hour = int(datetime.datetime.now().hour)

    if 0 <= hour < 12:
        speak("Good morning sir!")
    elif 12 <= hour < 18:
        speak("Good afternoon sir!")
    else:
        speak("Good evening sir!")
    speak("How can I help you?")


def send_email(recipient_email_id, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login('pradhanrajat2001@gmail.com', 'djrzhsjhxgstthod')
    server.sendmail('pradhanrajat2001@gmail.com', recipient_email_id, content)
    server.quit()


def send_attachment_email(recipient_email_id, subject, content, file_location):
    msg = MIMEMultipart()
    msg['From'] = "pradhanrajat2001@gmail.com"
    msg['To'] = recipient_email_id
    msg['Subject'] = subject

    msg.attach(MIMEText(content, 'plain'))

    file_name = os.path.basename(file_location)
    attachment = open(file_location, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition',
                    "attachment; file_name = % s" % file_name)

    msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login('pradhanrajat2001@gmail.com', 'djrzhsjhxgstthod')
    text = msg.as_string()
    server.sendmail('pradhanrajat2001@gmail.com', recipient_email_id, content)
    server.quit()
    speak("Email has been successfully sent!")


def news():
    main_url = 'https://newsapi.org/v2/top-headlines?sources=google-news-in&apiKey=a56cf5f0490f4ee7a8522f30883fed30'
    main_page = requests.get(main_url).json()
    articles = main_page["articles"]
    head = []
    day = ["first", "second", "third", "fourth", "fifth",
           "sixth", "seventh", "eighth", "ninth", "tenth"]
    for ar in articles:
        head.append(ar["title"])
    for i in range(len(day)):
        speak(f"Todays {day[i]} news is: {head[i]}")


def read_pdf():
    speak("Sir please enter the path of the file to read.")
    path = input("Enter PDF file path: ")
    pdf_file = open(path, 'rb')
    pdf_reader = PyPDF2.PdfFileReader(pdf_file)
    pages = pdf_reader.numPages
    speak(f"Total {pages} pages are there in this PDF.")
    speak("Sir please enter the page number to read.")
    page_no = int(input("Enter page number: ")) - 1
    page = pdf_reader.getPage(page_no)
    text = page.extractText()
    speak(text)


if __name__ == "__main__":
    wish()

    while True:
        query = take_command().lower()

        if "open notepad" in query:
            path = "C:\\Windows\\system32\\notepad.exe"
            os.startfile(path)

        elif "open command prompt" in query:
            os.system("start cmd")

        elif "open google chrome" in query:
            path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
            os.startfile(path)

        elif "open microsoft edge" in query:
            path = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
            os.startfile(path)

        elif "open adobe illustrator" in query:
            path = "D:\\Softwares\\Adobe Illustrator 2021\\Support Files\\Contents\\Windows\\Illustrator.exe"
            os.startfile(path)

        elif "open adobe photoshop" in query:
            path = "D:\\Softwares\\Adobe Photoshop 2021\\Photoshop.exe"
            os.startfile(path)

        elif "open eagle get" in query:
            path = "C:\\Program Files (x86)\\EagleGet\\EagleGet.exe"
            os.startfile(path)

        elif "open excel" in query:
            path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
            os.startfile(path)

        elif "open powerpoint" in query:
            path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"
            os.startfile(path)

        elif "open word" in query:
            path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
            os.startfile(path)

        elif "open onedrive" in query:
            path = "C:\\Program Files (x86)\\Microsoft OneDrive\\OneDrive.exe"
            os.startfile(path)

        elif "open nvidia geforce experience" in query:
            path = "C:\\Program Files\\NVIDIA Corporation\\NVIDIA GeForce Experience\\NVIDIA GeForce Experience.exe"
            os.startfile(path)

        elif "open vlc media player" in query:
            path = "C:\\Program Files\\VideoLAN\\VLC\\vlc.exe"
            os.startfile(path)

        elif "open visual studio code" in query:
            path = "D:\\Softwares\\Microsoft VS Code\\Code.exe"
            os.startfile(path)

        elif "open winrar" in query:
            path = "C:\\Program Files\\WinRAR\\WinRAR.exe"
            os.startfile(path)

        elif "open camera" in query:
            cap = cv2.VideoCapture(0)

            while True:
                check, img = cap.read()
                cv2.imshow('WebCam', img)
                fps = 30
                key = cv2.waitKey(int(1000/fps))
                if key == 27:
                    break

            cap.release()
            cv2.destroyAllWindows()

        # Play music
        elif "play music" in query:
            music_dir = "D:\\Music"
            songs = os.listdir(music_dir)
            rd = random.choice(songs)
            os.startfile(os.path.join(music_dir, rd))

        # Find IP address
        elif "ip address" in query:
            ip = get('https://api.ipify.org').text
            speak(f"Your IP address is {ip}")

        #Search in Wikipedia
        elif "wikipedia" in query:
            speak("Searching Wikipedia...")
            query = query.replace("wikipedia", "")
            result = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia,")
            speak(result)

        elif "open google" in query:
            speak("Sir, what should I search on Google?")
            search_query = take_command().lower()
            webbrowser.open(search_query)

        elif "open youtube" in query:
            webbrowser.open("youtube.com")

        elif "open facebook" in query:
            webbrowser.open("facebook.com")

        elif "open instagram" in query:
            webbrowser.open("instagram.com")

        elif "open amazon" in query:
            webbrowser.open("amazon.in")

        elif "open flipkart" in query:
            webbrowser.open("flipkart.com")

        elif "open stackoverflow" in query:
            webbrowser.open("stackoverflow.com")

        # Send WhatsApp message
        elif "send a message" in query:
            speak("Sir, to whom do you want to send message?")
            recipient = take_command().lower()
            speak("What is the message sir?")
            message = take_command().lower()
            current_hour = int(datetime.datetime.now().hour)
            current_minute = int(datetime.datetime.now().minute)

            if "aman" in recipient:
                kit.sendwhatmsg("+918770620410", message,
                                current_hour, current_minute + 1)

            elif "ajay" in recipient:
                kit.sendwhatmsg("+917697342045", message,
                                current_hour, current_minute + 1)

            elif "swarnima" in recipient:
                kit.sendwhatmsg("+916261333771", message,
                                current_hour, current_minute + 1)

            elif "ankit" in recipient:
                kit.sendwhatmsg("+917828488998", message,
                                current_hour, current_minute + 1)

            elif "aakash" in recipient:
                kit.sendwhatmsg("+917974219880", message,
                                current_hour, current_minute + 1)

            elif "ritvij" in recipient:
                kit.sendwhatmsg("+919770622632", message,
                                current_hour, current_minute + 1)

            elif "aditya" in recipient:
                kit.sendwhatmsg("+918349410342", message,
                                current_hour, current_minute + 1)

            elif "hemant" in recipient:
                kit.sendwhatmsg("+919340740239", message,
                                current_hour, current_minute + 1)

            elif "mom" in recipient:
                kit.sendwhatmsg("+917354376202", message,
                                current_hour, current_minute + 1)

        elif "play video on youtube" in query:
            speak("Sir, what you want me to play on YouTube?")
            search_query = take_command().lower()
            kit.playonyt(search_query)

        # Send email
        elif "send an email" in query:
            try:
                speak("Sir, to whom do you want to send email?")
                recipient = take_command().lower()
                speak("What subject do you want to put in the mail sir?")
                subject = take_command().lower()
                speak("What do you want to be written in the email sir?")
                content = take_command().lower()
                speak("Do you want to attach any file sir?")
                response = take_command().lower()

                if "yes" in response:
                    speak("Sir, please enter the path of the file you want to attach")
                    file_location = input("Enter path: ")
                    speak("Please wait sir, sending the email...")
                    send_attachment_email(
                        "pradhanrajat2001@gmail.com", subject, content, file_location)

                if "myself" in recipient:
                    send_email('pradhanrajat2001@gmail.com', content)

                elif "ajay" in recipient:
                    send_email('ajaykumarsahu0505@gmail.com', content)

                elif "ankit" in recipient:
                    send_email('aags1911@gmail.com', content)

                elif "aman" in recipient:
                    send_email('khandweaman@gmail.com', content)

                speak("Email has been successfully sent.")

            except Exception as e:
                print(e)
                speak("Sorry sir, I am not able to send the email.")

        elif "close notepad" in query:
            speak("Closing Notepad")
            os.system("taskkill /f /im notepad.exe")

        # Set an alarm
        elif "set alarm" in query:
            nn = int(datetime.datetime.now().hour)

            if nn == 22:
                music_dir = 'D:\\Music'
                songs = os.listdir(music_dir)
                os.startfile(os.path.join(music_dir, songs[0]))

        # Listen a joke
        elif "joke" in query:
            joke = pyjokes.get_joke()
            speak(joke)

        elif "shut down the system" in query:
            os.system("shutdown /s /t /5")

        elif "restart the system" in query:
            os.system("shutdown /r /t 5")

        elif "sleep the system" in query:
            os.system("rundll32.exe powrprof.dll, SetSuspendState 0, 1, 0")

        # Switch windows
        elif "switch window" in query:
            pyautogui.keyDown("alt")
            pyautogui.press("tab")
            time.sleep(1)
            pyautogui.keyUp("alt")

        # Listen to news
        elif "news" in query:
            speak("Please wait sir, fetching the latest news...")
            news()

        # Find geographical location based on IP address
        elif "where am i" in query or "where are we" in query:
            speak("Please wait sir, let me check...")
            try:
                ip_address = requests.get('https://api.ipify.org').text
                print(ip_address)
                url = 'https://get.geojs.io/v1/ip/geo/' + ip_address + '.json'
                geo_requests = requests.get(url)
                geo_data = geo_requests.json()
                city = geo_data['city']
                country = geo_data['country']
                speak(f"Sir, we are in {city} city of {country}")
            except Exception as e:
                speak("Sorry sir, I am not able to find our location.")

        # See Instagram profile of someone
        elif "instagram profile" in query or "profile on instagram" in query:
            speak("Sir please enter the username.")
            username = input("Enter username: ")
            webbrowser.open(f"instagram.com/{username}")
            speak(f"Sir here is the profile of the user {username}")
            time.sleep(5)
            speak("Sir, do you want to download the profile picture of this account?")
            response = take_command().lower()
            if "yes" in response:
                mod = instaloader.Instaloader()
                mod.download_profile(username, profile_pic_only=True)
                speak("Profile picture downloaded sir, saved to the main folder")
            else:
                pass

        # Take screenshots
        elif "screenshot" in query:
            speak("Sir please tell me the name of this screenshot file.")
            file_name = take_command().lower()
            speak(
                "Sir please hold the screen for few seconds, I am taking the screenshot.")
            time.sleep(3)
            image = pyautogui.screenshot()
            image.save(f"{file_name}.png")
            speak("Took the screenshot sir, saved to the main folder")

        # Read any PDF file
        elif "read pdf" in query:
            read_pdf()

        # Hide/unhide files
        elif "hide" in query or "unhide" in query or "visible" in query:
            speak(
                "Sir please tell me if you want to hide all files or make all files visible?")
            response = take_command().lower()
            if "hide" in response:
                os.system("attrib +h /s /d")
                speak("Sir, all the files in this folder are now hidden.")

            elif "visible" in response:
                os.system("attrib -h /s /d")
                speak(
                    "Sir, all files in this folder are now visible to everyone. I hope you are taking this step on your own risk")

            elif "leave it" in response or "leave for now" in response:
                speak("Okay sir!")

        # Perform basic mathematical calculations
        elif "calculation" in query or "calculate" in query:
            r = sr.Recognizer()
            with sr.Microphone() as source:
                speak("What you want me to calculate sir?")
                print("Listening...")
                r.adjust_for_ambient_noise(source)
                audio = r.listen(source)

            my_string = r.recognize_google(audio)
            print(my_string)

            def get_operator_fn(op):
                return {
                    "+": operator.add,
                    "-": operator.sub,
                    "x": operator.mul,
                    "divided": operator.__truediv__,
                }[op]

            def eval_binary_expr(op1, oper, op2):
                op1, op2 = int(op1), int(op2)
                return get_operator_fn(oper)(op1, op2)

            speak("Your result is:")
            speak(eval_binary_expr(*(my_string.split())))

        # Check temperature
        elif "temperature" in query:
            search = "temperature in Raigarh"
            url = f"https://www.google.co.in/search?q={search}"
            r = requests.get(url)
            data = BeautifulSoup(r.text, "html.parser")
            temp = data.find("div", class_="BNeawe").text
            speak(f"Current {search} is {temp}")

        # How to do anything
        elif "how to" in query:
            speak("Sir, tell me what you want to know?")
            how = take_command().lower()
            max_results = 1
            how_to = search_wikihow(how, max_results)
            assert len(how_to) == 1
            speak(how_to[0].summary)

        # Check battery percentage
        elif "how much power left" in query or "how much power we have" in query or "battery" in query:
            battery = psutil.sensors_battery()
            percentage = battery.percent
            speak(f"Sir we have {percentage}% power left")
            if percentage >= 81:
                speak("We have enough power to continue our work")
            elif percentage >= 41 and percentage <= 80:
                speak("You should look for a charging point")
            elif percentage >= 11 and percentage <= 40:
                speak("We don't have enough power, please connect to a charging point")
            elif percentage <= 10:
                speak(
                    "We have critically very low power, please connect to a charging point or the system is going to shutdown soon")

        # Check internet speed
        elif "internet speed" in query:
            speak("Please wait a few seconds...")
            st = speedtest.Speedtest()
            dl = st.download()
            up = st.upload()
            dl_mbps = round(dl/(8*1024*1024), 1)
            up_mbps = round(up/(8*1024*1024), 1)
            speak(
                f"Sir we have {dl_mbps} MB per second download speed and {up_mbps} MB per second upload speed")

        elif "no" in query:
            speak("Okay sir, have a good day!")
            sys.exit()

        speak("Sir, do you have any other work?")
