import os
import sys
import smtplib
import webbrowser
import requests
import wikipedia
import pywhatkit as kit
import pyjokes
import cv2
import pyautogui
import platform
import psutil
import datetime
import pyttsx3
import speech_recognition as sr
from googletrans import Translator
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from tkinter import Tk, Label, Button, simpledialog
import threading
import schedule
import time
import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer

nltk.download('stopwords')
nltk.download('punkt')
nltk.download('wordnet')

# Initialize Text-to-Speech engine
engine = pyttsx3.init()
engine.setProperty('rate', 150)
engine.setProperty('volume', 1.0)

recognizer = sr.Recognizer()

# Initialize NLTK components
lemmatizer = WordNetLemmatizer()
stop_words = set(stopwords.words('english'))

# Function to speak
def speak(audio):
    engine.say(audio)
    engine.runAndWait()

# Function to recognize voice command
def take_command(initial_wake_up=False):
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)

        try:
            query = recognizer.recognize_google(audio, language="en_in").lower()
            if initial_wake_up and "hello" in query:
                speak("Yes, how can I help you?")
                audio = recognizer.listen(source)
                command = recognizer.recognize_google(audio, language="en_in").lower()
                return command
            elif not initial_wake_up:
                return query
        except sr.UnknownValueError:
            speak("Sorry, I didn't catch that. Please try again.")
        except sr.RequestError as e:
            speak(f"Could not request results from Google Speech Recognition service; {e}")
        return None

# Function to set reminder
def set_reminder():
    reminder_time = simpledialog.askstring("Reminder Time", "Specify reminder time (hh:mm):")
    reminder_message = simpledialog.askstring("Reminder Message", "Enter reminder message:")
    schedule.every().day.at(reminder_time).do(lambda: speak(reminder_message))

# Function to run schedule in a separate thread
def run_schedule():
    while True:
        schedule.run_pending()
        time.sleep(1)

threading.Thread(target=run_schedule).start()

# Function to handle translation request
def handle_translate_request(text_to_translate, target_language):
    translator = Translator()
    translated_text = translator.translate(text_to_translate, dest=target_language).text
    speak(f"Translated text: {translated_text}")

def process_query(query):
    tokens = word_tokenize(query)
    tokens = [lemmatizer.lemmatize(token) for token in tokens if token.lower() not in stop_words]
    return ' '.join(tokens)

# Function to get weather information
def get_weather(city):
    api_key = os.getenv('OPENWEATHER_API_KEY')
    url = f'http://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric'
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        weather_desc = data['weather'][0]['description']
        temp = data['main']['temp']
        speak(f"The weather in {city} is {weather_desc} with a temperature of {temp} degrees Celsius.")
    else:
        speak(f"Sorry, I couldn't fetch the weather information for {city}.")

def sleep_system():
    os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")

def open_task_manager():
    os.system("taskmgr")

def open_control_panel():
    os.system("control")

def download_from_internet(url, save_path):
    try:
        response = requests.get(url)
        with open(save_path, 'wb') as f:
            f.write(response.content)
        speak("Download completed.")
    except Exception as e:
        speak(f"Failed to download the file: {str(e)}")

def get_system_info():
    system_info = {
        'System': platform.system(),
        'Node Name': platform.node(),
        'Release': platform.release(),
        'Version': platform.version(),
        'Machine': platform.machine(),
        'Processor': platform.processor(),
        'CPU Usage': psutil.cpu_percent(interval=1),
        'Memory': psutil.virtual_memory().percent
    }
    return system_info

def fetch_data(query):
    url = f"https://en.wikipedia.org/api/rest_v1/page/summary/{query}"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        return data.get('extract', '')
    else:
        return None

valid_credentials = {
    "username": "saugat",
}

def authenticate_user():
    while True:
        username = simpledialog.askstring("Login", "Enter your username:")
        if username in valid_credentials:
            return True, username
        else:
            continue

def authorize_user(user):
    return user in valid_credentials

def main_loop():
    authenticated, username = authenticate_user()
    while True:
        query = take_command()
        if query and authorize_user(username):
            if 'tell me about' in query:
                topic = query.replace('tell me about', '').strip()
                speak(f"Fetching information about {topic}...")
                try:
                    result = fetch_data(topic)
                    if result:
                        speak(result)
                    else:
                        speak(f"I couldn't find information about {topic}.")
                except Exception as e:
                    speak(f"An error occurred: {str(e)}")
            elif 'open command prompt' in query:
                speak("Opening command prompt")
                os.system("start cmd")
            elif "set a reminder" in query:
                speak("Setting a reminder")
                set_reminder()
            elif 'shutdown' in query:
                speak("Shutting down the system")
                os.system("shutdown /s")
            elif 'restart' in query:
                speak("Restarting the system")
                os.system("shutdown /r")
            elif 'weather in' in query:
                city = query.split("weather in ")[1]
                get_weather(city)
            elif 'logout' in query:
                speak("Logging out")
                os.system("shutdown /l")
            elif 'sleep system' in query:
                speak("Putting the system to sleep")
                sleep_system()
            elif 'open task manager' in query:
                speak("Opening task manager")
                open_task_manager()
            elif 'open control panel' in query:
                speak("Opening control panel")
                open_control_panel()
            elif 'empty recycle bin' in query:
                speak("Recycling Bin emptied")
                os.system("powershell.exe -Command Clear-RecycleBin -Force")
            elif 'switch the window' in query:
                pyautogui.keyDown("alt")
                pyautogui.press("tab")
                pyautogui.sleep(1)
                pyautogui.keyUp("alt")
            elif 'sleep' in query:
                speak("Okay, I am going to sleep. You can call me anytime.")
                sys.exit()
            elif 'open camera' in query:
                speak("Opening camera")
                cap = cv2.VideoCapture(0)
                while True:
                    ret, img = cap.read()
                    cv2.imshow('webcam', img)
                    if cv2.waitKey(50) == 27:
                        break
                cap.release()
                cv2.destroyAllWindows()
            elif 'wikipedia' in query:
                speak("Searching Wikipedia")
                query = query.replace("wikipedia", "")
                result = wikipedia.summary(query, sentences=2)
                speak("According to Wikipedia")
                speak(result)
            elif 'search' in query or 'play' in query:
                query = query.replace("search", "").replace("play", "")
                search_query = take_command()
                webbrowser.open(search_query)
            elif 'open google' in query:
                speak("Opening Google")
                cm = take_command().lower()
                webbrowser.open(f"https://www.google.com/search?q={cm}")
            elif 'open facebook' in query:
                speak("Opening Facebook")
                webbrowser.open("https://www.facebook.com")
            elif 'open youtube' in query:
                speak("Opening YouTube")
                webbrowser.open("https://www.youtube.com")
            elif 'open stackoverflow' in query:
                speak("Opening Stack Overflow")
                webbrowser.open("https://www.stackoverflow.com")
            elif 'download' in query:
                speak("Please provide the URL of the file you want to download.")
                url = take_command()
                speak("Where do you want to save the file?")
                save_path = take_command()
                download_from_internet(url, save_path)
            elif 'song on youtube' in query:
                speak("What do you want to listen?")
                song = take_command()
                kit.playonyt(song)
            elif 'joke' in query:
                joke = pyjokes.get_joke()
                speak(joke)
            elif 'open code' in query:
                speak("Opening Visual Studio Code")
                path = r"C:\Users\sauga\Downloads\Programs\VSCodeUserSetup-x64-1.68.1.exe"
                os.startfile(path)
            elif 'open chrome' in query:
                speak("Opening Chrome")
                path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
                os.startfile(path)
            elif "send email to saugat" in query:
                speak("What should I say?")
                message = take_command()
                if "send file" in message:
                    send_email_with_attachment()
                else:
                    send_email(message)
            elif 'hello' in query:
                speak("Hi, how can I help you?")
            elif 'hi' in query:
                speak("Hello, how can I assist you?")
            elif 'translate' in query:
                text_to_translate = query.replace('translate', '').strip()
                target_language = 'en'  # Set the default target language
                handle_translate_request(text_to_translate, target_language)
            elif 'close wikipedia' in query:
                speak("Closing Wikipedia")
                os.system("TASKKILL /F /T /IM msedge.exe")
            elif 'time' in query:
                current_time = datetime.datetime.now().strftime("%H:%M")
                speak(f"The current time is {current_time}")
            else:
                speak(f"You said: {query}")
        else:
            speak("You are not authorized to perform this action.")

def send_email(message):
    email = os.getenv('EMAIL')
    password = os.getenv('EMAIL_PASSWORD')
    send_to_email = os.getenv('RECIPIENT_EMAIL')
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(email, password)
    server.sendmail(email, send_to_email, message)
    server.quit()
    speak("Email has been sent to saugat")

def send_email_with_attachment():
    email = os.getenv('EMAIL')
    password = os.getenv('EMAIL_PASSWORD')
    send_to_email = os.getenv('RECIPIENT_EMAIL')

    subject = simpledialog.askstring("Email Subject", "Enter the subject:")
    body = simpledialog.askstring("Email Body", "Enter the message:")
    file_location = simpledialog.askstring("File Location", "Enter the file path to attach:")

    msg = MIMEMultipart()
    msg['From'] = email
    msg['To'] = send_to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    file_name = os.path.basename(file_location)
    attachment = open(file_location, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename= {file_name}")

    msg.attach(part)

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(email, password)
    server.sendmail(email, send_to_email, msg.as_string())
    server.quit()
    speak("Email has been sent to saugat")

# Create the GUI
root = Tk()
root.title("Voice Assistant")
root.geometry("400x300")

# Add a label
label = Label(root, text="Press the button to start the Voice Assistant", font=("Arial", 14))
label.pack(pady=20)

# Add a start button
start_button = Button(root, text="Start", command=main_loop, font=("Arial", 12))
start_button.pack(pady=10)

# Run the GUI
root.mainloop()
