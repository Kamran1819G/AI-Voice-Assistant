from time import sleep
import webbrowser
import datetime
import wikipedia
import subprocess
import keyboard
import os
import pyttsx3
import speech_recognition as sr


def greetMe():
    hour = datetime.datetime.now().hour
    if hour >= 0 and hour < 12:
        print("Good Morning Sir")
        Speak("Good Morning Sir")
    elif hour >= 12 and hour < 18:
        print("Good Afternoon Sir")
        Speak("Good Afternoon Sir")
    else:
        print("Good Evening Sir")
        Speak("Good Evening Sir")


def take_commands():
    r = sr.Recognizer()  # initializing speech_recognition
    with sr.Microphone() as source:  # opening physical microphone of computer
        print("")
        print("How can I help you?")
        Speak("How can I help you?")
        print('Listening...')
        r.pause_threshold = 0.8
        audio = r.listen(source)    # storing audio/sound to audio variable
        try:
            print("Recognizing...")
            # Recognizing audio using google api
            Query = r.recognize_google(audio)
            print("you said : '", Query, "'")
        except Exception as e:
            print(e)
            print("Say that again sir")
            Speak("Say that again sir")
            # returning none if there are errors
            return "None"
    # returning audio as text
    return Query

# creating Speak() function to giving Speaking power
# to our voice assistant


def Speak(audio):
    # initializing pyttsx3 module
    engine = pyttsx3.init()

    # Speaking rate Setting
    # getting details of current speaking rate
    rate = engine.getProperty('rate')
    engine.setProperty('rate', 175)     # setting up new voice rate

    # Volume Setting
    # getting to know current volume level (min=0 and max=1)
    volume = engine.getProperty('volume')
    # setting up volume level  between 0 and 1
    engine.setProperty('volume', 1.0)

    # Voice Setting
    voices = engine.getProperty('voices')
    # 0 for male voice and 1 for female voice
    engine.setProperty('voice', voices[0].id)

    # anything we pass inside engine.say(), will be spoken by our voice assistant
    engine.say(audio)
    engine.runAndWait()


# Driver Code
if __name__ == '__main__':
    greetMe()
    # using while loop to communicate infinitely
    while True:
        command = take_commands().lower()
        if "exit" in command or "quit" in command or "shut up" in command or "close" in command or "bye" in command or "goodbye" in command:
            print("Sure Sir, as your wish")
            Speak("Sure Sir, as your wish")
            break
        if "about yourself" in command or "who are you" in command or "your name" in command:
            print("I am your personal assistant, I am created by you")
            Speak("I am your personal assistant, I am created by you")
        if "search" in command:
            command = command.replace("search", "")
            webbrowser.open_new_tab("https://www.google.com/search?q="+command)
        if "who is" in command:
            Speak('Searching Wikipedia...')
            command = command.replace("who is", "")
            results = wikipedia.summary(command, sentences=3)
            print("According to Wikipedia")
            Speak("According to Wikipedia")
            print(results)
            Speak(results)
        if "open google" in command:
            webbrowser.open_new_tab("https://www.google.com")
            print("Openning Google")
            Speak("Openning Google")
        if "open youtube" in command:
            webbrowser.open_new_tab("https://www.youtube.com")
            print("Openning Youtube")
            Speak("Openning Youtube")
        if "open map" in command:
            webbrowser.open_new_tab("https://www.google.com/maps")
            print("Openning Google Maps")
            Speak("Openning Google Maps")
        if "open gmail" in command:
            webbrowser.open_new_tab("gmail.com")
            print("Openning Google Mail")
            Speak("Openning Google Mail")
        if "open insta" in command or "open instagram" in command:
            webbrowser.open_new_tab("https://www.instagram.com/")
            print("Opening Instagram")
            Speak("Opening Instagram")
        if "time" in command:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            print(f"Sir, the time is {strTime}")
            Speak(f"Sir, the time is {strTime}")
        if "open terminal" in command or "open cmd" in command:
            subprocess.call('cmd.exe')
            print("Opening Terminal")
            Speak("Opening Terminal")
        if "open word" in command or "open microsoft word" in command:
            path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
            try:
                subprocess.call(path)
                print("Opening Microsoft Word")
                Speak("Opening Microsoft Word")
            except:
                print("Sorry, I am not able to open Microsoft Word")
                Speak("Sorry, I am not able to open Microsoft Word")
        if "open powerpoint" in command or "open microsoft powerpoint" in command:
            path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"
            try:
                subprocess.call(path)
                print("Opening Microsoft PowerPoint")
                Speak("Opening Microsoft PowerPoint")
            except:
                print("Sorry, I am not able to open Microsoft PowerPoint")
                Speak("Sorry, I am not able to open Microsoft PowerPoint")
        if "open excel" in command or "open microsoft excel" in command:
            path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
            try:
                subprocess.call(path)
                print("Opening Microsoft Excel")
                Speak("Opening Microsoft Excel")
            except:
                print("Sorry, I am not able to open Microsoft Excel")
                Speak("Sorry, I am not able to open Microsoft Excel")
        if "minimize all" in command or "minimise all" in command:
            keyboard.press_and_release('win + m')
            print("Minimizing all windows")
            Speak("Minimizing all windows")
        if "maximize all" in command or "maximise all" in command:
            keyboard.press_and_release('win + shift + m')
            print("Maximizing all windows")
            Speak("Maximizing all windows")
        if "take screenshot" in command or "take screenshot" in command or "screenshot" in command:
            keyboard.press_and_release('alt + tab')
            keyboard.press_and_release('win + shift + s')
            print("Taking screenshot")
            Speak("Taking screenshot")
        if "lock my pc" in command or "lock my computer" in command or "lock my device" in command:
            print("Locking your PC")
            Speak("Locking your PC")
            os.system("rundll32.exe user32.dll, LockWorkStation")
            break
        if "shutdown my pc" in command or "shutdown my computer" in command or "shutdown my device" in command:
            print("Shutting down your PC")
            Speak("Shutting down your PC")
            os.system("shutdown /s /t 1")
            break
