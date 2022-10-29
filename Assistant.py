import webbrowser
import datetime
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
        print('Listening')
        r.pause_threshold = 0.8
        audio = r.listen(source)    # storing audio/sound to audio variable
        try:
            print("Recognizing")
            # Recognizing audio using google api
            Query = r.recognize_google(audio)
            print("Your Command : '", Query, "'")
        except Exception as e:
            print(e)
            print("Say that again sir")
            Speak("Say that again sir")
            # returning none if there are errors
            return "None"
    # returning audio as text
    return Query

# Speak() function to giving Speaking power to our voice assistant


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
        if "exit" in command or "quit" in command or "shut up" in command:
            print("Sure Sir, as your wish")
            Speak("Sure Sir, as your wish")
            break
        if "about yourself" in command:
            print("I am your personal assistant, I am created by you")
            Speak("I am your personal assistant, I am created by you")
        if "open insta" in command or "open instagram" in command:
            webbrowser.open_new_tab("https://www.instagram.com/")
            print("Opening instagram")
            Speak("Opening instagram")
        if "open google" in command:
            webbrowser.open_new_tab("https://www.google.com")
            print("Google chrome is open now")
            Speak("Google chrome is open now")
        if "search" in command:
            command = command.replace("search", "")
            webbrowser.open_new_tab(command)
        if "open gmail" in command:
            webbrowser.open_new_tab("gmail.com")
            print("Google Mail open now")
            Speak("Google Mail open now")
        if "time" in command:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            print(f"Sir, the time is {strTime}")
            Speak(f"Sir, the time is {strTime}")
