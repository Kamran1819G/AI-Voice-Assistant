import os
import webbrowser
import subprocess
import datetime

try:
    import speech_recognition as sr
    import pyaudio
    import wikipedia
    import keyboard
    import pyttsx3
    import pyjokes
except Exception as e:
    try:
        print("Some modules are missing, Installing all require modules")
        os.system('pip install -r requirements.txt')
    except Exception as e:
        os.system('python -u assistant.py')


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

# def wake_up():
#    r = sr.Recognizer()
#    with sr.Microphone() as source:
#    audio = r.listen(source)
#       try:
#           Query = r.recognize_google(audio, language='en-in')
#       except Exception as e:
#           return "None"
#    return Query


def take_commands():
    r = sr.Recognizer()  # initializing speech_recognition
    with sr.Microphone() as source:  # opening physical microphone of computer
        print("")
        print("Hi, How can I help you?")
        Speak("Hi, How can I help you?")
        print('Listening...')
        # seconds of non-speaking audio before a phrase is considered complete
        r.pause_threshold = 0.6
        audio = r.listen(source)    # storing audio/sound to audio variable
        try:
            print("Recognizing...")
            # Recognizing audio using google api
            Query = r.recognize_google(audio, language='en-in')
            print("you said : '", Query, "'")
        except Exception as e:
            print(e)
            print("Please Say that again Sir")
            Speak("Please Say that again Sir")
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
    #    WAKE = "ok binod"
    # using while loop to communicate infinitely
    while True:
        #        print("Wake me up by saying 'ok siri'")
        #        command = wake_up().lower()

        #        if command.count(WAKE) > 0:
        command = take_commands().lower()

        # General commands
        if "who is siri" in command:
            print("Siri is a virtual assistant developed by Apple Inc.")
            Speak("Siri is a virtual assistant developed by Apple Inc.")
        if "who is google assistant" in command:
            print("Google Assistant is a virtual assistant developed by Google.")
            Speak("Google Assistant is a virtual assistant developed by Google.")
        if "who is alexa" in command:
            print("Alexa is a virtual assistant developed by Amazon.")
            Speak("Alexa is a virtual assistant developed by Amazon.")
        if "who is jarvis" in command:
            print("Jarvis is a virtual assistant developed by Tony Stark.")
            Speak("Jarvis is a virtual assistant developed by Tony Stark.")
        if "who is cortana" in command:
            print("Cortana is a virtual assistant developed by Microsoft.")
            Speak("Cortana is a virtual assistant developed by Microsoft.")
        if "about yourself" in command or "who are you" in command or "your name" in command:
            print("I am your personal assistant, I am created by you")
            Speak("I am your personal assistant, I am created by you")
        if "joke" in command:
            joke = pyjokes.get_joke()
            print(joke)
            Speak(joke)
        if "time" in command:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            print(f"Sir, the time is {strTime}")
            Speak(f"Sir, the time is {strTime}")
        if "exit" in command or "quit" in command or "shut up" in command or "bye" in command or "goodbye" in command:
            print("Sure Sir, as your wish")
            Speak("Sure Sir, as your wish")
            break

        # For Opening Web Applications
        if "search" in command:
            command = command.replace("search", "")
            webbrowser.open_new_tab(
                "https://www.google.com/search?q="+command)
        if "who is" in command or "wikipedia" in command:
            Speak('Searching Wikipedia...')
            command = command.replace("who is", "") and command.replace(
                "wikipedia", "") and command.replace("from wikipedia", "")
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

        # For Opening User Applications
        if "open word" in command or "open microsoft word" in command:
            path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
            try:
                print("Opening Microsoft Word")
                Speak("Opening Microsoft Word")
                subprocess.call(path)
            except:
                print("Sorry, I am not able to open Microsoft Word")
                Speak("Sorry, I am not able to open Microsoft Word")
        if "open powerpoint" in command or "open microsoft powerpoint" in command:
            path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"
            try:
                print("Opening Microsoft PowerPoint")
                Speak("Opening Microsoft PowerPoint")
                subprocess.call(path)
            except:
                print("Sorry, I am not able to open Microsoft PowerPoint")
                Speak("Sorry, I am not able to open Microsoft PowerPoint")
        if "open excel" in command or "open microsoft excel" in command:
            path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
            try:
                print("Opening Microsoft Excel")
                Speak("Opening Microsoft Excel")
                subprocess.call(path)
            except:
                print("Sorry, I am not able to open Microsoft Excel")
                Speak("Sorry, I am not able to open Microsoft Excel")
        if "open code" in command or "open visual studio code" in command:
            path = "C:\\Users\\"+os.getlogin()+"\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            try:
                print("Opening Visual Studio Code")
                Speak("Opening Visual Studio Code")
                subprocess.call(path)
            except:
                print("Sorry, I am not able to open Visual Studio Code")
                Speak("Sorry, I am not able to open Visual Studio Code")
        if "open eclipse" in command:
            path = "C:\\Users\\"+os.getlogin()+"\\eclipse\\java-2022-06\\eclipse\\eclipse.exe"
            try:
                print("Opening Eclipse")
                Speak("Opening Eclipse")
                subprocess.call(path)
            except:
                print("Sorry, I am not able to open Eclipse")
                Speak("Sorry, I am not able to open Eclipse")
        if "open figma" in command:
            path = "C:\\Users\\"+os.getlogin()+"\\AppData\\Local\\Figma\\Figma.exe"
            try:
                print("Opening Figma")
                Speak("Opening Figma")
                subprocess.call(path)
            except:
                print("Sorry, I am not able to open Figma")
                Speak("Sorry, I am not able to open Figma")
        if "open canva" in command:
            path = "C:\\Users\\"+os.getlogin()+"\\AppData\\Local\\Programs\\Canva\\Canva.exe"
            try:
                print("Opening Canva")
                Speak("Opening Canva")
                subprocess.call(path)
            except:
                print("Sorry, I am not able to open Canva")
                Speak("Sorry, I am not able to open Canva")
        if "open discord" in command:
            path = "C:\\Users\\"+os.getlogin() + \
                "\\AppData\\Local\\Discord\\Update.exe --processStart Discord.exe"
            try:
                print("Opening Discord")
                Speak("Opening Discord")
                subprocess.call(path)
            except:
                print("Sorry, I am not able to open Discord")
                Speak("Sorry, I am not able to open Discord")

        # For Openning System Applications
        if "open file explorer" in command:
            print("Opening File Explorer")
            Speak("Opening File Explorer")
            subprocess.Popen("C:\\Windows\\explorer.exe")
        if "open camera" in command:
            print("Opening Camera")
            Speak("Opening Camera")
            os.system("start microsoft.windows.camera:")
        if "open screen keyboard" in command or "open on screen keyboard" in command:
            print("Opening on screen keyboard")
            Speak("Opening on screen keyboard")
            os.system("osk")
        if "open calculator" in command:
            print("Opening Calculator")
            Speak("Opening Calculator")
            os.system("calc")
        if "open notepad" in command:
            print("Opening Notepad")
            Speak("Opening Notepad")
            os.system("notepad")
        if "open terminal" in command or "open cmd" in command:
            print("Opening Terminal")
            Speak("Opening Terminal")
            subprocess.call('cmd.exe')
        if "open control panel" in command:
            print("Opening Control Panel")
            Speak("Opening Control Panel")
            os.system("control panel")
        if "open task manager" in command:
            print("Opening Task Manager")
            Speak("Opening Task Manager")
            os.system("taskmgr")
        if "open environment variables" in command:
            print("Opening Environment Variables")
            Speak("Opening Environment Variables")
            os.system("rundll32 sysdm.cpl,EditEnvironmentVariables")
        if "open system properties" in command or "open system information" in command or "show system information" in command or "show system properties" in command:
            print("Opening System Properties")
            Speak("Opening System Properties")
            os.system("sysdm.cpl")
        if "open screen snipping" in command or "open screen snip" in command:
            print("Opening Screen Snipping")
            Speak("Opening Screen Snipping")
            os.system("snippingtool")
        if "disk cleanup" in command:
            print("Openning disk cleanup")
            Speak("Openning disk cleanup")
            os.system("start cleanmgr")
        if "open disk defragmenter" in command:
            print("Openning disk defragmenter")
            Speak("Openning disk defragmenter")
            os.system("start dfrgui")
        if "open disk management" in command:
            print("Openning disk management")
            Speak("Openning disk management")
            os.system("start diskmgmt.msc")
        if "open device manager" in command:
            print("Openning device manager")
            Speak("Openning device manager")
            os.system("start devmgmt.msc")
        if "open microsoft store" in command:
            print("Openning microsoft store")
            Speak("Openning microsoft store")
            os.system("start ms-windows-store:")

        # For Manuplation of Window size
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

        # For System Lock,Sleep,Restart,Shutdown
        if "lock my pc" in command or "lock my computer" in command or "lock my device" in command:
            print("Locking your PC")
            Speak("Locking your PC")
            os.system("rundll32.exe user32.dll, LockWorkStation")
            break
        if "sleep" in command:
            print("Putting your PC to sleep")
            Speak("Putting your PC to sleep")
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
            break
        if "shutdown" in command:
            print("Shutting down your PC")
            Speak("Shutting down your PC")
            os.system("shutdown /s /t 1")
            break
        if "restart my pc" in command or "restart my computer" in command or "restart my device" in command:
            print("Restarting your PC")
            Speak("Restarting your PC")
            os.system("shutdown /r /t 1")
            break
        if "hibernate my pc" in command or "hibernate my computer" in command or "hibernate my device" in command:
            print("Hibernating your PC")
            Speak("Hibernating your PC")
            os.system("shutdown /h")
            break

        # For Openning System Settings
        if "turn on bluetooth" in command or "turn off bluetooth" in command or "open bluetooth settings" in command:
            try:
                print("Openning bluetooth settings")
                Speak("Openning bluetooth settings")
                os.system("start ms-settings:bluetooth")
            except:
                print("Sorry, I am not able to open bluetooth settings")
                Speak("Sorry, I am not able to open bluetooth settings")
        if "turn on wi-fi" in command or "turn off wi-fi" in command or "open wi-fi settings" in command:
            try:
                print("Openning Wi-Fi settings")
                Speak("Openning Wi-Fi settings")
                os.system("start ms-settings:network-wifi")
            except:
                print("Sorry, I am not able to open Wi-Fi settings")
                Speak("Sorry, I am not able to open Wi-Fi settings")
        if "open display settings" in command:
            try:
                print("Openning display settings")
                Speak("Openning display settings")
                os.system("start ms-settings:display")
            except:
                print("Sorry, I am not able to open display settings")
                Speak("Sorry, I am not able to open display settings")
        if "open sound settings" in command:
            try:
                print("Openning sound settings")
                Speak("Openning sound settings")
                os.system("start ms-settings:sound")
            except:
                print("Sorry, I am not able to open sound settings")
                Speak("Sorry, I am not able to open sound settings")
        if "turn on hotspot" in command or "turn off hotspot" in command or "open hotspot settings" in command:
            try:
                print("Openning hotspot settings")
                Speak("Openning hotspot settings")
                os.system("start ms-settings:network-mobilehotspot")
            except:
                print("Sorry, I am not able to open hotspot settings")
                Speak("Sorry, I am not able to open hotspot settings")
        if "turn on airplane mode" in command or "turn off airplane mode" in command or "open airplane mode settings" in command:
            try:
                print("Openning airplane mode settings")
                Speak("Openning airplane mode settings")
                os.system("start ms-settings:network-airplanemode")
            except:
                print("Sorry, I am not able to open airplane mode settings")
                Speak("Sorry, I am not able to open airplane mode settings")
        if "open battery settings" in command:
            try:
                print("Openning battery settings")
                Speak("Openning battery settings")
                os.system("start ms-settings:batterysaver")
            except:
                print("Sorry, I am not able to open battery settings")
                Speak("Sorry, I am not able to open battery settings")
        if "open location settings" in command:
            try:
                print("Openning location settings")
                Speak("Openning location settings")
                os.system("start ms-settings:privacy-location")
            except:
                print("Sorry, I am not able to open location settings")
                Speak("Sorry, I am not able to open location settings")
        if "turn on night light" in command or "turn off night light" in command or "open night light settings" in command:
            try:
                print("Openning night light settings")
                Speak("Openning night light settings")
                os.system("start ms-settings:nightlight")
            except:
                print("Sorry, I am not able to open night light settings")
                Speak("Sorry, I am not able to open night light settings")
        if "open theme settings" in command or "change theme" in command:
            try:
                print("Openning theme settings")
                Speak("Openning theme settings")
                os.system("start ms-settings:personalization")
            except:
                print("Sorry, I am not able to open theme settings")
                Speak("Sorry, I am not able to open theme settings")
        if "search for updates" in command or "check for updates" in command:
            try:
                print("Searching for updates")
                Speak("Searching for updates")
                os.system("start ms-settings:windowsupdate")
            except:
                print("Sorry, I am not able to search for updates")
                Speak("Sorry, I am not able to search for updates")
