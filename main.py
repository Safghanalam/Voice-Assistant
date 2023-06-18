# Importing Needed Libraries and Modules
import os
import pyaudio
import speech_recognition as sr
import win32com.client
import webbrowser
import datetime
import wikipedia
import pyautogui

# setting Speaker voice
speaker = win32com.client.Dispatch("SAPI.SpVoice")

# Making Functions
def say(text):
    speaker.speak(f"{text}")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        # r.pause_threshold = 1  #Defualt Value is 0.8
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)
        try:
            print("Recognizing .....")
            query = r.recognize_google(audio, language="en-in")
            print(f"User Said : {query}")
            return query
        except Exception as e:
            return "Some error occured, Sorry from Jarvis"

def wishme():
    print("Hello Sir")
    hour = datetime.datetime.now().hour
    if (hour >= 0 and hour <= 12):
        print("Good Morning, Sir")
        say("Good Morning, Sir")
    elif (hour > 12 and hour < 18):
        print("Good Evening, Sir")
        say("Good Evening, Sir")
    else:
        print("Good Night, Sir")
        say("Good Night, Sir")

# Main Function

if __name__ == '__main__':
    # say("Hello")
    wishme()
    while True:
        print("Listening .....")
        query = takeCommand()
        sites = [["YouTube", "https://www.youtube.com/"], ["Google", "https://www.google.com/"],
                 ["Linkedin", "https://www.linkedin.com/"], ["facebook", "https://www.facebook.com/"],
                 ["twitter", "https://www.twitter.com/"], ["instagram","https://www.instagram.com/"], ["github","https://github.com/"],["Hacker rank","https://www.hackerrank.com/"]]

        # Opening Site
        # Todo: Add more sites
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]} Sir ....")
                print(f"Opening {site[0]} Sir ....")
                webbrowser.open(site[1])

        apps = [["vs Code", "C:\\Program Files\\Microsoft VS Code\\Code.exe"],
                ["browser", "C:\\Program Files\\Mozilla Firefox\\firefox.exe"],
                ["pycharm", "C:\\Program Files\\JetBrains\\PyCharm 2023.1.2\\bin\\pycharm64.exe"],
                ["microsoft word", "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.exe"],
                ["Microsoft Excel", "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.exe"],
                ["Microsoft powerpoint", "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.exe"]]

        for app in apps:
            if f"open {app[0]}".lower() in query.lower():
                print(f"Opening {app[0]}")
                say(f"Opening {app[0]}")
                os.startfile(app[1])

        # Open Music
        # Todo : Add more Songs
        if "play music".lower() in query.lower():
            musicpath = "C:\\Users\\khans\\Downloads\\y2mate.com - The Kidd Mohali Anthem ft Sikander Kahlon.mp3"
            os.startfile(musicpath)

        # Date Time
        elif "the time".lower() in query.lower():
            strfdate = datetime.datetime.now().strftime("%H:%M:%S")
            print(strfdate)
            say(f"The time is {strfdate}, Sir")

        # Opening Applications
        # Todo : Create a list of Apllication like sites and iterate
        # elif "open vs code".lower() in query.lower():
        #     os.startfile("C:\\Program Files\\Microsoft VS Code\\Code.exe")
        #     print("Opening Vs Code")
        #     say("Opening Vs Code")

        elif "wikipedia".lower() in query.lower():
            try:
                say("Ok wait sir, I'm searching...")
                query = query.replace("wikipedia", "")
                result = wikipedia.summary(query, sentences=2)
                print(result)
                print("Would you like to listen the result, Sir")
                say("Would you like to listen the result, Sir")
                print("Taking Decision .... ")
                decision = takeCommand()
                if "yes".lower() in decision:
                    say(result)
                else:
                    pass
            except:
                say("Can't find this page sir, please ask something else")

        elif "Offline".lower() in query.lower():
            hour = datetime.datetime.now().hour
            if (hour >= 0 and hour <= 12):
                print("Okay Sir i am going offline, Have a Good day, Sir")
                say("Okay Sir i am going offline, Have a Good day, Sir")
            elif (hour > 12 and hour < 18):
                print("Okay Sir i am going offline, Have a Good Evening, Sir")
                say("Okay Sir i am going offline, Have a Good Evening, Sir")
            else:
                print("Okay Sir i am going offline, Have a Good Night, Sir")
                say("Okay  Sir i am going offline, Have a Good Night, Sir")
            quit()

        elif "screenshot".lower() in query.lower():
            image = pyautogui.screenshot()
            hour = datetime.datetime.now().hour
            min = datetime.datetime.now().minute
            sec = datetime.datetime.now().second
            image.save(f"C:\\Users\\khans\\Pictures\\Jarvis\\screenshot{hour}{min}{sec}.png")
            print("Screenshot Taken")
            say("Screenshot Taken")

        elif "remember that" in query:
            say("What should I remember")
            data = takeCommand()
            say("You said me to remember that" + data)
            print("You said me to remember that " + str(data))
            remember = open("data.txt", "a+")
            remember.write(data+"\n")
            remember.close()

        elif "do you remember anything" in query:
            remember = open("data.txt", "r")
            say("You told me to remember that" + remember.read())
            print("You told me to remember that " + str(remember))

        # else:
        #     print("Sorry Sir, Can not process what you said, please say again")
        #     say("Sorry Sir, Can not process what you said, please say again")
