import datetime
import webbrowser
import speech_recognition as sr
import os
import win32com.client
import random
from num2words import num2words
from openai import OpenAI
speaker = win32com.client.Dispatch("SAPI.SpVoice")

firstError = "I am not able to recognize your task please check your microphone and try again"


def say(text):
    speaker.Speak(text)

def takeCommand():
    r = sr.Recognizer()
    try:
        with sr.Microphone() as source:
            r.pause_threshold = 1
            audio = r.listen(source)
            try :
                query = r.recognize_google(audio, language='en-in')
                print(f"user said : {query}")
                return query
            except Exception as e:

                return "Something Went Wrong"
    except Exception as e:
        return firstError

if __name__ == "__main__":
    print('Welcome To Sillycon Bot')
    say("Yes sir what i do for you")
    while True:
        print("Listening...")
        query = takeCommand()
        sites = [["youtube","https://www.youtube.com"],["google", "https://www.google.com"],["instagram", "https://www.instagram.com"],["wikipedia", "https://www.wikipedia.com"],["ratikant", "https://mahabocw.in/"],]
        for site in sites:
            if f"open {site[0]}".lower() in query.lower():
                say(f"opening {site[0]} sir...")
                webbrowser.open(site[1])

        if "the time" in query:
            now = datetime.datetime.now()
            hour = now.strftime("%I")  # 12-hour format
            minute = now.strftime("%M")
            period = now.strftime("%p")  # AM or PM
            # Convert hour and minute to words (remove leading 0s)
            hour_words = num2words(int(hour))
            minute_words = num2words(int(minute)) if int(minute) != 0 else "o'clock"
            say(f"It's {hour_words} {minute_words} {period}, sir.")

        if "open chrome".lower() in query.lower():
            os.startfile(r"C:\Program Files\Google\Chrome\Application\chrome.exe")



        #   This Code Is For Video Purpose
        # if "video" in query:
        #     folder_path = "E:/RP Movies/"
        #     video_extensions = ('.mp4', '.mkv', '.avi', '.mov')
        #
        #     videos = [f for f in os.listdir(folder_path) if f.endswith(video_extensions)]
        #     if videos:
        #         random_video = random.choice(videos)
        #         os.startfile(os.path.join(folder_path, random_video))

        # This Code Is For Music Purpose
        # if "play music" in query:
        #     music_folder = "E:/RP Music"
        #     music_extensions = ('.mp3', '.wav', '.aac')
        #
        #     songs = [f for f in os.listdir(music_folder) if f.endswith(music_extensions)]
        #     if songs:
        #         song = random.choice(songs)
        #         os.startfile(os.path.join(music_folder, song))


        if(query == firstError):
            say(query)
            break
        elif(query == "exit"):
            say("Thank You For Using Mi Sir")
            break
        say(query)



