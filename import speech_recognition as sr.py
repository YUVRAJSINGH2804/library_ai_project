import speech_recognition as sr
from gtts import gTTS
import pandas as pd
import os
import pygame
import pyttsx3
import datetime

# Load the Excel sheet with book information
excel_path = "C:/Users/Yuvraj/Downloads/Book1.xlsx"  # Change this to the path of your Excel file
df = pd.read_excel(excel_path)

# Initialize the recognizer
recognizer = sr.Recognizer()
engine=pyttsx3.init('sapi5')
voices= engine.getProperty('voices')
engine.setProperty('voices', voices[0].id)


def wish():

    hour = int(datetime.datetime.now().hour)

    if hour>=0 and hour<=12:
        speak("good morning")

    elif hour>12 and hour<18:
        speak("good afternoon")

    else:
        speak("good evening")
    speak("I am Jarvis sir . please tell me how can i help you")

def speak(audio):
    engine.say(audio)
    print(audio)
    engine.runAndWait()

# Function to recognize voice command
def recognize_voice():
    with sr.Microphone() as source:
        print("Listening...")
        audio = recognizer.listen(source)

        try:
            print("Recognizing...")
            command = recognizer.recognize_google(audio).lower()
            print("You said:", command)
            return command
        except sr.UnknownValueError:
            print("Could not understand audio")
            return None
        except sr.RequestError as e:
            print("Could not request results; {0}".format(e))
            return None

# Function to find book location
def find_book_location(book_name):
    book_row = df[df['Book Name'].str.lower() == book_name]
    if not book_row.empty:
        location = book_row.iloc[0]['Location']
        return f"The book '{book_name}' is located at '{location}'."
    else:
        return f"Sorry, the book '{book_name}' was not found in the library."

# Main loop
while True:
    wish()
    print("Please give a voice command...")
    voice_command = recognize_voice()

    if voice_command:
        if "exit" in voice_command:
            print("Exiting...")
            break
        else:
            response_text = find_book_location(voice_command)
            print(response_text)

            # Generate voice output
            tts = gTTS(text=response_text, lang='en')
            tts.save('output.mp3')
            
            # Play the audio using pygame
            pygame.mixer.init()
            pygame.mixer.music.load('output.mp3')
            pygame.mixer.music.play()

            break  # Stop after providing book location


