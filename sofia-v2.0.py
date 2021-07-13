from __future__ import print_function
from mailmerge import MailMerge
import speech_recognition as sr
import pyttsx3
import sys
import re
import webbrowser
import smtplib
import requests
import subprocess
from pyowm import OWM
from bs4 import BeautifulSoup as soup
from time import strftime

#from datetime import date

def sofiaResponse(audio):
    engine = pyttsx3.init()
    engine.say(audio)
    engine.runAndWait()
    
def myCommand():
    "listens for commands"
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print('SAY SOMETHING.......')
        r.pause_threshold = 1
        r.adjust_for_ambient_noise(source, duration=1)
        audio = r.listen(source)
    try:
        command = r.recognize_google(audio).lower()
        print('You said: ' + command + '\n')
    #loop back to continue to listen for commands if unrecognizable speech is received
    except sr.UnknownValueError:
        print('....')
        command = myCommand();
    return command

def assistant(command):
    "if statements for executing commands"
    if 'name' in command:
       # reg_ex = re.search('name (.*)', command)
       r1 = sr.Recognizer()
       with sr.Microphone() as source:
           print('Please tell the Patient Name...')
           r1.pause_threshold = 1
           r1.adjust_for_ambient_noise(source, duration=1)
           audioName = r1.listen(source)
       try:
           commandName=r1.recognize_google(audioName).lower()
           print('You said: '+commandName + '\n')
           sofiaResponse('Inserting name into the prescription...')
           template = "Prescription-Template.docx"
           document = MailMerge(template)
           print(document.get_merge_fields())
           document.merge(
                   PatientName=commandName)
           document.write('Prescription3.docx')
       except sr.UnknownValueError:
           print('.......')
           
    elif 'age'in command:
       r1 = sr.Recognizer()
       with sr.Microphone() as source:
           print('Please tell the Patient Age...')
           r1.pause_threshold = 1
           r1.adjust_for_ambient_noise(source, duration=1)
           audioAge = r1.listen(source)
       try:
           commandName=r1.recognize_google(audioAge).lower()
           print('You said: '+commandName + '\n')
           sofiaResponse('Inserting Age into the prescription...')
           template = "Prescription-Template.docx"
           document = MailMerge(template)
           print(document.get_merge_fields())
           document.merge(
                   PatientAge=commandName)
           document.write('Prescription3.docx')
       except sr.UnknownValueError:
           print('.......')
        
           
    elif 'shutdown' in command:
        sofiaResponse('Bye bye Sir. Have a nice day')
        sys.exit()
        
    elif 'hello' in command:
        day_time = int(strftime('%H'))
        if day_time < 12:
            sofiaResponse('Hello Sir. Good morning')
        elif 12 <= day_time < 18:
            sofiaResponse('Hello Sir. Good afternoon')
        else:
            sofiaResponse('Hello Sir. Good evening')
    elif 'help me' in command:
        sofiaResponse("""
        You can use these commands and I'll help you out:
1. Open reddit subreddit : Opens the subreddit in default browser.
        2. Open xyz.com : replace xyz with any website name
        3. Send email/email : Follow up questions such as recipient name, content will be asked in order.
        4. Current weather in {cityname} : Tells you the current condition and temperture
        5. Hello
        6. play me a video : Plays song in your VLC media player
        7. change wallpaper : Change desktop wallpaper
        8. news for today : reads top news of today
        9. time : Current system time
        10. top stories from google news (RSS feeds)
        11. tell me about xyz : tells you about xyz
        """)
       
        
        
sofiaResponse('Hi User, I am Sofia and I am your personal voice assistant, Please give a command or say "help me" and I will tell you what all I can do for you.')
#loop to continue executing multiple commands
#sofiaResponse('Hello harshal...We will win Smart India Hackathon!!')
while True:
    assistant(myCommand())