from gtts import gTTS
import speech_recognition as sr
import pyttsx3
import os
import re
import webbrowser
import requests
import win32com.client as wincl
import socket
import random
import datetime
import json
import time
import string
from pathlib import Path
from selenium import webdriver
from urllib3 import quote

def speak_out(audio, voice):
    #"speaks audio passed as argument"
    engine = pyttsx3.init() # object creation
    voices = engine.getProperty('voices')       #getting details of current voice
    engine.setProperty('voice', voices[voice].id)   #changing index, changes voices
    pyttsx3.speak(audio)
    engine.setProperty('rate', 300)     # setting up new voice rate



def listen():
    #"listens for commands"

    r = sr.Recognizer()

    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source, duration=0.5)
        audio = r.listen(source)

    try:
        command = r.recognize_google(audio).lower()
        return command

    #loop back to continue to listen for commands if unrecognizable speech is received
    except sr.UnknownValueError:
        speak_out('Your last command couldn\'t be heard ! Please try again',voice)
        #speak.Speak('Your last command couldn\'t be heard')
        command = listen()
        return command



def observe():
    r = sr.Recognizer()

    with sr.Microphone() as source:
        r.pause_threshold = 1
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)

    try:
        read = r.recognize_google(audio).lower()
        if myname in read:
            return read

    #loop back to continue to listen for commands if unrecognizable speech is received
    except sr.UnknownValueError:
            observe()



def is_connected():
    try:
        # connect to the host -- tells us if the host is actually
        # reachable
        socket.create_connection(("1.1.1.1", 53))
        return True
    except OSError:
        pass
    return False



def greet():
    morning=['good morning','hiii','hello','heyy']
    afternoon=['good afternoon','hiii','hello','heyy']
    evening=['good evening','hiii','hello','heyy']
    night=['hiii','hello','heyy']
    hour=int(datetime.datetime.now().hour)
    if hour>=4 and hour<12:
        speak_out(random.choice(morning) + usr_name,voice)
    
    elif hour>=12 and hour<=5:
        speak_out(random.choice(afternoon) + usr_name,voice)
   
    elif hour>=6 and hour<9:
        speak_out(random.choice(evening) + usr_name,voice)
    
    else:
        speak_out(random.choice(night) + usr_name,voice)
    



def web_search():
    search=command
    speak_out('I can search the web for you. Do you want to continue?',voice)
    listen()
    if 'yes' in command or 'yeah' in command or 'yep' in command or 'sure' in command or 'yup' in command or 'okay' in command:
        omit=['open','run,','execute','start','begin','create','initiate']
        search=search.replace(omit,'')
        search=search.replace(' ','+')
        webbrowser.open_new_tab('http://www.google.com/search?' + search)
    elif 'no' in command or 'nope' in command or 'never' in command:
        speak_out('Okay...',voice)
    else:
        speak_out('Sorry... I did not get you. Please try again...',voice)
        time.sleep(2)
        web_search()



def joke():
    res = requests.get(
            'https://icanhazdadjoke.com/',
            headers={"Accept":"application/json"}
            )
    if res.status_code == requests.codes.ok:
        speak_out('Here is an awesome joke for you- ',voice)
        speak_out(str(res.json()['joke']),voice)
    else:
        speak_out('oops!I ran out of jokes',voice)



def shutdown():
    speak_out('Do you want to shut down your PC?',voice)
    choice=listen()
    if 'yes' in choice or 'yeah' in choice or 'yep' in choice or 'sure' in choice or 'yup' in choice or 'okay' in choice:
        speak_out('Your system will shut down in 10 seconds. Make sure that you save all applications.',voice)
        time.sleep(10)
        os.system('shutdown /s')
    elif 'no' in choice or 'nope' in choice or 'never' in choice:
        speak_out('Okay...',voice)
    else:
        speak_out('Sorry... I did not get you. Please try again',voice)
        time.sleep(2)
        shutdown()



def restart():
    speak_out('Do you want to restart your PC?',voice)
    choice=listen()
    if 'yes' in choice or 'yeah' in choice or 'yep' in choice or 'sure' in choice or 'yup' in choice or 'okay' in choice:
        speak_out('Your system will restart in 10 seconds. Make sure that you save all applications.',voice)
        time.sleep(10)
        os.system('shutdown /r')
    elif 'no' in choice or 'nope' in choice or 'never' in choice:
        speak_out('Okay...',voice)
    else:
        speak_out('Sorry... I did not get you. Please try again',voice)
        time.sleep(2)
        restart()



def logoff():
    speak_out('Do you want to log your PC off?',voice)
    choice=listen()
    if 'yes' in choice or 'yeah' in choice or 'yep' in choice or 'sure' in choice or 'yup' in choice or 'okay' in choice:
        speak_out('Your system will log off in 10 seconds. Make sure that you save all applications.',voice)
        time.sleep(10)
        os.system('shutdown /l')
    elif 'no' in choice or 'nope' in choice or 'never' in choice:
        speak_out('Okay...',voice)
    else:
        speak_out('Sorry... I did not get you. Please try again',voice)
        time.sleep(2)
        logoff()



def hibernate():
    speak_out('Do you want to hibernate your PC?',voice)
    choice=listen()
    if 'yes' in choice or 'yeah' in choice or 'yep' in choice or 'sure' in choice or 'yup' in choice or 'okay' in choice:
        speak_out('Your system will hibernate in 10 seconds. All your applications will run as they are on the next boot..',voice)
        time.sleep(10)
        os.system('shutdown /h')
    elif 'no' in choice or 'nope' in choice or 'never' in choice:
        speak_out('Okay...',voice)
    else:
        speak_out('Sorry... I did not get you. Please try again',voice)
        time.sleep(2)
        hibernate()



def search_web(): 
    if 'youtube' in command: 
        speak_out("Opening in youtube",voice) 
        query= command.replace("youtube",'')
        query= command.replace(" ",'+')
        query= command.replace("search",'')
        if 'in' in command:
            query= command.replace('in','')
        if 'for' in command:
            query= command.replace("for",'')

        webbrowser.open("http://www.youtube.com/results?search_query =" + query) 
        return
  
    elif 'wikipedia' in command: 
  
        speak_out("Opening Wikipedia",voice) 
        indx = command.split().index('wikipedia') 
        query = command.split()[indx + 1:] 
        driver.get("https://en.wikipedia.org/wiki/" + '_'.join(query)) 
        return
  
    else: 
  
        if 'google' in command: 
  
            indx = command.split().index('google') 
            query = command.split()[indx + 1:] 
            driver.get("https://www.google.com/search?q =" + '+'.join(query)) 
  
        elif 'search' in command: 
  
            indx = command.split().index('search') 
            query = command.split()[indx + 1:] 
            driver.get("https://www.google.com/search?q =" + '+'.join(query)) 
  
        else: 
  
            driver.get("https://www.google.com/search?q =" + '+'.join(command.split())) 


def open_app():
    speak_out('Please speak the name of the application or website that you want to open once more.',voice)
    select=listen()
        
    if 'gmail' in select or 'mail' in select:
        #reg_ex = re.search('open gmail (.*)', command)
        url = 'https://www.gmail.com/'
        webbrowser.open(url)
        speak_out('Done!',voice)

    elif 'calender' in select or 'schedule' in select:
        url = 'https://calendar.google.com/'
        webbrowser.open(url)
        speak_out('Done!',voice)

    elif 'website' in select or '.com' in select or '.org' in select or '.in' in select or '.io' in select:
        webbrowser.open(select)
        speak_out('Done!',voice)

    else:
        try:
            speak_out( select + ' will be opened if it is in the environment path',voice)
            z=select +' > nul 2> nul'
            print(z)
            os.system(z)
        except:
            speak_out('Sorry, I cannot open ' + select + '. Please try again.',voice)

        



def exitme():
    speak_out('Speak yes if you want me to exit an application or speak no if you want to close me.',voice)
    choice=listen()
    if 'yes' in choice or 'yeah' in choice or 'yep' in choice or 'sure' in choice or 'yup' in choice or 'okay' in choice:
        speak_out('You chose to close other applications... Please speak the name of the application that you want to exit.',voice)
        select=listen()
        try:
            x=('taskkill /F /IM ')
            y=('.exe /T > nul 2> nul')
            z=x+select+y
            speak_out('If '+ select + 'is open, ' + select +' will close in 10 seconds... Make sure to save any unsaved changes',voice)
            os.system(z)
        except:
            speak_out('Sorry, I cannot close ' + select + '...',voice)
            os.system('cls')

    elif 'no' in choice or 'nope' in choice or 'never' in choice:
        exitall()
    else:
        speak_out('You did not enter a valid choice... Please try again.',voice)
        exitme()



def tell_me():
    message = 'Ask me to do something, but please be short.'
    speak_out(message,voice)

def process():
    
    if ('hello' in command) or ('hi' in command) or ('hey' in command):
        tell_me()
        

    elif('open' in command) or ('run' in command) or ('execute' in command) or ('start' in command) or ('begin' in command) or ('create' in command) or ('initiate' in command):
        open_app()
 

    elif 'joke' in command:
        joke()


    elif 'shut down' in command:
        shutdown()


    elif 'search' in command:
        search_web()


    elif 'camera' in command or 'web cam' in command or 'take photo' in command or 'take video' in command:
        speak_out('opening camera....',voice)
        os.system('start microsoft.windows.camera:')


    elif 'exit' in command or 'close' in command or 'quit' in command or 'shut down' in command or 'shut' in command:
        exitme()



def exitall():
    speak_out('I will exit in 5 seconds.',voice)
    exit_st=['Byeee...','See you soon...','Thanks for using me...', 'Have a great day...','Take care...']
    speak_out(random.choice(exit_st) + usr_name,voice)
    exit()


 
def setup():
    speak_out('Hiii!!!... I am your digital assistant!.. I am there to help you whenever you want...',voice)
    speak_out('Before getting started, I would like to know you more, so that I can do my best in assisting you...',voice)
    speak_out('Would you like to set me up?',voice)
    print('Speak yes or no.')
    choice=listen()
    if'yes' in choice or 'yeah' in choice or 'yep' in choice or 'sure' in choice or 'yup' in choice or 'okay' in choice:
        speak_out('Okay, so lets get started....',voice)
        os.system('cls')
        print('*********PLEASE ENTER YOUR INFORMATION*********')
        speak_out('Please note that you have to type your information',voice)
        print('please enter your name: ',end='')
        speak_out('Please enter your name',voice)
        uname=string.capwords(input())
        print('please enter your city: ',end='')
        speak_out('Please enter you city',voice)
        location=string.capwords(input())
        speak_out('Okay... Now I have sufficient information about you. Now its time to customize me.',voice)
        os.system('cls')
        print('*********PLEASE CUSTOMIZE ME*********')
        speak_out('What would you like to call me?',voice)
        speak_out('Please give me a name.',voice)
        name=input('Please give me a name: ')
        def ask_voice():
            speak_out('What should my voice sound like?',voice)
            print('Enter 1 for female voice and 0 for male voice: ', end='')
            speak_out('Type 1 if you want me to sound like a female or type 0 if you want me to sound like a male.',voice)
            in_voice=input()
            if in_voice=='0' or in_voice=='1':
                return in_voice
            else:
                speak_out('Enter 0 or 1 only. Please retry',voice)
                ask_voice()
        in_voice=ask_voice()
        os.system('cls')
        speak_out('Thats it! I am all set and ready to asist you.',voice)
        
        info={"uname":uname, "location": location, "voice":in_voice,"name":name}
        data=json.dumps(info)
        with open('user_info.json','w') as f:
            f.write(data)
            f.close()
        speak_out('I have to shut down in order to save your details. Please start me again.',voice)
    else:
        speak_out('No worries, you can set me up later on...',voice)


i=0
while(True):
    data=json.load(open('user_info.json'))
    usr_name=(data["uname"])
    voice=int((data["voice"]))
    myname=(data["name"])
    if usr_name=='Master':
        setup()


    x=is_connected()
    if x!=True:
        print('Sorry, but I cannot work offline. See you soon!!!')
        speak_out('Sorry, but I cannot work offline. See you soon!!!',voice)
        exitall()
    if i==0:
        greet()
        tell_me()
        command=listen()
        print("You said: " + command)
        i=i+1
        #if usr_name in command:
        process()


    elif  i==1:
        greet()
        speak_out('You can take my name whenever you want and I will be ready to serve you',voice)
        speak_out('How may I help you?',voice)
        command=listen()
        speak_out("You said: " + command,voice)
        i=i+1
        process()
        speak_out('I will now run in background. You can wake me up by taking my name.',voice)

    else:
        command=observe()
        if command !=None:
            print("You said: " + command)
            process()
            i=i+1