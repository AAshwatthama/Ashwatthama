import win32com.client
import speech_recognition as sr
import os
import webbrowser
import openai
from pathlib import Path 
from AppOpener import open

#webdriver package



speaker = win32com.client.Dispatch('SAPI.SpVoice')
chatstr=''
def chat(prompt):
    global chatstr
    print(chatstr)
    # client=OpenAI()
    openai.api_key='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'
    chatstr+=f'Vaibhav:{text}\n Aswatthama:'
    responce=openai.completions.create(
        model='gpt-3.5-turbo-0301',
        prompt=chatstr,
        # temprature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    say(responce['choice'][0]['text'])
    chatstr +=f"{responce['choice'][0]['text']}\n"
    return responce['choice'][0]['text']
def say(text):
    speaker.Speak(text)

def takecommand():
    r=sr.Recognizer()
    try:
        with sr.Microphone() as source: #distutils package is removed in python version 3.12(Install setuptools in your machine)
            r.pause_threshold=0.6
            print('listening..')
            audio=r.listen(source) #create audio
            query=r.recognize_google(audio,language='en-in') #create text using audio
            
            print(f'User said:{query}')
            return query
    except Exception as e:
        print(e)
print('Hello I am Ashwatthama A.I')
say('Hello I am Ashwatthama A I') 

# say('Hello I am Ashwatthama A I')
text=takecommand()
try:
        
    sites=[['youtube','https://www.youtube.com/'],['google','https://www.google.com/#sbfbu=1&pi='],['gmail','https://mail.google.com/mail/u/0/#inbox'],['whatsapp web','https://web.whatsapp.com/'],['stack overflow','https://stackoverflow.com/questions/69919970/no-module-named-distutils-but-distutils-installed']
        ,['linkedin','https://www.linkedin.com/feed/']
        ,['python documentation','https://www.python.org/'],['spotify','https://open.spotify.com/']]
    for site in sites:
        if f'open {site[0]}'.lower() in text.lower():
            webbrowser.open(site[1])
            say(f'opening {site[0]} sir')  
# except:
    # pass
# try:
    if 'Play music'.lower() in text.lower():
        path=r"E:\downfall-21371.mp3"
        # os.startfile("E:\downfall-21371.mp3") 
        os.system(path)
# except: 
    # pass
# try:
    drives=[['c drive','c:'],['d drive','d:'],['e drive','e:'],['f drive','f:'],['g drive','g:']]
    for drive in drives:
        if f'open {drive[0]}' in text.lower():
            os.startfile(drive[1])
# except:
    # pass
# try:
    if 'open chat' in text.lower():
        while True:
            chat(chatstr)
# except:
    # pass

# try:
    if 'open file' in text:
        say('which file you want to open sir')
        filename=takecommand()
        y=os.path.abspath(filename)
        x=y.replace(r'\main','')
        os.startfile(x)
# except:
    # pass

# try:
    if 'open app' in text:
        say('which app you want to open sir')
        a=takecommand()
        open(a)

except:
    pass



