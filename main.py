import speech_recognition as sr
import win32com.client
import webbrowser
import datetime
import os
import yagmail

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        r.adjust_for_ambient_noise(source)
        r.energy_threshold = 150
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")
        #speaker.Speak(query)

    except Exception as e:
        # print(e)
        print("Say that again please...")
        return "None"
    return query

emails={'tania': 'mahata.tania@gmail.com','tanya': 'mahata.tania@gmail.com','taniya':'mahata.tania@gmail.com','jagadish':'jtmahata@gmail.com','ashitha':'ashithanair2002@gmail.com','my mail':'mahata.tanisha76@gmail.com'}
if __name__ == "__main__":
    speaker.Speak("Hello I am Tanisha's Bot, how may i help you?")
    emailsent=True
    while True:
        if emailsent==True:
            query = takeCommand().lower()
        else:
            query='send an email'

        if 'quit' in query:
            speaker.Speak("Quitting Mam...")
            break
        if 'hello' in query:
            speaker.Speak("Hi..")
        elif 'open youtube' in query.lower():
            speaker.Speak('Opening youtube mam...')
            webbrowser.open('https://www.youtube.com/')
        elif "the time" in query:
            hour = datetime.datetime.now().strftime("%H")
            min = datetime.datetime.now().strftime("%M")
            speaker.Speak(f"Mam time is {hour} and {min} minutes")
        elif "the date" in query:
            date= datetime.date.today()
            speaker.Speak(f"Today's Date is {date}")
        elif 'open spotify' in query:
            os.system('start Spotify')
        elif 'search' in query:
            speaker.Speak("Searching for you...")
            mainpart= query.split('search')[1].split(' ')
            searchstr= ('+'.join(mainpart[1:]))
            webbrowser.open(f'https://www.google.com/search?q={searchstr}')
        elif 'send an email' in query:
            speaker.Speak("Who do you want to send the mail to?")
            try:

                name1= takeCommand()
                reciever = emails[name1.lower()]
                speaker.Speak("From which email?")
                name2 = takeCommand()
                sender = yagmail.SMTP(emails[name2.lower()])
                print(reciever, sender)
                sender.send(to=reciever, subject='This is an automated mail', contents='Hello didi i created a bot to send this!')
                emailsent=True
            except Exception as e:
                print('No name like that exists!')
                speaker.Speak('No name like that exists!')
                emailsent=False



