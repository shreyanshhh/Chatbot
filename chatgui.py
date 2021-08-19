import nltk
from nltk.stem import WordNetLemmatizer
lemmatizer = WordNetLemmatizer()
import keras
import numpy as np
import json
import pickle
from keras.models import load_model
model = load_model('chatbot_model.h5')

import random
import cv2
import tkinter
from tkinter import *
from PIL import Image, ImageTk

import win32com.client
speaker = win32com.client.Dispatch("SAPI.SpVoice")

import speech_recognition as sr
r=sr.Recognizer()

intents = json.loads(open('intents.json').read())
words = pickle.load(open('words.pkl','rb'))
classes = pickle.load(open('classes.pkl','rb'))

def clean_up_sentence(sentence):
    sentence_words = nltk.word_tokenize(sentence)
    sentence_words = [lemmatizer.lemmatize(word.lower()) for word in sentence_words]
    return sentence_words

def bow(sentence, words, show_details=True):
    # tokenize the pattern
    sentence_words = clean_up_sentence(sentence)
    print(sentence_words)
    # bag of words - matrix of N words, vocabulary matrix
    bag = [0]*len(words)
    for s in sentence_words:
        for i,w in enumerate(words):
            if w == s:
                # assign 1 if current word is in the vocabulary position
                bag[i] = 1
                if show_details:
                    print ("found in bag: %s" ,w)
    return(np.array(bag))

def predict_class(sentence, model):
    # filter out predictions below a threshold
    p = bow(sentence, words,show_details=False)
    res = model.predict(np.array([p]))[0]
    ERROR_THRESHOLD = 0.25
    results = [[i,r] for i,r in enumerate(res) if r>ERROR_THRESHOLD]
    # sort by strength of probability
    results.sort(key=lambda x: x[1], reverse=True)
    return_list = []
    for r in results:
        return_list.append({"intent": classes[r[0]], "probability": str(r[1])})
    return return_list



def getResponse(ints, intents_json):
    tag = ints[0]['intent']
    list_of_intents = intents_json['intents']
    for i in list_of_intents:
        if(i['tag']== tag):
            result = random.choice(i['responses'])
            break
    return result


def chatbot_response(msg):
        ints = predict_class(msg, model)
        print(ints)
        res = getResponse(ints, intents)
        return res

def send():
    msg = EntryBox.get("1.0",'end-1c').strip()
    EntryBox.delete("0.0",END)

    if msg != '':
        ChatLog.config(state=NORMAL)
        ChatLog.insert(END,"You: "+msg+ '\n\n')
        res=chatbot_response(msg)
        
        
        ChatLog.insert(END,"Robo: "+res+ '\n\n')
        ChatLog.config(state=DISABLED)
        speaker.Speak(res)
        ChatLog.yview(END)

def speech():
    L = Label(base,text="Listening!!!")
    L.place(x=50,y=350,height=30)
    with sr.Microphone() as source:
        print("Talk")
        
        audio_text = r.listen(source)    
        try:
            # using google speech recognition
            msg = r.recognize_google(audio_text)
            if msg != '':
                ChatLog.config(state=NORMAL)
                ChatLog.insert(END,"You: "+msg+ '\n\n')
                res=chatbot_response(msg)
                ChatLog.insert(END,"Robo: "+res+ '\n\n')
                ChatLog.config(state=DISABLED)
##                speaker.Speak(res)
                ChatLog.yview(END)
                print("Text: ",msg)
        except:
             print("Sorry, I did not get that")
    L.after(1000, L.destroy)
    print("connected")
    

base = Tk()
base.title("Chat Bot")
base.geometry("500x500")
base.resizable(width = 'false', height= 'false')
image2 =Image.open('back.jpg')
background_image=ImageTk.PhotoImage(image2)
background_label = Label(base, image=background_image)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

image=Image.open('mic2.jpg')
img=ImageTk.PhotoImage(image)
but=Button(base,width=30, image = img,command=speech)
but.place(x=6,y=350,height=40)

ChatLog = Text(base, bd=0, bg="white", height="8", width="50", font="Arial",)
ChatLog.config(state=DISABLED)

scrollbar = Scrollbar(base,command = ChatLog.yview)
ChatLog['yscrollcommand']=scrollbar.set
scrollbar.place(x=376,y=25, height=300)

button = Button(activebackground='red', activeforeground='black', bg='green',text='send message',width=25,command=send)
button.place(x=6, y=401, height=90)

EntryBox = Text(base, bd=0, bg="white",width="32", height="5", font="Arial")
EntryBox.place(x=200,y=400,height=90)

ChatLog.place(x=6,y=25, height=300, width=370)
base.mainloop()
