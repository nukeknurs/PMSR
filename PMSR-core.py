import speech_recognition as sr  # import the library
import win32com.client
import pandas as pd

def strkom_do_parametrow(kom):
    for i in kom:
        try:
            int(i)
            ind=kom.find(i)
            break
        except:
            pass
    return [kom[:ind],kom[ind:]]

#def keyword(keyword):




act_cell=input('Podaj komórkę początkową np. A1')
nazwakomorki=strkom_do_parametrow(act_cell)
str_ind=nazwakomorki[0]
num_ind=int(nazwakomorki[1])
nazwakomorki=nazwakomorki[0]+nazwakomorki[1]
ExcelApp = win32com.client.GetActiveObject("Excel.Application")
ExcelApp.Visible = True

workbook = ExcelApp.Workbooks.Open(r"C:/Users/Odin/Desktop/outputkupy.xlsx")

text=''
preferred_phrases=['edytuj','Zakończ']
r = sr.Recognizer()  # initialize recognizer
while True:
    with sr.Microphone() as source:  # mention source it will be either Microphone or audio files.
        print("Speak Anything :")
        audio = r.listen(source)  # listen to the source
        try:
            text = r.recognize_google(audio,language='PL')
            if text.split()[0] in preferred_phrases:
                if text.split()[0] == 'edytuj':
                    if text.split()[1]=='komórka':
                        try:
                            l = strkom_do_parametrow(text.split()[2])
                            str_ind_e = l[0]
                            num_ind_e = int(l[1])
                            print(str_ind,num_ind)
                        except:
                            print('Coś poszło nie tak z nazwą komorki')
                            next()
                        print("Na co zedytować?")
                        audio = r.listen(source)
                        text = r.recognize_google(audio, language='PL')
                        ExcelApp.Range(str_ind_e + str(num_ind_e)).Value = [text]
                        next()


                if text.split()[0] == 'zakończ':
                    break
            ExcelApp.Range(str_ind+str(num_ind)).Value =[text]
            print('xd')
            num_ind=num_ind+1
            #zaktualizuj_komorke(str_ind,num_ind)
            print("You said : {}".format(text))
        except:
            print("Sorry could not recognize your voice")