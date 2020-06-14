#This hash sign represents a beginning of a comment. You will find many such below in the code telling you to make necessary changes according to your device.
# Make sure you have a microphone connected to run this program.
# please install following libraries
import pyautogui
import mysql.connector
import speech_recognition as sr
import time
import re
import sys
import os
import win32com.client
import random
import webbrowser
import cv2
path="'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe' %s"#Enter file location of the default browser for this application
google="http://www.google.com/search?q="
speak = win32com.client.Dispatch('Sapi.SpVoice')
speak.Rate=0.0001
speak.Volume=100
try:
    
    r = sr.Recognizer()
    mic=sr.Microphone()
    a1=['move to position (x,y)','move up by 100','move down by 100','move right by 100','move left by 100']
    a2=['right click','left click','middle click','double click']
    a3=['type']
    a4=['press']
    a5=['show commands']
    a8=['drag to position(x,y)','drag up by 100','drag down by 100','drag right by 100','drag left by 100']
    a7=['open calculator','search','quit','shutdown','restart','capture image','capture screenshot']
    bb=a1+a2+a3+a4+a5+a7+a8
    a6=['-','=',',',"'",'.','tab', 'enter', 'space', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', 'accept', 'add', 'alt', 'altleft', 'altright', 'apps', 'backspace', 'browserback', 'browserfavorites', 'browserforward', 'browserhome', 'browserrefresh', 'browsersearch', 'browserstop', 'capslock', 'clear', 'convert', 'ctrl', 'ctrlleft', 'ctrlright', 'decimal', 'del', 'delete', 'divide', 'down', 'end', 'enter', 'esc', 'escape', 'execute', 'f1', 'f10', 'f11', 'f12', 'f13', 'f14', 'f15', 'f16', 'f17', 'f18', 'f19', 'f2', 'f20', 'f21', 'f22', 'f23', 'f24', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9', 'final', 'fn', 'hanguel', 'hangul', 'hanja', 'help', 'home', 'insert', 'junja', 'kana', 'kanji', 'launchapp1', 'launchapp2', 'launchmail', 'launchmediaselect', 'left', 'modechange', 'multiply', 'nexttrack', 'nonconvert', 'num0', 'num1', 'num2', 'num3', 'num4', 'num5', 'num6', 'num7', 'num8', 'num9', 'numlock', 'pagedown', 'pageup', 'pause', 'pgdn', 'pgup', 'playpause', 'prevtrack', 'print', 'printscreen', 'prntscrn', 'prtsc', 'prtscr', 'return', 'right', 'scrolllock', 'select', 'separator', 'shift', 'shiftleft', 'shiftright', 'sleep', 'space', 'stop', 'subtract', 'tab', 'up', 'volumedown', 'volumemute', 'volumeup', 'win', 'winleft', 'winright', 'yen', 'command', 'option', 'optionleft', 'optionright']
    conn=mysql.connector.connect(host='localhost',user='root',password='1234')#change hostname,usernam,password according to your mysql
    cur=conn.cursor()
    try:
        cur.execute('create database commands')
        conn.commit()
    except mysql.connector.errors.DatabaseError:
        pass
    conn=mysql.connector.connect(host='localhost',user='root',passwd='1234',database='commands')#change hostname,usernam,password according to your mysql
    cur=conn.cursor()
    try:
        cur.execute('create table commands(name varchar(100), value int(10))')
        conn.commit()
        for i in bb:
            cur.execute('insert into commands value("'+i+'",0)')
            conn.commit()
            
    except mysql.connector.errors.ProgrammingError:
        pass

    def text2int (text):
        if ',' in text:
            text=text.replace(',',' comma')
        ar=text.split()
        ans=ar[:]
        units = [
            "zero", "one", "two", "three", "four", "five", "six", "seven", "eight",
            "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen",
            "sixteen", "seventeen", "eighteen", "nineteen",
            ]
        tens = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"]
        scales=['hundred','thousand']
        s=['and']
        for i in range (len(ar)):
            if ar[i] in units+tens+scales+s:
                ans1=0
                arr=[]
                try:
                    while ar[i] in units+tens+scales+s:
                        arr.append(ar[i])
                        i+=1
                        
                except IndexError:
                    pass
                if 'thousand' in arr:
                    ans1+=1000
                    ak=arr[arr.index('thousand'):]
                if 'hundred' in arr:
                    if arr.index('hundred')!=0:
                        for j in range(arr.index('hundred')):
                            if arr[j] in units:
                                ans1+=units.index(arr[j])*100
                            elif arr[j] in tens:
                                ans1+=tens.index(arr[j])*1000
                        ak=arr[arr.index('hundred'):]
                    else:
                        ans1+=100
                        ak=arr[arr.index('hundred'):]
                if 'hundred' not in arr and 'thousand' not in arr:
                    ak=arr[:]
                for j in ak:
                    if j in units:
                        ans1+=units.index(j)
                    elif j in tens:
                        ans1+=tens.index(j)*10
                ho=ans.index(arr[0])
                for j in arr:
                    ans.remove(j)
                ans.insert(ho,str(ans1))
                
                c=[y for y in ans if y in units or y in tens or y in scales]
                
                if len(c)>0:
                    k=text2int(" ".join(ans))
                    return k
                else:
                    fans=" ".join(ans)
                    if 'comma' in fans:
                        fans=fans.replace('comma',',')
                        
                    return fans
                
        return text

    def check(a,a0):
        k=a0.find(a)
        if k==-1:
            return False
        else:
            return True
    def checkf(a0):
        for n, i in enumerate(a0):
            if i == 'spacebar':
                a0[n] = 'space'
            elif i=='minus':
                a0[n]='-'
            elif i=='equal':
                a0[n]='='
            elif i=='comma':
                a0[n]=','
            elif i=='quote':
                a0[n]="'"
            elif i=='dot':
                a0[n]='.'
            elif i=='pause':
                a0[n]='.'
            elif i=='stop':
                a0[n]='.'
            elif i=='control':
                a0[n]='ctrl'
            elif i=='function':
                a0[n]='fn'
        units=["zero", "one", "two", "three", "four", "five", "six", "seven", "eight","nine"]
        for m in a0:
            if m in units:
                m=str(units.indes(m))
        c=[x for x in a0 if x in a6]
        return c
    def fn():
        with mic as source:
            print('Listening')
            speak.Speak('I am listening')
            audio= r.record(source=mic, duration=10)
            speak.Speak('Wait till I execute')
            print('Executing')
            
        try:
            a0=r.recognize_google(audio, language='en-IN')
            a0=a0.lower()
            print("You said " + a0)
            speak.Speak('You said'+a0)
            execute(a0)
            speak.Speak('Command executed')
            cur.execute('select name from commands order by value asc')
            rr=cur.fetchall()
            aa=rr[0:9]
            print('Try using these commands:',end="")
            for ii in range(3):
                kk=random.choice(aa)
                print(kk[0],end=", ")
                aa.remove(kk)
            time.sleep(2)
            print()
            fn()
        except sr.UnknownValueError:
            speak.Speak("I did't get any value. Start speaking when I am listening.")
            fn()
    def commands():
        print('This is the list of available commands')
        speak.Speak('This is the list of available commands')
        a1=['move to position (x,y); maximum value:'+str(pyautogui.size())+'you present mouse location is:'+str(pyautogui.position()),'move up( specify the amount like move up by 100)','move down( move down by 100)','move right( move right by 100)','move left( move left by 100)']
        a8=['drag to position (x,y); maximum value:'+str(pyautogui.size())+'you present mouse location is:'+str(pyautogui.position()),'drag up(specify the amount  drag up by 100)','drag down( drag down by 100)','drag right( drag right by 100)','drag left( drag left by 100)']
        a2=['right click','left click','middle click','double click']
        a3=['type( type python)']
        a4=['press (after this the name of keys like press ctrl a)']
        a5=['show commands']
        a7=['open (starts any file in your computer)','search(google serch)','quit(it will end the program)','shutdown','restart','capture image','capture screenshot']

        r=1
        a=a1+a8+a2+a3+a4+a5+a7
        for i in a:
            print(r,i)
            r+=1
            speak.Speak(i)
        
        print("Chose one(only 1 request is processed at a time) of the commands(out of these predefined ones) to execute, I would recommend to speak exact same commands in same order (part in brackets is not necessary).")
        speak.Speak('Chose one(only 1 request is processed at a time) of the commands(out of these predefined ones) to execute, I would recommend to speak exact same commands in same order (part in brackets is not necessary). I am waiting for 5 seconds before starting.')
        time.sleep(5)
        fn()
    def execute(a0):
        if check('move',a0)==True:
            a0=text2int(a0)
            num=list(map(int,re.findall(r'\d+',a0)))
            if check('up',a0)==True:
                pyautogui.moveRel(0, -(num[0]), duration=1)
                cur.execute('update commands set value=value+1 where name="move up by 100"')
                conn.commit()
            elif check('position',a0)==True:
                x,y=num[0],num[1]
                pyautogui.moveTo(x,y,1)
                cur.execute('update commands set value=value+1 where name="move to position (x,y)"')
                conn.commit()
            elif check('down',a0)==True:
                pyautogui.moveRel(0, (num[0]), duration=1)
                cur.execute('update commands set value=value+1 where name="move down by 100"')
                conn.commit()
            elif check('right',a0)==True:
                pyautogui.moveRel(num[0], 0, duration=1)
                cur.execute('update commands set value=value+1 where name="move right by 100"')
                conn.commit()
            elif check('left',a0)==True:
                pyautogui.moveRel(-(num[0]), 0, duration=1)
                cur.execute('update commands set value=value+1 where name="move left by 100"')
                conn.commit()
            else:
                print('Please recheck the recognition result(if it matches with what you spoke),try to give commands more similar to predefined ones.Sorry, I am not able to detect it.')
                speak.Speak('Please recheck the recognition result(if it matches with what you spoke),try to give commands more similar to predefined ones.Sorry, I am not able to detect it.')
                fn()
        elif check('drag',a0)==True:
            a0=text2int(a0)
            num=list(map(int,re.findall(r'\d+',a0)))
            if check('up',a0)==True:
                pyautogui.dragRel(0, -(num[0]), duration=1)
                cur.execute('update commands set value=value+1 where name="drag up by 100"')
                conn.commit()
            elif check('position',a0)==True:
                x,y=num[0],num[1]
                pyautogui.dragTo(x,y,1)
                cur.execute('update commands set value=value+1 where name="drag to position (x,y)"')
                conn.commit()
            elif check('down',a0)==True:
                pyautogui.dragRel(0, (num[0]), duration=1)
                cur.execute('update commands set value=value+1 where name="drag down by 100"')
                conn.commit()
            elif check('right',a0)==True:
                pyautogui.dragRel(num[0], 0, duration=1)
                cur.execute('update commands set value=value+1 where name="drag right by 100"')
                conn.commit()
            elif check('left',a0)==True:
                pyautogui.dragRel(-(num[0]), 0, duration=1)
                cur.execute('update commands set value=value+1 where name="drag left by 100"')
                conn.commit()
            else:
                print('Please recheck the recognition result(if it matches with what you spoke),try to give commands more similar to predefined ones.Sorry, I am not able to detect it.')
                speak.Speak('Please recheck the recognition result(if it matches with what you spoke),try to give commands more similar to predefined ones.Sorry, I am not able to detect it.')
                fn()
        elif check('click',a0)==True:
            if check('right',a0)==True or check('write',a0)==True:
                pyautogui.rightClick()
                cur.execute('update commands set value=value+1 where name="right click"')
                conn.commit()
            elif check('left',a0)==True:
                pyautogui.leftClick()
                cur.execute('update commands set value=value+1 where name="left click"')
                conn.commit()
            elif check('middle',a0)==True:
                pyautogui.middleClick()
                cur.execute('update commands set value=value+1 where name="middle click"')
                conn.commit()
            elif check('double',a0)==True:
                pyautogui.doubleClick()
                cur.execute('update commands set value=value+1 where name="double click"')
                conn.commit()
            else:
                print('Please recheck the recognition result(if it matches with what you spoke),try to give commands more similar to predefined ones.Sorry, I am not able to detect it.')
                speak.Speak('Please recheck the recognition result(if it matches with what you spoke),try to give commands more similar to predefined ones.Sorry, I am not able to detect it.')
                fn()
        elif check('type',a0)==True:
            aa=a0.split('type')
            pyautogui.typewrite(aa[-1])
            cur.execute('update commands set value=value+1 where name="type"')
            conn.commit()
        elif check('command',a0)==True:
            commands()
            cur.execute('update commands set value=value+1 where name="show commands"')
            conn.commit()
        elif check('capture',a0)==True:
            if check('image',a0)==True:
                speak.Speak('3')
                print(3)
                time.sleep(0.5)
                speak.Speak('2')
                print(2)
                time.sleep(0.5)
                speak.Speak('1')
                print(1)
                time.sleep(0.5)
                print("say Cheese")
                speak.Speak("say cheese")
                time.sleep(0.5)
                videoCaptureObject = cv2.VideoCapture(0)
                ret,frame = videoCaptureObject.read()
                cv2.imwrite("images/CapturedImage.jpg",frame)
                videoCaptureObject.release()
                cv2.destroyAllWindows()
                cur.execute('update commands set value=value+1 where name="capture image"')
                conn.commit()
            elif check('screen',a0)==True:
                pyautogui.screenshot().save(r'images/CapturedScreenshot.jpg')
                cur.execute('update commands set value=value+1 where name="capture screnshot"')
                conn.commit()
            else:
                print('Please recheck the recognition result(if it matches with what you spoke),try to give commands more similar to predefined ones.Sorry, I am not able to detect it.')
                speak.Speak('Please recheck the recognition result(if it matches with what you spoke),try to give commands more similar to predefined ones.Sorry, I am not able to detect it.')
                fn()
        
        elif check('quit',a0)==True:
            print('Thanks for using this program')
            speak.Speak('Thanks for using this program')
            time.sleep(2)
            cur.execute('update commands set value=value+1 where name="quit"')
            conn.commit()
            sys.exit()
        elif check('press',a0)==True:
            cur.execute('update commands set value=value+1 where name="press"')
            conn.commit()
            a0=a0.split()
            z=checkf(a0)
            for i in z:
                pyautogui.keyDown(i)
            for j in z:
                pyautogui.keyUp(i)
        elif check('shut',a0)==True:
            speak.Speak('Shutting down')
            cur.execute('update commands set value=value+1 where name="shutdown"')
            conn.commit()
            cur.close()
            conn.close()
            os.system('shutdown /s /t 1')
            exit()
        elif check('restart',a0)==True:
            speak.Speak('Restarting')
            cur.execute('update commands set value=value+1 where name="restart"')
            conn.commit()
            cur.close()
            conn.close()
            os.system('shutdown /r /t 1')
            exit()
        elif check('search',a0)==True:
            cur.execute("update commands set value=value+1 where name='search'")
            conn.commit()
            at=a0.split()
            aaa=at.index('search')
            at=at[aaa+1:]
            at="+".join(at)
            run=google+at
            webbrowser.get(path).open(run)
        elif check('open',a0)==True:
            cur.execute('update commands set value=value+1 where name="open calculator"')
            conn.commit()
            pyautogui.press('win')
            at=a0.split()
            aaa=at.index('open')
            at=at[aaa+1:]
            pyautogui.typewrite(' '.join(at))
            pyautogui.press('enter')
        else:
            print("We haven't received any command similar to the command list, please try again")
            speak.Speak("We haven't received any command similar to the command list, please try again")
            fn()
        
    cur.execute('select name from commands order by value desc')
    rr=cur.fetchall()
    aa=rr[0:9]
    print('Try using these commands:',end=" ")
    speak.Speak('Try using these commands')
    for ii in range(3):
        kk=random.choice(aa)
        print(kk[0],end=", ")
        speak.Speak(kk[0])
        aa.remove(kk)
        time.sleep(0.5)
    print('for curated list of all commands use the command "show commands"')
    speak.Speak('for curated list of commands use the command "show commands"')
    fn()
except OSError:
    print('please check if your microphone is connected')
    speak.Speak('please check if your microphone is connected')
