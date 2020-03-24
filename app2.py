import os.path
from os import path
import time
import win32com.client as comclt
import win32gui
import threading
import subprocess
import random
from os import system, name 

print(" ╔═════════════════════════════════════╗")
print(" ║                                     ║")
print(" ║ █████ █   █ ████   ███   ███   ███  ║")
time.sleep(0.03)
print(" ║   █   █   █ █   █ █   █ █   █ █   █ ║")
time.sleep(0.03)
print(" ║   █   █   █ █   █ █   █ █  ██ █   █ ║")
time.sleep(0.03)
print(" ║   █    █ █  ████   ███  █ █ █  ███  ║")
time.sleep(0.03)
print(" ║   █     █   █     █   █ ██  █ █   █ ║")
time.sleep(0.03)
print(" ║   █     █   █     █   █ █   █ █   █ ║")
time.sleep(0.03)
print(" ║   █     █   █      ███   ███   ███  ║")
time.sleep(0.03)

print(" ╟─────────────────────────────────────╢")
print(" ║ Creator by TYP808 © 2019            ║")
print(" ╚═════════════════════════════════════╝")
print("")

#====================================================

wsh = comclt.Dispatch("WScript.Shell")

nfile = "num.txt"
cfg_file = "config.ini"

def_online = "20-40"
def_waiting = "30-120"
sec_min = 3
sec_max = 1800
app_name = "Skype"
skype_path = "C:/Program Files (x86)/Microsoft/Skype for Desktop/Skype.exe"

    
def loadConfig():
    
    global _cfg
    
    if(path.isfile(cfg_file)):
    
        file = open(cfg_file, "r")
        _cfg={}

        for str in file:

            str = str.rstrip('\r\n')
            
            if(len(str) > 0):
                
                part = str.split("=")
                cfg_name = part[0].replace(" ", "")
                
                if(len(part[1].split('"')[0].replace(" ", "")) == 0):
                
                    cfg_val = part[1].split('"')[1]
                    _cfg[cfg_name] = cfg_val
                    
                elif(len(part[1].split("'")[0].replace(" ", "")) == 0):
                
                    cfg_val = part[1].split("'")[1]
                    _cfg[cfg_name] = cfg_val
                    
                elif(len(part[1].split("`")[0].replace(" ", "")) == 0):
                
                    cfg_val = part[1].split("`")[1]
                    _cfg[cfg_name] = cfg_val
                    
                    
        setConfig()
                    

    else:
        print(" Файл " + cfg_file + " отсутствует!")
        
        flag = False
        while flag is False:

            answer = input(" Хотите чтобы программа сама создала файл " + cfg_file + "? (y/n): ")
            if(answer == "y" or answer == ""):
            
                flag = True

                intQuest()       
                
            elif(answer == "n"):
                exit()
            else:
                print(" Дайте ответ в формате (y/n)!")       
        



def setConfig():

    global def_online
    global def_waiting
    global app_name
    global skype_path
    
    def_online = _cfg["online"]
    def_waiting = _cfg["waiting"]
    app_name = _cfg["app_name"]
    skype_path = _cfg["skype_path"]
    
    checkNumFile()

def intQuest():

    global def_online
    global def_waiting
    
    param = ""
    print("")

    print("================== Настройка интервала ==================")

    print("")

    flag = False
    while flag is False:
        
        answer = input(" 1. █ Введите диапазон ожидания после набора в секундах (по умолчанию " + def_online + " нажмите Enter): ")
        print("    │")
        if(not answer == ""):
                    
            array = answer.split("-")
            array_len = len(array)

            if(array_len > 1 and array_len < 3): #Размер массива
                    
                if(array[0].isdigit() and array[1].isdigit()): #Число или нет
                        
                    min = int(array[0])
                    max = int(array[1])
                    
                    if(min < max or min == max): #Min меньше Max больше
                    
                        if((min > sec_min-1 and min < sec_max+1) and (max > sec_min-1 and max < sec_max+1)): #Ограничения по величине
                            
                            def_online = answer
                            print("    └─█ OK! Диапазон ожидания на линии " + def_online + " сек. ")
                            param = "online = " + '"' + def_online + '"\n'
                            flag = True
                            
                        else:
                            flag = False  
                            #print("Error 4")
                            
                    else:
                        flag = False  
                        #print("Error 3")
                        
                else:
                    flag = False  
                    #print("Error 2")
                     
            else:
                flag = False  
                #print("Error 1")

            if(flag == False):
                print("    █ !!! Ошибка введите диапазон в формате min-max !!!")
                print("    │")

        else:
            flag = True
            print("    └─█ OK! Установлено по умолчанию " + def_online + ".")
            param = "online = " + '"' + def_online + '"\n'


    print("")



    flag = False
    while flag is False:

        answer = input(" 2. █ Введите диапазон ожидания перед новым звонком в секундах (по умолчанию " + def_waiting + " нажмите Enter): ")
        print("    │")
        if(not answer == ""):
                    
            array = answer.split("-")
            array_len = len(array)

            if(array_len > 1 and array_len < 3): #Размер массива
                    
                if(array[0].isdigit() and array[1].isdigit()): #Число или нет
                        
                    min = int(array[0])
                    max = int(array[1])
                    
                    if(min < max or min == max): #Min меньше Max больше
                    
                        if((min > sec_min-1 and min < sec_max+1) and (max > sec_min-1 and max < sec_max+1)): #Ограничения по величине
                            
                            def_waiting = answer
                            print("    └─█ OK! Диапазон ожидания перед новым звонком " + def_waiting + " сек. ")
                            param = param + "waiting = " + '"' + def_waiting + '"\n'
                            flag = True
                            
                        else:
                            flag = False  
                            #print("Error 4")
                            
                    else:
                        flag = False  
                        #print("Error 3")
                        
                else:
                    flag = False  
                    #print("Error 2")
                     
            else:
                flag = False  
                #print("Error 1")

            if(flag == False):
                print("    █ !!! Ошибка введите диапазон в формате min-max !!!")
                print("    │")

        else:
            flag = True
            print("    └─█ OK! Установлено по умолчанию " + def_waiting + ".")
            param = param + "waiting = " + '"' + def_waiting + '"\n'



    print("")
    print("===================== Настройка завершена =====================")

    print("")
    
    param = param + "app_name = " + '"' + app_name + '"\n'
    param = param + "skype_path = " + '"' + skype_path + '"\n'
    
    f = open(cfg_file, "a")
    f.write(param)
    f.close()    

    loadConfig()


def checkNumFile():

    if(path.isfile(nfile)):
        loadNumFile()
    else:

        print(" Файл " + nfile + " отсутствует!")
        
        flag = False
        while flag is False:

            answer = input(" Хотите чтобы программа сама создала файл " + nfile + "? (y/n): ")
            if(answer == "y" or answer == ""):
            
                flag = True
                f = open(nfile, "a")
                f.close()
                time.sleep(1)
                
                print(" Добавьте в файл телефонные номера в формате +79871231212. Каждый новый номер на новой строке.")
                
                os.system("notepad " + nfile)
                
                if(path.isfile(nfile)):
                    print("-------------------------------------")
                    loadNumFile()
                else:
                    print(" Произошла неизвестная ошибка!")       
                    
            elif(answer == "n"):
                exit()
            else:
                print(" Дайте ответ в формате (y/n)!")       
        

def loadNumFile():

    global numbers
    global numb_count
    
    numbers = [line.rstrip('\n') for line in open(nfile, "r")]
    numb_count = len(numbers)
    print(" Загружено номеров: " + str(numb_count))
    print("")

    answer = input(" Начать обзвон? (Жми Enter):")
    startCalling()

def startCalling():

    system('cls') 
    print("  Вызов: " + def_online + " / Пауза: " + def_waiting + " / Кол-во номеров: " + str(numb_count))
    print("╚═══════════════════════════ СТАРТ ═══════════════════════════╝")
    print("")
    global proc
    global result

    array_online = def_online.split("-")
    array_waiting = def_waiting.split("-")

    a = int(array_online[0])
    b = int(array_online[1])
    c = int(array_waiting[0])
    d = int(array_waiting[1])

    i = 0
    for number in numbers:
    
        i = i + 1
        online = int(random.randrange(a, b))
        waiting = int(random.randrange(c, d))
        
        print("╒══════ Готовность " + str(i) + " из " + str(numb_count) + ": " + number, end="\r")

        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        proc = subprocess.Popen('"' + skype_path + '" --_="skype:' + number + '"', startupinfo=si)
        
        appname = app_name
        window = win32gui.FindWindow(None, appname)
        try:
            win32gui.SetForegroundWindow(window)
            time.sleep(3)
            wsh.SendKeys("{ENTER}")
            win32gui.SetForegroundWindow(window)
            time.sleep(0.1)
            wsh.SendKeys("{ENTER}")
            win32gui.SetForegroundWindow(window)
            time.sleep(0.1)
            wsh.SendKeys("{ENTER}")
            time.sleep(1)
            wsh.SendKeys("{ESC}")
            print("├ Вызов")
            print("├ Ждем на линии: " + str(online) + " сек.")
            
            print("│")

            waitingBar(online)
            
            print("│")
            
            win32gui.SetForegroundWindow(window)
            print("├ Завершение вызова")
            time.sleep(3)
            wsh.SendKeys("^e")
            time.sleep(1)
            wsh.SendKeys("{ESC}")
            completedNumb(number)
            print("└──────────────────────────────────────────────────────────────")
            
            if(not numb_count == i):
            
                print("╒══════ Пауза перед новым вызовом: " + str(waiting) + " сек.")
                print("│")
                    
                waitingBar(waiting)
                
                print("└──────────────────────────────────────────────────────────────")
                

        except:
            print(" Window (" + appname + ") not found!")
     
        result = proc.communicate()
        
        text = str(result[0])
        errors = str(result[1])
        
        
    print("")
    print(" ═══════════════════════════ ФИНИШ ═══════════════════════════")
    print("")
    #input("Чтобы завершить, нажмите Enter ...")
    #os.system("pause")   

    flag = False
    while flag is False:

        print(" y или enter - Загрузить список номеров и запустить снова")
        print(" n - Закрыть приложение")
        answer = input(": ")
        
        if(answer == "y" or answer == ""):
            flag = True
            print("")
            loadConfig()
        elif(answer == "n"):
            exit()
        else:
            print(" Дайте ответ в формате (y/n)!")       


def waitingBar(sec):

    sleep_int = sec/100
    p = "█"
    pin = ""   
    r = 0.00000000000002

    for x in range(100):
        
        r = r + sleep_int
        v = str(r).split(".")[0]
        
        if(str((x/2)).split(".")[1] == "5"):
            pin = pin + p
        
        if(x < 9):
            sp ="  "
        elif(x > 9 and x < 99):
            sp =" "
        elif(x == 99):
            sp =""
        else:
            sp = " "


        print("├ " + str(x+1) + "% " + sp + pin + " " + v + "", end="\r")

        time.sleep(sleep_int)
        
    print()
           


def completedNumb(comp_numb):

    f = open("comp.txt", "a")
    f.write(comp_numb + '\n')
    f.close()



#===============================================

loadConfig()
