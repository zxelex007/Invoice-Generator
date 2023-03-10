import gspread
import speech_recognition as sr

sa = gspread.service_account()
sh = sa.open("Invoice1")
ws = sh.worksheet("Sheet1")
invoice_no = input("Enter your Invoice No:\n")
ws.update("F10",invoice_no) 

                        

def greet():
    ask = input("Press: 1) 1 to insert Item Name\n       2) 2 to insert Quantity\n       3) 3 to insert Price")
    if ask == "1":
        ItemName()
    elif ask == '2':
        Quantity()
    elif ask == '3':
        Price()

def ItemName():
    r=sr.Recognizer()
    i = 18
    j = 1
    for item in range(100):
        cell = f"C{str(i)}"
        
        check = input("Press: q to quit\n       c to continue")
        if check == "c":
            with sr.Microphone() as source:
                print("Item Name: listening...")
                r.pause_threshold = 1
                audio=r.listen(source)  
                try:
                    print("Recognizing...")  
                    Item_name = r.recognize_google(audio,language='en-in')    
                    print(f"User said:{Item_name.upper()}\n")
                    conf = input("Press: (y): if the text is correctly recognized\n (n) if the text is not recognized correctly")
                    if conf != 'n':
                        try:
                            iname = Item_name.remove(" ")
                        except:
                            iname = Item_name
                            pass
                        
                        
                        ws.update(cell,iname.upper())
                        ws.update(f"B{i}",j)
                        j+=1
                    else: pass
                except: 
                    print("Couldn't able to identify your voice, say that again please...")
        else: break
def Cname():
    r=sr.Recognizer()
    i = 18
    with sr.Microphone() as source:
        print("Company Name: listening...")
        r.pause_threshold = 1
        audio=r.listen(source)  
        try:
            print("Recognizing...")  
            cname = r.recognize_google(audio,language='en-in')    
            print(f"User said:{cname.upper()}\n")
            conf = input("Press: (y): if the text is correctly recognized\n (n) if the text is not recognized correctly")
            if conf != 'n':
                iname = cname
                ws.update("B10",iname.upper())
                        
            else: pass
        except: 
            print("Couldn't able to identify your voice, say that again please...")
        
def Address():
    r=sr.Recognizer()
    i = 18
    
    
        
    with sr.Microphone() as source:
        print("Address: listening...")
        r.pause_threshold = 1
        audio=r.listen(source)  
        try:
            print("Recognizing...")  
            add = r.recognize_google(audio,language='en-in')    
            print(f"User said:{add.upper()}\n")
            conf = input("Press: (y): if the text is correctly recognized\n (n) if the text is not recognized correctly")
            if conf != 'n':
                iname = add
                ws.update("B10",iname.upper())
                        
            else: pass
        except: 
            print("Couldn't able to identify your voice, say that again please...")
def Quantity():
    r=sr.Recognizer()
    i = 18
    for item in range(100):
        cell = f"E{str(i)}"
        check = input("Press: q to quit\n       c to continue")
        if check == "c":
            with sr.Microphone() as source:
                print("Item Name: listening...")
                r.pause_threshold = 1
                audio=r.listen(source)  
                try:
                    print("Recognizing...")  
                    Quantity = r.recognize_google(audio,language='en-in')    
                    print(f"User said:{Quantity.upper()}\n")
                    conf = input("Press: (y): if the text is correctly recognized\n       (n) if the text is not recognized correctly")
                    if conf == 'y':
                        try:
                            iname = Quantity.remove(" ")
                        except:
                            iname = Quantity
                            pass
                        ws.update(cell,iname.upper())
                        i+=1
                    else: pass
                except: 
                    print("Couldn't able to identify your voice, say that again please...")
        else: break
def Price():
    r=sr.Recognizer()
    i = 18
    for item in range(100):
        cell = f"F{str(i)}"
        check = input("Press: q to quit\n       c to continue")
        if check == "c":
            with sr.Microphone() as source:
                print("Price: listening...")
                r.pause_threshold = 1
                audio=r.listen(source)  
                try:
                    print("Recognizing...")  
                    Price = r.recognize_google(audio,language='en-in')    
                    print(f"User said:{Price.upper()}\n")
                    conf = input("Press: (y): if the text is correctly recognized\n       (n) if the text is not recognized correctly")
                    if conf == 'y':
                        try:
                            iname = Price.remove(" ")
                        except:
                            iname = Price
                            pass
                        ws.update(cell,iname.upper())
                        i+=1
                    else: pass
                except: 
                    print("Couldn't able to identify your voice, say that again please...")
        else: break
def Tax():
    r=sr.Recognizer()
    i = 18
    for item in range(100):
        cell = f"G{str(i)}"
        check = input("Press: q to quit\n       c to continue")
        if check == "c":
            with sr.Microphone() as source:
                print("Item Name: listening...")
                r.pause_threshold = 1
                audio=r.listen(source)  
                try:
                    print("Recognizing...")  
                    tax = r.recognize_google(audio,language='en-in')    
                    print(f"User said:{tax.upper()}\n")
                    conf = input("Press: (y): if the text is correctly recognized\n       (n) if the text is not recognized correctly")
                    if conf == 'y':
                        try:
                            iname = tax.remove(" ")
                        except:
                            iname = tax
                            pass
                        ws.update(cell,iname.upper())
                        i+=1
                    else: pass
                except: 
                    print("Couldn't able to identify your voice, say that again please...")
        else: break

    r=sr.Recognizer()
    i = 18
    for item in range(100):
        cell = f"H{str(i)}"
        check = input("Press: q to quit\n c to continue")
        if check == "c":
            with sr.Microphone() as source:
                print("Item Name: listening...")
                r.pause_threshold = 1
                audio=r.listen(source)  
                try:
                    print("Recognizing...")  
                    amount = r.recognize_google(audio,language='en-in')    
                    print(f"User said:{amount.upper()}\n")
                    conf = input("Press: (y): if the text is correctly recognized\n       (n) if the text is not recognized correctly")
                    if conf == 'y':
                        iname = amount.remove(" ")
                        ws.update(cell,iname.upper())
                        i+=1
                    else: pass
                except: 
                    print("Couldn't able to identify your voice, say that again please...")
        else: break
Cname()
Address()

greet()
