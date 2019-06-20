import smtplib, ssl
import imaplib
import email
import speech_recognition as sr
import win32com.client as wincl
speak=wincl.Dispatch("SAPI.SpVoice")

port = 587  # For starttls
smtp_server = "smtp.gmail.com"
speak.Speak("welcome") 

speak.Speak("enter 1 for sending mail and 2 for receiving mail")
r=sr.Recognizer()
with sr.Microphone() as source:
    audio=r.listen(source)
    try:
        op = r.recognize_google(audio, language='en-IN')
        print("you chose "+op)
    except:
        print("google can't recognize your voice")

if "one" or "1" in op:
    speak.Speak("enter your mail id") 
    se=input("enter your mail id:- ")
    print(se)

    speak.Speak("enter your password") 
    password=input("enter password:- ")
    print(password)

    speak.Speak("enter receiver's mail id")
    re=input("enter receiver's mail id:-")
    print(re)

    speak.Speak("enter message") 
    cr=sr.Recognizer()
    with sr.Microphone() as source:
        audio=cr.listen(source)
        try:
            message = cr.recognize_google(audio, language='en-IN')
            print(message)
        except:
            print("google can't recognize your voice")

    context = ssl.create_default_context()
    with smtplib.SMTP(smtp_server, port) as server:
        server.ehlo()  
        server.starttls(context=context)
        server.ehlo()  
        server.login(se,password)
        server.sendmail(se, re, message)
    
    print("mail sent")
    speak.Speak("your mail has been sent")
    server.close()

#elif "two" or "2" in op:
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(se, password)
    mail.list()
    mail.select("inbox")
    result, data = mail.uid('search', None, "ALL") 
    latest_email_uid = data[0].split()[-1]
    result, data = mail.uid('fetch', latest_email_uid, '(RFC822)')
    raw_email = data[0][1] 
    msg = email.message_from_bytes(raw_email)
    print ("From:-"+msg['From'])
    speak.Speak("Sender is "+msg['From'])
    speak.Speak("Message is")
    print (msg.get_payload(decode=True))
    speak.Speak(msg.get_payload(decode=True))
    #speak.Speak("Body of mail is" +msg['Body'])
    mail.close()
    mail.logout()
    
    
    