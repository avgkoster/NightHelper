import smtplib,mimetypes,email,email.mime,email.mime.application,email.mime.multipart


#Отправка отчета на почту
def OtprEmail(title,froms,to,filenames,server,log,password):
    msg =  email.mime.multipart.MIMEMultipart()
    msg['Subject']=title
    msg['From']=froms
    msg['To']=to

    #Excel файл
    filename=filenames
    fop = open(filename,'rb')
    att =  email.mime.base.MIMEBase('application','vnd.ms-excel')
    att.set_payload(fop.read())
    email.encoders.encode_base64(att)
    att.add_header('Content-Dispostion','attachment',filename=filenames)
    msg.attach(att)

    #Отправка
    server = smtplib.SMTP(server)
    server.starttls()
    server.login(log,password) 
    server.sendmail(froms,to,msg.as_string())
    server.quit()
