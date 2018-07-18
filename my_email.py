#!/usr/bin/env python
#coding:utf-8

#this is outlook email test recv and send by xx@xxxx.com
#reference http://www.snb-vba.eu/VBA_Outlook_external_en.html#L_15.2.1


import time
import threading
import logging
import platform
import smtplib
import poplib
from email import encoders
from email.header import Header
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.utils import parseaddr, formataddr
from email.parser import Parser


log_file_name = "mail.log"


#获取设备类型
if "windows" == platform.system():
    #sys.path.append(os.path.abspath(".") + "\\libs")
    import win32com.client as win32
    print "This platform is windows!"
else:
    #sys.path.append(os.path.abspath(".") + "/libs")
    print "This platform is linux!"


#send email by outlook(just in win)
def send_email_by_outlook():
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    receicer = ['xx@xxxx.com']
    # sender = ['xx@xxxx.com']
    # mail.sender = sender[0]
    mail.To = receicer[0]
    mail.Cc = receicer[0]
    mail.Subject = 'This is test'
    mail.Body = "hello,python!"
    mail.Attachments.Add('night.jpg')
    mail.Send()
    return

#recv email by outlook(just in win)
def receive_emial_by_outlook():
    outlook = win32.Dispatch('outlook.application')
    #打印收件箱邮件数量
    mail_count = outlook.GetNamespace("MAPI").GetDefaultFolder(6).Items.Count
    while True:
        new_mail_count = outlook.GetNamespace("MAPI").GetDefaultFolder(6).Items.Count
        if new_mail_count > mail_count:
            for mail_index in range(mail_count,new_mail_count):
                mail = outlook.GetNamespace("MAPI").GetDefaultFolder(6).Items(mail_index + 1)
                logging.debug("From:" + mail.sender() + "\nTo:" + mail.To + "\nCC:" + mail.Cc + "\nSubject:" + mail.Subject + "\nContent:\n" + mail.Body)
                mail_count = new_mail_count
        else:
            if mail_count > new_mail_count:
                mail_count = new_mail_count
            logging.debug("There is no new mail!mail_count:%d new_mail_count:%d" %(mail_count,new_mail_count))

        time.sleep(3)
    return



#format mail address
def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(),addr))


#send mail by smtp
def send_email_by_smtp():
	ser_ip = "smtp.163.com"
	ser_port = 25
        
        mail_sender_count = raw_input("Mail Count:")
        mail_sender_passwd = raw_input("Mail passwd:")
        mail_recver_count = raw_input("Mail recver count:")
        subject = raw_input("subject:") 
        content = raw_input("mail content:")

        '''
        #发送纯文本文件
        msg = MIMEText(subject, "plain", "utf-8")
        msg['From'] = _format_addr("小可爱 <%s>" %mail_sender_count)
        msg['To'] = mail_recver_count
        msg['Subject'] = Header(subject, 'utf-8').encode()
        '''

        #发送带附件邮件
        msg = MIMEMultipart()
        msg["From"] = _format_addr("小可爱 <%s>" % mail_sender_count)
        msg["To"] = mail_recver_count;
        msg["Subject"] = Header(subject, "utf-8").encode()
        msg.attach(MIMEText(content, "plain", "utf-8"))



        try:
            f = open("night.jpg", "rb")
            mime = MIMEBase("image", "jpg", filename = "night.jpg")
            #mime = MIMEBase("file", "log", filename = "20--v003.log")
            mime.add_header("Content-Disposition","attachment",filename="night.jpg")
            mime.add_header("Content-ID","<0>")
            mime.add_header("X-Attachment-ID","0")
            mime.set_payload(f.read())
            encoders.encode_base64(mime)
            msg.attach(mime)
            f.close()
            print "Attach file successful!"

        except:
            print "Attach file fail!"


        try:
            smtpobj = smtplib.SMTP(ser_ip, ser_port)
            #smtpobj.set_debuglevel(1)
	    smtpobj.login(mail_sender_count, mail_sender_passwd)
            print "login mail MTA successful!"
        except:
            print "login mail MTA fail!"
            return

	
        try:
            smtpobj.sendmail(mail_sender_count, mail_recver_count, msg.as_string())
            print "send mail successful!"
        except:
            print "send mail fail!"
	
	smtpobj.quit()
	return


#decode mail_str
def decode_str(str):
    value,charset = decode_header(str)[0]
    if charset:
        value = value.decode(charset)
    return value

#guess content's charset
def guess_charset(content):
    charset = content.get_charset()
    if charset is None:
        content_type = content.get("Content-Type","").lower()
        pos = content_type.find("charset=")
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset


#输出邮件内容
def print_info(mail):
    for header in ["Subject", "From", "To"]:
        value = mail.get(header, "")
        if header == "Subject" and value:
            print "Subject:%s"%decode_str(value)
        if header == "From" and value:
            name_str,addr = parseaddr(value)
            name = decode_str(name_str)
            print "From:%s %s"%(name,addr)
        if header == "To" and value:
            name_str,addr = parseaddr(value)
            name = decode_str(name_str)
            print "To:%s %s"%(name,addr)
        
    if (mail.is_multipart()):
        parts = mail.get_payload()
        for n,part in enumerate(parts):
            print "------------part %d start-------------"%n
            print_info(part)
            print "------------part %d end---------------\n\n"%n
    else:
        content_type = mail.get_content_type()
        if content_type == "text/plain":
            content = mail.get_payload(decode=True)
            charset = guess_charset(mail)
            if charset:
                content = content.decode(charset)
            print "  content:%s"%content
        else:
            print "  Attachment:%s"%content_type 
    
    return





#recv email by pop3
def recv_email_by_pop3():
    pop3_ser = "pop3.163.com"
    pop3_ser_port = 110
    mail_sender_count = raw_input("Mail Count:")
    mail_sender_passwd = raw_input("Mail passwd:")

    try:
        pop3obj = poplib.POP3(pop3_ser)
        #pop3obj.set_debuglevel(1)
        print pop3obj.getwelcome().decode("utf-8")
        print "login mail pop3 server successful!"
    except:
        print "login mail pop3 server fail!"
        return
    
    try:
        #身份认证
        pop3obj.user(mail_sender_count)
        pop3obj.pass_(mail_sender_passwd)
        print "Mail server auth success!"
    except:
        print "Mail server auth fail!" 
        return

    #stat返回邮件数量和占用空间
    print "Messages:%s Size:%s" % pop3obj.stat()
    rest_pace,mails,octets = pop3obj.list()
    #print mails,rest_pace,octets
    index = len(mails)
    rest_pace,lines,octets = pop3obj.retr(index)
    mail_content = b'\r\n'.join(lines).decode("utf-8")
    mail = Parser().parsestr(mail_content)
    
    print_info(mail)

    return




def log_config():
    log_format = "%(asctime)s - %(levelname)s - %(funcName)s:%(lineno)s: %(message)s"
    logging.basicConfig(filename = log_file_name, level = logging.DEBUG, format = log_format)
    '''
    logging.debug("this is log debug test!")
    logging.info("this is log info test!")
    logging.warning("this is log warning test!")
    logging.error("this is log error test!")
    logging.critical("this is log critical test!")
    '''


def main():
    #加载log配置
    log_config()
    
    #发送outlook邮件
    #send_email_by_outlook()

    #启动outlook收邮件线程
    #recv_thread = threading.Thread(target = receive_emial_by_outlook())
    #recv_thread.start()
    
    #send_email_by_smtp()
    recv_email_by_pop3()

    return


if __name__ == '__main__':
    main()
