#!/usr/bin/env python
#coding:utf-8

#this is outlook email test recv and send by xx@xxxx.com
#reference http://www.snb-vba.eu/VBA_Outlook_external_en.html#L_15.2.1


import time
import threading
import logging
#import win32com.client as win32
from email.mime.text import MIMEText
import smtplib

msg = MIMEText("hello, python!", "plain", "utf-8")

'''
#send email by outlook
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
    mail.Attachments.Add('D:\\1.jpg')
    mail.Send()
    return

#recv email by outlook
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
'''

#send mail by smtp
def send_email_by_smtp():
	ser_ip = "smtp.163.com"
	ser_port = 25
        from_addr = input("mail_addr:")
        passwd = input("passwd:")
        to_addr = input("to_addr:")
	
	smtpobj = smtplib.SMTP(ser_ip, ser_port)
	#smtpobj.set_debuglevel(1)
	login_log = smtpobj.login(from_addr, passwd)
        print login_log
	#smtpobj.sendmail(from_addr, to_addr, msg.as_string())
	
	smtpobj.quit()
	return




def log_config():
    log_format = "%(asctime)s - %(levelname)s - %(funcName)s:%(lineno)s: %(message)s"
    logging.basicConfig(filename = "D:\\email.log", level = logging.DEBUG, format = log_format)
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
    
    send_email_by_smtp()
    
    #发送outlook邮件
    #send_email_by_outlook()

    #启动收邮件线程
    #recv_thread = threading.Thread(target = receive_emial_by_outlook())
    #recv_thread.start()
    
    return


if __name__ == '__main__':
    main()
