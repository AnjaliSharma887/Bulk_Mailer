import smtplib
import socket
socket.getaddrinfo('localhost', 8080)

def email_send_func(to_,sub_,msg_,email_,pw_):
    print(to_,sub_,msg_,email_,pw_)
    s=smtplib.SMTP("smtp.gmail.com",587)
    s.starttls()
    s.login(email_,pw_)
    msg_="Subject:{}\n\n{}".format(sub_,msg_)
    s.sendmail(email_,to_,msg_)
    x=s.ehlo()
    if x[0]==250:
        return "s"
    else:
        return "f"
    s.close()


    