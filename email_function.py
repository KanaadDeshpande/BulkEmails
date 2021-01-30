import smtplib

def email_sent_function(to_, subject_, message_, from_, password_):
    # Creating SMTP Session for GMail
    s = smtplib.SMTP("smtp.gmail.com", 587)
    s.starttls() # Starting the transport layer
    s.login(from_, password_)
    message = "Subject: {}\n\n{}".format(subject_, message_)
    s.sendmail(from_, to_, message)
    x = s.ehlo() # Mail Result
    if x[0] == 250:
        return "s"
    else:
        return "f"
    s.close()
    
    
