import logging.handlers
import smtplib
import time
from email.message import EmailMessage
import servicemanager
import socket
import sys
import win32event
import win32service
import win32serviceutil
import datetime
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

"""
Setting up logger parameters. 
Windows path must be included to log data when running as a service
Therefore we use logging.handlers here.
"""

log_file_path = "C:\\Users\\Username\\SomeOtherDirectory\\data.log"
logger = logging.getLogger("Logger")
logger.setLevel(logging.INFO)
handler = logging.handlers.RotatingFileHandler(log_file_path)
logger.addHandler(handler)


def send_email(subject, content):
    email = 'YourEmail'
    pwd = 'YourPassword'

    message = EmailMessage()
    message['Subject'] = subject
    message['From'] = email
    message['To'] = 'your.email@something.com'
    message['CC'] = 'your.buddy@something.com'
    message.set_content(content)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(email, pwd)
        smtp.send_message(message)

    return


def service_started():
    logger.info(f'The service has started at: {datetime.datetime.now().replace(microsecond=0)}')
    subject = f"The service has started"
    content = F"This is the email informing about windows service start"
    send_email(subject, content)
    return


def service_stopped():
    logger.info(f'The service has stopped at: {datetime.datetime.now().replace(microsecond=0)}')
    subject = f"The service has been stopped!"
    content = F"This is the email informing about windows service being stopped"
    send_email(subject, content)
    return


def establish_connection():
    engine = create_engine("postgresql+psycopg2://postgres:test@localhost/your_database")

    conn = engine.connect()
    Session = sessionmaker()
    Session.configure(bind=engine)
    session = Session()

    return session, conn


def some_function_you_use_with_database(session, conn):
    pass


def some_function_you_use_results_from_db_and_update(session, conn):
    pass


class TestService(win32serviceutil.ServiceFramework):
    _svc_name_ = 'Test_Service'
    _svc_display_name_ = 'Test_Service'
    _svc_description_ = 'This is a test service'

    def __init__(self, args):
        logger.info('Init')
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)

    def start(self):
        logger.info('Start')
        self.isrunning = True
        service_started()

    def stop(self):
        logger.info('Stop')
        self.isrunning = False
        service_stopped()

    def SvcStop(self):
        self.stop()
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        logger.info('Run')
        self.start()
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, '')
                              )
        self.main()

    def main(self):
        logger.info('Main function called')
        while self.isrunning:
            subject = F"Some service initiated..."
            content = F"Service is up and running...."
            send_email(subject, content)
            session, conn = establish_connection()
            results = some_function_you_use_with_database(session, conn)
            some_function_you_use_results_from_db_and_update(results, session, conn)
            time.sleep(900)


if __name__ == '__main__':
    if len(sys.argv) == 1:
        logger.info('Calling handle command line')
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(TestService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(TestService)
