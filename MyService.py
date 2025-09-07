import win32serviceutil
import win32service
import win32event
import servicemanager
import time


class MyService(win32serviceutil.ServiceFramework):
    _svc_name_ = "MyPythonService"
    _svc_display_name_ = "My Python Test Service"
    _svc_description_ = "A simple Python service that writes logs every 5 seconds."

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.running = True

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.running = False

    def SvcDoRun(self):
        servicemanager.LogInfoMsg("MyPythonService is starting...")
        self.main()

    def main(self):
        while self.running:
            try:
                with open("C:\\service_log.txt", "a") as f:
                    f.write("Service is running...\n")
            except Exception as e:
                servicemanager.LogErrorMsg(str(e))
            time.sleep(5)


if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(MyService)
