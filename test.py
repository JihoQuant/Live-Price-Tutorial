import win32com.client
import ctypes
import pythoncom

CREON_STATUS = win32com.client.Dispatch('CpUtil.CpCybos')

def check_creon():
    if ctypes.windll.shell32.IsUserAnAdmin():
        print('관리자 권한으로 실행됨')
    else:
        print('관리자 권한이 아님')
        return False

    if CREON_STATUS.IsConnect == 0:
        print('CREON 연결 안 됨')
        return False

    return True

class LivePriceEvent:
    def set_client(self, client, callback):
        self.client = client
        self.callback = callback

    def OnReceived(self):
        code = self.client.GetHeaderValue(0)
        name = self.client.GetHeaderValue(1)
        price = self.client.GetHeaderValue(13)
        time = self.client.GetHeaderValue(18)

        print(code, name, price, time)


class LivePriceReceiver:
    def __init__(self):
        self.client = win32com.client.Dispatch('DsCbo1.StockCur')
        self.is_subscribe = False

    def subscribe(self, code='A078930'):
        if self.is_subscribe:
            self.unsubscribe()

        print(code)
        self.client.SetInputValue(0, code)

        handler = win32com.client.WithEvents(self.client, LivePriceEvent)
        handler.set_client(self.client, self.print_values)
        handler.client.Subscribe()

        self.is_subscribe = True

    def print_values(self, code, name, price):
        print(code, name, price)

    def unsubscribe(self):
        if self.is_subscribe:
            self.client.Unsubscribe()
        self.is_subscribe = False


check_creon()
receiver = LivePriceReceiver()
receiver.subscribe("A005930")

try:
    while True:
        pythoncom.PumpWaitingMessages()
except:
    pass
finally:
    receiver.unsubscribe()
