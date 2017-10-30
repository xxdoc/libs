import win32con, win32api, win32gui, ctypes, ctypes.wintypes, sys, struct
from array import array

# pip install pypiwin32

def help():
    print "%d arguments given: %s" % (len(sys.argv), str(sys.argv))
    print "no args = listen for WM_COPYDATA messages"
    print "2 args: hwnd msg"
    print "other = this help screen"

WM_DISPLAY_TEXT = 3

class COPYDATASTRUCT(ctypes.Structure):
    _fields_ = [
        ('dwData', ctypes.wintypes.LPARAM),
        ('cbData', ctypes.wintypes.DWORD),
        ('lpData', ctypes.c_void_p)
    ]
PCOPYDATASTRUCT = ctypes.POINTER(COPYDATASTRUCT)
    
def sendCommand(w, message):
    print "sending msg'%s' to %s" % (message, w)
    isWindow = win32gui.IsWindow(int(w))
    if isWindow != 1:
        print "not a valid hwnd"
        return
    CopyDataStruct = "IIP"
    char_buffer = array('c', message)
    char_buffer_address = char_buffer.buffer_info()[0]
    char_buffer_size = char_buffer.buffer_info()[1]
    cds = struct.pack(CopyDataStruct, WM_DISPLAY_TEXT, char_buffer_size, char_buffer_address)
    v = win32gui.SendMessage(int(w), win32con.WM_COPYDATA, 0, cds)
    print "done retval = %d" % v
    
def recvCommand(msg):
    print "in recvCommand: %s" % msg
    ary = msg.split("=")
    if len(ary) == 2:
        if ary[0] == "PINGME":
            sendCommand(int(ary[1]), "Why hum-diddly dog it dun works roscoe")
    
class Listener:
    # https://stackoverflow.com/questions/5249903/receiving-wm-copydata-in-python   
    def __init__(self):
        message_map = {
            win32con.WM_COPYDATA: self.OnCopyData
        }
        wc = win32gui.WNDCLASS()
        wc.lpfnWndProc = message_map
        wc.lpszClassName = 'MyWindowClass'
        hinst = wc.hInstance = win32api.GetModuleHandle(None)
        classAtom = win32gui.RegisterClass(wc)
        self.hwnd = win32gui.CreateWindow (
            classAtom,
            "win32gui test",
            0,
            0, 
            0,
            win32con.CW_USEDEFAULT, 
            win32con.CW_USEDEFAULT,
            0, 
            0,
            hinst, 
            None
        )
        print "Python listening for WM_COPYDATA on hwnd = %d" % self.hwnd
        
    def OnCopyData(self, hwnd, msg, wparam, lparam):
        print "Copy data msg received! hwnd=%d msg=0x%x wparam=0x%x lparam=0x%x" % (hwnd,msg,wparam,lparam)
        pCDS = ctypes.cast(lparam, PCOPYDATASTRUCT)
        if pCDS.contents.dwData != WM_DISPLAY_TEXT:
            print "Not WM_DISPLAY_TEXT dwData=%d cbData=0x%x lpData=0x%x" % (pCDS.contents.dwData,pCDS.contents.cbData,pCDS.contents.lpData)
            return
        
        print "WM_DISPLAY_TEXT received cbData=0x%x lpData=0x%x" % (pCDS.contents.cbData,pCDS.contents.lpData)
        msg = ctypes.string_at(pCDS.contents.lpData)
        recvCommand(msg)
        return 1





if len(sys.argv) == 1:
    l = Listener()
    win32gui.PumpMessages()        
else:
    if len(sys.argv) != 3:
        help()
    else:
        h = sys.argv[1]
        msg = sys.argv[2]
        sendCommand(h,msg)

