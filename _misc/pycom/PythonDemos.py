import sys
import ctypes
import win32com
import string
from ctypes import c_int, WINFUNCTYPE, windll
from ctypes.wintypes import HWND, LPCSTR, UINT
import win32com.client
from win32com.server.exception import Exception
import winerror

prototype = WINFUNCTYPE(c_int, HWND, LPCSTR, LPCSTR, UINT)
paramflags = (1, "hwnd", 0), (1, "text", "Hi"), (1, "caption", None), (1, "flags", 0)
MsgBox = prototype(("MessageBoxA", windll.user32), paramflags)

class Class2:
    _public_methods_ = [ 'test' ]
    _reg_progid_ = "PythonDemos.Class2"
    # NEVER copy the following ID
    # Use "print pythoncom.CreateGuid()" to make a new one.
    _reg_clsid_ = "{41E24E95-D45A-11D2-852C-204C4F4F5021}"

    def test(self):
        #MsgBox(0,"in PythonDemos.Class2.test()")
        return 40
        
class PythonUtilities:
    _public_methods_ = [ 'SplitString', 'CallVB', 'getPyObj', 'Exec' ]
    _reg_progid_ = "PythonDemos.Utilities"
    # NEVER copy the following ID
    # Use "print pythoncom.CreateGuid()" to make a new one.
    _reg_clsid_ = "{41E24E95-D45A-11D2-852C-204C4F4F5020}"

    def __init__(self):
        self.dict = {}    

    def SplitString(self, val, item=None):
        #file = open('c:\\testfile.txt','w') 
        #file.write(str(sys.modules))
        #file.close() 
        if item != None: item = str(item)
        return string.split(str(val), item)

    def CallVB(self,lpFN):
        #MsgBox(0,"in python vb callback address = " + hex(lpFN))
        prototype = ctypes.WINFUNCTYPE( ctypes.c_int, ctypes.c_int, ctypes.c_int) # retval, arg1, arg2
        #MsgBox(0,"1")
        vbFunc = prototype(int(lpFN))   
        #MsgBox(0,"2")
        buf = ctypes.create_string_buffer('Hello, World')
        #MsgBox(0,"3")
        rv = vbFunc(ctypes.addressof(buf) , 0x11223344)
        #MsgBox(0,"back in python return value was: " + str(rv))
        return rv + 1

    def getPyObj(self): 
        
        #MsgBox(0,"in getPyObj")
        
        try:
            x = win32com.client.Dispatch("PythonDemos.Class2")
            #x = win32com.client.Dispatch("Scripting.FileSystemObject")
        except Exception as e:
            MsgBox(0,"Caught error in CreateObject: " + str(e))
            return -1
        
        #MsgBox(0,str(x))
        
        return x
    
    def Exec(self, exp):
        """Execute a statement.
        """
        if type(exp) not in [str, unicode]:
            raise Exception(desc="Must be a string",scode=winerror.DISP_E_TYPEMISMATCH)
        exec str(exp) in self.dict    

# Add code so that when this script is run by
# Python.exe, it self-registers.
if __name__=='__main__':
    print "Registering COM server"
    import win32com.server.register
    win32com.server.register.UseCommandLine(PythonUtilities)
    win32com.server.register.UseCommandLine(Class2)
    print "syntax error check"