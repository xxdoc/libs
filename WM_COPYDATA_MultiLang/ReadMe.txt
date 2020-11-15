
Last Update: 10.6.17


this project depends on the ActiveX control spSubclass.dll included in this 
package. It has to be registered from an elevated 32bit command prompt or
you can run the installer here:

  http://sandsprite.com/CodeStuff/subclassSetup.exe

it is open source you can also compile it yourself:

  https://github.com/dzzie/libs/tree/master/Subclass


start the vb_server first. it will launch which ever client sample you want.

of course you need java installed for the java sample and .net installed for the
c# example.


This is a simple example on how you can use the WM_COPYDATA window message
to communicate and share data between C, C#, JAVA, Delphi and VB apps.

In this example we use a VB Server. 

C clients are provided in both 32 and 64 bit executables for testing.

C Project files are for VS2008.

The Vb server subclasses its main window so it can detect when the 
window message comes in. When it finds a message meant for it, it will
extract the data and display it to the user.

For simplicity sake, I used my spSubclass.dll library to handle the
VB subclassing to keep the code lighter and cleaner in the app.

The VC app takes 1 command line argument which is the message you want
to be send back to VB server. The VC app will locate the HWND itself
using the Findwindow API. VB Mainwindow uses different classname if
running in IDE vs EXE so it checks for both.

Vista+ added a feature called UIPI it is part of UAC. It is only a problem
if the server is running at a higher permission level than the client.

I have included code to show you how to deal with this.

