'**************************************************
'* NT Service sample                              *
'* http://www.smsoft.ru                           *
'* e-mail: sm@smsoft.ru                           *
'* The code is freeware. It may be used           *
'* in programs of any kind without permission     *
'**************************************************

Writing NT service using VB6
   
This sample program was built to show how to program Windows NT/2000/XP services using
Visual Basic. 

Of course, you can use free NTSVC.OCX control from Microsoft to create your service,
and this way is simple and reliable, but it has one disadvantage: you can't set
"Unattended execution" option for it. You can't use any OCX controls when "Unattended
execution" is set. 

This sample is written using VB6 without any external components. It takes into account
multithreading problems which seriousely limit VB functionality. Some used API functions
are declared in type library, and it is a workaround.

As the service was compiled for unattended execution, it has no visual interface.
Use NT Event Viewer to read messages written by the service to the Application Log. 

The functional part of service (which is absent in this sample) must be object-oriented
and event-driven. All events must be processed during few seconds, otherwise the service
will be unable to respond on requests from the Service Dispatcher.

See also:

Microsoft Knowledge Base Q137890, Q170883, Q175948,

http://msdn.microsoft.com/library/techart/msdn_ntsrvocx.htm,

http://msdn.microsoft.com/library/periodic/period98/service.htm,

http://msdn.microsoft.com/library/periodic/period98/vb98j1.htm

http://vbwire.com/advanced/howto/service.asp 

http://vbwire.com/advanced/howto/service2.asp

Matthew Curland's article "Create Worker DLL Threads" in the 06/99 issue of Visual Basic
Programmer's Journal.
 

30 September 2001

Update:
1. All API functions calls except GetVersionEx changed to Unicode versions, in code
and in type library. Added new Enum members to type library to support new Windows 2000
control codes.
2. Added MsgWaitObj function to prevent blocking of messages processing. All calls of
WaitForSingleObject and WaitForMultipleObjects in Sub Main replaced with MsgWaitObj.

06 June 2004