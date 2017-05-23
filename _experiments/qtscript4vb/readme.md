<pre>

blog post: 
	<a href="http://sandsprite.com/blogs/index.php?uid=11&pid=351">http://sandsprite.com/blogs/index.php?uid=11&pid=351</a>
	
video walkthrough:
	<a href="https://www.youtube.com/watch?v=FM81NBBJu6Q">https://www.youtube.com/watch?v=FM81NBBJu6Q</a>
	

NOTE:
  i have included copies of the debug builds of QT4.8 dlls required to run this test
  as a convience. (26mb but saves you a 234mb dl + install)
  
  otherwise to run this we assume you have QT4.8 installed
  AND that its \bin directory has been added to your PATH envirnoment variable
  
  

the following are for VS2008 development

 <a href="https://download.qt.io/archive/qt/4.8/4.8.5/">https://download.qt.io/archive/qt/4.8/4.8.5/</a>
   qt-win-opensource-4.8.5-vs2008.exe = 234 MB

to compile the dll with VS2008 use the following:
 <a href="https://download.qt.io/official_releases/vsaddin/">https://download.qt.io/official_releases/vsaddin/</a>
   qt-vs-addin-1.1.11-opensource.exe = 112 MB

the runtime distribution requirements are:

release mode: 12.2mb
  QtCore4.dll
  QtGui4.dll
  QtScript4.dll
  QtScriptTools4.dll

debug mode dll: 26mb
  QtScriptToolsd4.dll
  QtCored4.dll
  QtGuid4.dll
  QtScriptd4.dll

if we use a scriptAgent instead of the full fledged script debugger ui the release 
mode dependancy is down to just QtScript4,QtScriptTools4 1.8mb, i will experiment 
with that in another project..
</pre>

![screenshot](https://raw.githubusercontent.com/dzzie/QtScript4vb/master/qtscript_debug_ui.png)