
todo: switch over to PROCESS_QUERY_LIMITED_INFORMATION IF vista+

' x64 safe functions..
'    DumpProcess, DumpMemory, GetProcessModules, EnumDrivers, GetRunningProcesses, GetMemoryMap
'
' x64unsafe: ReadMemory

x64 support requires x64helper.exe in same directory as dll
EnumTasks, EnumMutexes requires EnumMutexes.dll

both of these external resources are contained in the proc_lib2 dll and will
be dropped to disk first time they are required.

source and binary can be found here: 

  https://github.com/dzzie/SysAnalyzer


CCmdOutput has been reworked and simplified. This dll will soon replace proc_lib