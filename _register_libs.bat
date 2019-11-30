@echo off
goto check_Permissions

:do_install
	regsvr32 ./dzrt/dzrt.dll /s
	regsvr32 ./pe_lib/sppe.dll /s
	regsvr32 ./pe_lib2/sppe2.dll /s
	regsvr32 ./pe_lib3/sppe3.dll /s
	regsvr32 ./proc_lib/proclib.dll /s
	regsvr32 ./proc_lib2/proclib2.dll /s
	regsvr32 ./subclass/spSubclass.dll /s
	regsvr32 ./subclass2/spSubclass2.dll /s
	regsvr32 ./vb6_utypes/vbUtypes.dll /s
	regsvr32 ./vbDevKit/vbDevKit.dll /s
    goto fini
    
:check_Permissions
    echo Administrative permissions required. Detecting permissions...

    net session >nul 2>&1
    if %errorLevel% == 0 (
        echo Administrative permissions confirmed.
		goto do_install
    ) else (
        echo Failure: Current permissions inadequate.
    )

    pause >nul
	
:fini
    pause