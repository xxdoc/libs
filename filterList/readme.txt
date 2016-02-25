
small usercontrol that gives you a listview control
with a built in filter textbox on the bottom.

very similar to use as original, couple bonus functions thrown
in on top. 

simple but very useful.

set integer FilterColumn to determine which column it searchs 
0 is .text rest is .subitem index

obscure feature: the filter text box can also double as a command
line processor.

currently supports

Const cmdHelp = "Supports following commands: \n" & _
                    "/fc[number]     set filter column number \n" & _
                    "/copy           copy entire listview contents \n" & _
                    "/copysel        copy selected items in listview \n" & _
                    "/cc[number]     copy all elements from column number \n" & _
                    "/multi          toggle multi selection mode \n" & _
                    "/hide           toggle hide selection mode \n" & _
                    "/help           display this help message"