
author:  david zimmer [dzzie@yahoo.com]
site:    http://sandsprite.com
license: free for any use

small usercontrol that gives you a listview control
with a built in filter textbox on the bottom.

very similar to use as original, couple bonus functions thrown
in on top. 

simple but very useful.

set integer FilterColumn to determine which column it searchs 
1 based. You can set this any time, including before setting
column headers. You can also specify it in the call to 
SetColumnHeaders by including an * in the column you want.

The user can also change FilterColumn on the fly from the popup
menu, or through entering /[index] in filter textbox and hitting
return.

you can apply multiple filters by seperating values with commas.
You can use subtractive filters by entering a - as the first character
in the filter textbox, this also supports a csv list.

The filter popup has a help message with more details.

When the control is locked no events will be generated or
processed. filter textbox locked and grayed.

If allowDelete property is set, user can hit delete key to
remove items from list box. This supports removing items
from the filtered results as well. (Even if the user resorted
the columns with the built in column click sort handler)

When you resize the control, the last listview item column 
header will grow. You specify initial column widths in set header
call. When running in the IDE, there is a debug menu item available
that will give you the current column width settings to copy out for
the set column header call. So just manually adjust them, then use
the menu item, then you can easily set them as startup defaults.

the current list count is always available on the popup menu along
with some basic macros to allow the user to copy the whole table,
copy a specific column, copy selected entries etc. 


examples:

lvFilter.SetColumnHeaders "test1,test2,test3*,test4", "870,1440,1440"

Set li = lvFilter.AddItem("text" & i)
li.subItems(1) = "taco1 " & i

Set li = lvFilter.AddItem("text", "item1", "item2", "item3")
lvFilter.SetLiColor li, vbBlue