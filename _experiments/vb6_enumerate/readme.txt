
trying to use a user function as a built in language feature
to replicate the python 

for i,s in enumerate(array):

type construct.

enumerate() works with:
- arrays
- collections (can optionally return item key as well)
- strings: if optional key is included string is split at key, else walks each letter
- textbox: if optional key is included text split at key, else split at vbcrlf
- listbox
- combobox


and can auto set the index, value and collection key
as the ary/col is walked.

it will detect if your startIndex is not initilized and you
change objects. 

you can not nest calls to enumerate or you will get an endless loop

in trying to make this able to be used with no setup other than the call
to mimic a native language feature we do sacrifice that..

this replaces the case where you want to walk an array or collection
but still require the item index so can not use for each

as a bonus this method also allows you to use any variable type you want
as the enumerator where as for each requires a variant type. This is useful
for arrays/collections of class objects so you have intellisense support for
the object. 

for collections, the enumerator can also give you the collection items key
which can be really handy to have as well..

I dunno..this isnt an everyday construct you would need, but it will be useful
at times. 

I think I am going to add this to my dzrt library in the globals class so it is
usable in any projects without having to declare it..it will be quite close to
a built in language feature at that point..

note: we dont support listviews because 

dim li as listitem
for each li in lv.listitems

can not be improved upon..i dont think I ever need the item index and it
might even be available in the listitem properties..