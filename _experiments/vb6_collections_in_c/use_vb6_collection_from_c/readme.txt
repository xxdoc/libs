
in this first test we use a C++ app to create a vb6 collection
and then add some items to it. 

finally we have the vb6 app cycle through the collection and let
us know if it worked.

note you can not create a vb6 collection on your own from c++
apparently you would have to create a clone of a collection object
in all native c++ as the vb6 one is not creatable.

to get around this for my test i just made a small vb6 dll that
i load from c++ and it creates the collection for me and hands
me a reference to it.

my normal use will be from a vb6 app calling a c++ dll, so I 
can pass in a reference to the collection to the c code anyway.