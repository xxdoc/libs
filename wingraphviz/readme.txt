
WinGraphviz is a single DLL COM object directly usable from vb6.
It uses the code from Graphviz which is public and freely available.

The "current" version of this DLL is based on the 1.8.x of Graphviz
which is from about 2005 I think. I have noticed some crashes, and some hangs
when trying to use complex examples from the current documentation. But it works
very well for the cases I have tried from my own code.

The Visual Basic six project file includes a basic graph and node generating class
that I wrote. These can generate the DOT textbased graph definition syntax to feed
to the library. The library is capable of very complex layouts. My classes only support
the basics that I see myself immediately needing. 