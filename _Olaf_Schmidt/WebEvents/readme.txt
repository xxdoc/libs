
Author:    Olaf Schmidt
Link:      http://www.vbforums.com/showthread.php?847773-VB6-elevated-IE-Control-usage-with-HTML5-elements-and-COM-Event-connectors
Download:  http://vbrichclient.com/Downloads/WebEvents.zip

Just a small Demo, how one can work these days with the IE-Control on a VB6-Form.

There's three main-topics which this Demo addresses:
- how to elevate the IE-Version from its default (which for compatibility-reasons always mimicks the age-old version 7)
- how to load the IE-Control dynamically, after the elevation above went through
- how to connect Elements on a page comfortably to normal VB6-EventHandlers

But also addressed is stuff like:
- how to load ones own HTML-template-code from a string into the Control
- how to enable the "themed look" of the Browser-Controls (avoiding the old "sunken edge 3D-style")
- how to work with the HTML5-canvas (in a "cairo-like-fashion") to produce antialiased output

The Event-approach as shown in the Demo does not require any References
or COMponent-check-ins, or Typelibs - the whole thing is based on a plain, virginal VB6-Project
which does not have any dependencies (and thus should work without installation anywhere when compiled).