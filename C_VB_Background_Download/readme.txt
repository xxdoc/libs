
test of a multithreaded background downloader in C for vb.

c threads do direct memory writes to private vb class variables
vb class makes sure downloader is complete before it allows its class instance to terminate

wininet code may still need some tweaks for certs and ssl or follow redirects
or detecting if server denied it the link because of wrong tls version etc.

I was mostly focused on testing the multithreading access and feedback without
having to add subclassing to the main form.

