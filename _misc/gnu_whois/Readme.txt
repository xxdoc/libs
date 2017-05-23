
http://sourceforge.net/projects/whoiswin/reviews/?sort=created_date&stars=0#reviews-n-ratings

 This project is basically an imported version of GNU-whois command-line tool. 
Since the original GNU-whois is an "inteligent" one, which choose the right (or 
likely right) whois server before sending the request messages, it is even 
better than the most GUI tools of Windows versions emerged on the Internet. At 
least, I feel convenient that I can use this tool even on a Windows box just as 
same as on a Linux one.

 The main part is unchanged, except the socket APIs and some posix functions. All 
the "adaptor" functions are in the win_funcs.c file. Some codes of error 
checking are removed for there is some difficulty for me to rewrite the Windows-
version ones, since these errors are unlikely to occure.

 The license of this project is GPL.


--To built
1. Run the batch make_headers.bat to generate the header files.

2. If you are under the mingw32, run the batch make_mingw32.bat.

3. If you are under an MS Visual Studio command line window, run the batch 
   make_vs.bat.
  
4. In both cases the output will be put under the dir 'output'.

5. Have fun.

By Yuyunwu 2007.12

yuyunwu1@gmail.com
