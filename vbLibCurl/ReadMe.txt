
modified sample for Jeffery Phillips vbLibCurl 
  https://sourceforge.net/projects/libcurl-vb/files/libcurl-vb/libcurl.vb%201.01/

this repo: https://github.com/dzzie/libs/tree/master/vbLibCurl

Additions:
-------------------------------
- initLib() to find/load C dll dependencies on the fly from different paths
- removed tlb requirements (all enums covered but not all api declares written yet)
- added higher level framework around low level api
- file progress, response object, abort
- download to memory only or file
- file downloads do not touch cache or temp

Notes:

Upgraded to the current libcurl v7.73 which was a drop in replacement 
for the 15yr old version originally used (v7.13). 

The newer versions will run on XP SP2 and newer because of the 
normaliz.Idn2Ascii import. The updated libcurl is required to talk to modern 
ssl servers. The old libcurl would run on win2k or newer. 

vb's built in Put file write command has a 2gb file size limit. I will leave it 
to the reader to switch over to API file writes if you need it.
------------------------------

https://curl.se/windows/
https://curl.se/windows/dl-7.73.0_1/openssl-1.1.1h_1-win32-mingw.zip

 libcurl 7.73.0_1 was built and statically linked with

    OpenSSL 1.1.1h [64bit/32bit]
    brotli 1.0.9 [64bit/32bit]
    libssh2 1.9.0 [64bit/32bit]
    nghttp2 1.41.0 [64bit/32bit]
    zlib 1.2.11 [64bit/32bit]
    zstd 1.4.5 [64bit/32bit] 

The following tools/compilers were used in the build process:

    binutils-mingw-w64-i686 2.35
    binutils-mingw-w64-x86_64 2.35
    clang 9.0.1
    gcc-mingw-w64-i686 10-win32
    gcc-mingw-w64-x86_64 10-win32
    mingw-w64 8.0.0-1 


we dont use curl.exe but i included it just in case it might be handy

