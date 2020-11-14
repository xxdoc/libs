
modified sample for Jeffery Phillips vbLibCurl 
  https://sourceforge.net/projects/libcurl-vb/files/libcurl-vb/libcurl.vb%201.01/

Upgraded to the current libcurl seems to be a drop in replacement 
for the 15yr old version originally used (v7.13)

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