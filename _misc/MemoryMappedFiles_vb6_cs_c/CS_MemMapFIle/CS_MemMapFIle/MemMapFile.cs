using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;

namespace ConsoleApplication1
{
    class MemMapFile  
    {
        
        [DllImport("kernel32")]
        public static extern int CreateFileMapping( int hFile, int lpAttributes, int flProtect,int dwMaximumSizeLow, int dwMaximumSizeHigh,String lpName);

        [DllImport("kernel32")]
        public static extern bool FlushViewOfFile( int lpBaseAddress, int dwNumBytesToFlush);

        [DllImport("kernel32")]
        public static extern uint MapViewOfFile(int hFileMappingObject, int dwDesiredAccess, int dwFileOffsetHigh, int dwFileOffsetLow, int dwNumBytesToMap);

        [DllImport("kernel32")]
        public static extern int OpenFileMapping(int dwDesiredAccess, bool bInheritHandle, String lpName);

        [DllImport("kernel32")]
        public static extern bool UnmapViewOfFile(int lpBaseAddress);

        [DllImport("kernel32")]
        public static extern bool CloseHandle(int handle);

        private const int PAGE_READWRITE = 0x4;
        private const int SECTION_MAP_WRITE = 0x2;
        private const int FILE_MAP_WRITE = SECTION_MAP_WRITE;
        private const int STANDARD_RIGHTS_REQUIRED = 0xF0000;
        private const int SECTION_QUERY = 0x1;
        private const int SECTION_MAP_READ = 0x4;
        private const int SECTION_MAP_EXECUTE = 0x8;
        private const int SECTION_EXTEND_SIZE = 0x10;

        private const int SECTION_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED | SECTION_QUERY | 
                                               SECTION_MAP_WRITE |  SECTION_MAP_READ |
                                               SECTION_MAP_EXECUTE | SECTION_EXTEND_SIZE;

        private const int FILE_MAP_ALL_ACCESS = 0xF001F;

        private string VFileName;
        private int MaxSize;
        private int hFile;
        private uint gAddr;

        public string ErrorMessage;
        public bool DebugMode  = false;

        public bool CreateMemMapFile(string fName, int mSize){

            byte[] b = new byte[mSize];
            VFileName = fName.ToUpper(); 
            MaxSize = mSize;
            
            if(hFile!=0){
                ErrorMessage = "Cannot open multiple virtural files with one class";
                return false;
            }
            
            hFile = CreateFileMapping(-1, 0, PAGE_READWRITE, 0, mSize, VFileName);

            if( hFile == 0 ){
                ErrorMessage = "Unable to create virtual file";
                return false;
            }
    
             gAddr = MapViewOfFile(hFile, FILE_MAP_ALL_ACCESS, 0, 0, mSize);

             if (gAddr == 0) return false;
             return true;
            
        }

        public bool WriteFile(byte[] bData){

            if(bData.Length == 0) return false;

            if(bData.Length > this.MaxSize){
                ErrorMessage = "Data is to large for buffer!";
                return false;
            }
    
            if(hFile==0){
                ErrorMessage = "Virtual File or Virtual File Interface not initialized";
                return false;
            }

            Marshal.Copy(bData, 0, (IntPtr)(this.gAddr), bData.Length);
            return true;
        }

        public bool ReadFile(byte[] bData, int length){

            if(length > this.MaxSize){
                ErrorMessage = "ReadLength to large for buffer!";
                return false;
            }
    
            if(hFile==0){
                ErrorMessage = "Virtual File or Virtual File Interface not initialized";
                return false;
            }

            Marshal.Copy((IntPtr)gAddr, bData,(int)0, length);
            return true;
        }


    }
}
