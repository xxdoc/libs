#include <stdio.h>
#include <Windows.h>
#include <conio.h>
#include <math.h>

/*

	The entropy code was adapted from the open source Detect It Easy signature scanner engine

	https://github.com/horsicq/Detect-It-Easy

*/

#define MINIMAL(a,b)            (((a) < (b)) ? (a) : (b))

bool readChunkFromFile(FILE* f, unsigned int nOffset, char *pBuffer, unsigned int nSize)
{
	if(ftell(f) != nOffset) fseek(f,0,nOffset);
	unsigned int sz = fread(pBuffer,1,nSize,f);
	if(sz != nSize) return false;
	return true;
}

int file_length(FILE *f)
{
	int pos;
	int end;

	pos = ftell(f);
	fseek (f, 0, SEEK_END);
	end = ftell(f);
	fseek (f, pos, SEEK_SET);

	return end;
}

float calcFileEntropy(FILE* f, unsigned int nOffset, unsigned int nDataSize)
{
	#define BUFFER_SIZE 0x1000
	unsigned int fSize = file_length(f);

    if(nDataSize == 0) return 0.0;
	
    if(nDataSize == -1)/* keyword signifies offset to eof */
    {
		nDataSize = fSize - nOffset;

        if(nDataSize==0)
        {
            return 0.0;
        }
    }

    if(nOffset >= fSize) return 0.0;
    if(nOffset + nDataSize > fSize) return 0.0;

    unsigned int nSize = nDataSize;

    float fEntropy=1.4426950408889634073599246810023;
    float bytes[256]= {0.0};
    float temp;

    unsigned int nTemp=0;
    char *pBuffer=new char[BUFFER_SIZE];

    while(nSize>0)
    {
        nTemp=MINIMAL(BUFFER_SIZE,nSize);

        if(!readChunkFromFile(f,nOffset,pBuffer,nTemp))
        {
            delete[] pBuffer;
            printf("Read error");
            return 0;
        }

        for(int i=0; i<nTemp; i++)
        {
            bytes[(unsigned char)pBuffer[i]]+=1;
        }

        nSize-=nTemp;
        nOffset+=nTemp;
    }

    delete[] pBuffer;

    for(int j=0; j<256; j++)
    {
        temp=bytes[j]/(float)nDataSize;

        if(temp)
        {
            fEntropy+=(-log(temp)/log((float)2))*bytes[j];
        }
    }

    fEntropy=fEntropy/(float)nDataSize;

    return fEntropy;
}

float calcMemEntropy(char* buf, unsigned int bufSize, unsigned int nOffset, unsigned int nDataSize)
{
	#define BUFFER_SIZE 0x1000
	unsigned int fSize = bufSize;

    if(nDataSize == 0) return 0.0;
	
    if(nDataSize == -1)/* keyword signifies offset to end of buf */
    {
		nDataSize = fSize - nOffset;
        if(nDataSize==0) return 0.0;
    }

    if(nOffset >= fSize) return 0.0;
    if(nOffset + nDataSize > fSize) return 0.0;

    unsigned int nSize = nDataSize;

    float fEntropy=1.4426950408889634073599246810023;
    float bytes[256]= {0.0};
    float temp;

    unsigned int nTemp=0;
    char *pBuffer = buf + nOffset;

    while(nSize>0)
    {
        nTemp=MINIMAL(BUFFER_SIZE,nSize);
		
        for(int i=0; i<nTemp; i++)
        {
            bytes[(unsigned char)pBuffer[i]]+=1;
        }

        nSize-=nTemp;
        nOffset+=nTemp;
		pBuffer+=nTemp;
    }

    for(int j=0; j<256; j++)
    {
        temp=bytes[j]/(float)nDataSize;

        if(temp)
        {
            fEntropy+=(-log(temp)/log((float)2))*bytes[j];
        }
    }

    fEntropy=fEntropy/(float)nDataSize;

    return fEntropy;
}


void main(void){

	char* fName = "C:\\windows\\notepad.exe";
	FILE* f = fopen(fName, "rb");

	if(f==NULL){
		printf("File not found");
		return;
	}
	
	float entropy = calcFileEntropy(f,0,-1);
	fclose(f);

	printf("Entropy of %s = %f\n", fName, entropy);
	
	char *pBuffer = new char[500];
	memset(pBuffer,0,500);
	for(int i=0;i < 255; i++){
		pBuffer[i] = i;
	}
	entropy = calcMemEntropy(pBuffer, 500, 0, 255);
    printf("Entropy of test mem alloc section 1= %f\n", entropy);

	entropy = calcMemEntropy(pBuffer, 500, 256, -1);
    printf("Entropy of test mem alloc section 2 = %f\n", entropy);

	delete[] pBuffer;

	#define TESTSZ  0x2550
	pBuffer = new char[TESTSZ]; //(char*)malloc(TESTSZ); //either you will see a difference in release mode since debug mode alloc sets mem to pattern..
	entropy = calcMemEntropy(pBuffer, TESTSZ, 0, -1);
	printf("Entropy of Unitilized mem alloc = %f\n", entropy);

	memset(pBuffer,0,TESTSZ);
	entropy = calcMemEntropy(pBuffer, TESTSZ, 0, -1);
	printf("Entropy of memset mem alloc = %f\n", entropy);
	
	printf("Press any key to exit...");
	getch();



}