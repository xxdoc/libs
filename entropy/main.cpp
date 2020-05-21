#include <stdio.h>
#include <Windows.h>
#include <conio.h>
#include <math.h>

#define uint8_t unsigned char
#define uint32_t unsigned int
#define int64_t  __int64

static double log2(double n){
  return log(n) / log(2.0);
}

bool FileExists(LPCTSTR szPath)
{
  if(szPath==NULL) return false;
  DWORD dwAttrib = GetFileAttributes(szPath);
  bool rv = (dwAttrib != INVALID_FILE_ATTRIBUTES && !(dwAttrib & FILE_ATTRIBUTE_DIRECTORY)) ? true : false;
  return rv;
}

int file_length(FILE *f)
{
	int pos, end;
	pos = ftell (f);
	fseek (f, 0, SEEK_END);
	end = ftell (f);
	fseek (f, pos, SEEK_SET);
	return end;
}

double yara_entropy2(uint8_t* s, uint32_t bufLen)
{
  size_t i;
  double entropy = 0.0;
  
  if(s == NULL || bufLen == 0) return 0;

  double data[256]= {0.0};
 
  for (i = 0; i < bufLen ; i++)
  {
    uint8_t c = s[i];
    data[c] += 1;
  }

  for (i = 0; i < 256; i++)
  {
    if (data[i] != 0)
    {
      double x = (double) (data[i]) / bufLen;
      entropy -= x * log2(x);
    }
  }

  return entropy;
}

double yara_entropy(void* s){
	if(s == NULL) return 0;
	return yara_entropy2((uint8_t*)s, strlen((const char*)s));
}

double yara_data_entropy(char* fpath, int64_t offset = 0, int64_t length = -1)
{
	//todo: optimize to only load portion of file specified...
	FILE *fp;
	double entropy = 0;

	if (!FileExists(fpath)) return 0;

	fp = fopen(fpath, "rb");
	if(fp==0) return 0;

	uint32_t size = file_length(fp);
	if(offset > size){fclose(fp); return 0;}

	if(length == -1) length = size - offset;
	if((offset+length) > size){fclose(fp); return 0;}

	uint8_t *buf = (uint8_t*)malloc(size); 
	memset(buf, 0, size);
	fread(buf, 1, size, fp);
	fclose(fp);

	entropy = yara_entropy2(&buf[offset], length); 
	free(buf);
	return entropy;
}

double die_entropy2(uint8_t* buf, uint32_t length)
{
	double dEntropy = 1.4426950408889634073599246810023;
    double bytes[256] = {0.0};
    double temp;

	for(int i=0; i<length; i++){
		uint8_t c = buf[i];
		bytes[c] += 1;
	}

    for(int j=0; j<256; j++){
        temp=bytes[j]/(double)length;
        if(temp) dEntropy+=(-log(temp)/log((double)2))*bytes[j];
    }

    dEntropy=dEntropy/(double)length;
    return dEntropy;

}

double die_entropy(void* s){
	if(s == NULL) return 0;
	return die_entropy2((uint8_t*)s, strlen((const char*)s));
}

double die_data_entropy(char* fpath, unsigned int offset = 0,unsigned int length = -1)
{
 
	FILE *fp; //todo: optimize to only load portion of file specified...
	if (!FileExists(fpath)) return 0;

	fp = fopen(fpath, "rb");
	if(fp==0) return 0;

	uint32_t size = file_length(fp);
	if(offset > size){fclose(fp); return 0;}

	if(length == -1) length = size - offset;
	if((offset+length) > size){fclose(fp); return 0;}

	uint8_t *buf = (uint8_t*)malloc(size); 
	memset(buf, 0, size);
	fread(buf, 1, size, fp);
	fclose(fp);

	double entropy = die_entropy2( &buf[offset], length);
	free(buf);
	return entropy;
}


void main(void){
	
	printf("yara_entropy(AAAAA) = %.4f\n", yara_entropy("AAAAA"));
	printf("yara_entropy(abcde) = %.4f\n", yara_entropy("abcde"));

	printf("die_entropy(AAAAA) = %.4f\n", die_entropy("AAAAA"));
	printf("die_entropy(abcde) = %.4f\n", die_entropy("abcde"));
	printf("die_entropy(abcdefghijklmnop) = %.4f\n", die_entropy("abcdefghijklmnop"));

	printf("yara_data_entropy(random.dat) = %.4f\n", yara_data_entropy("random.dat"));
	printf("yara_data_entropy(random.dat) = %.4f\n", yara_data_entropy("random.dat",0,0x20));

	printf("yara_data_entropy(a.dat) = %.4f\n", yara_data_entropy("a.dat"));
	printf("yara_data_entropy(a.dat) = %.4f\n", yara_data_entropy("a.dat",0,0x20));

	printf("die_data_entropy(random.dat) = %.4f\n", die_data_entropy("random.dat"));
	printf("die_data_entropy(random.dat) = %.4f\n", die_data_entropy("random.dat",0,0x20));

	printf("die_data_entropy(a.dat) = %.4f\n", die_data_entropy("a.dat"));
	printf("die_data_entropy(a.dat) = %.4f\n", die_data_entropy("a.dat",0,0x20));



	printf("Press any key to exit...");
	getch();

}