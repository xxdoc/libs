#include <windows.h>
#include <stdio.h>
#include <stddef.h>
#include <conio.h>

#include "./../libdasm.h"

void main(void){

	INSTRUCTION i = {0};
	OPERAND o ;

	//Size of i = 0xd4, sizeof o = 0x38
	printf("Size of i = 0x%x, sizeof o = 0x%x\n", sizeof(i), sizeof(o));

	//OffsetOf i.mode = 8
	printf("OffsetOf i.mode = %x\n", offsetof(INSTRUCTION,mode) );

	printf("Press any key to exit...");
	getch();



}
