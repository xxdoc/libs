
#include <windows.h>
#include "msvbvm60.tlh"
#include <conio.h>

//#import "C:\\windows\system32\msvbvm60.dll" no_namespace

#import "c:\Project1.dll" 

Project1::_Class1Ptr c;

void addStr(_CollectionPtr p , char* str){
	_variant_t vv;
	vv.SetString(str);
	p->Add( &vv.GetVARIANT() );

	/*VARIANT v;
	VariantInit(&v);
	v.bstrVal = SysAllocString(L"this is string 1!");
	v.vt = VT_BSTR;*/
}

void main(void)
{
  //{A4C4671C-499F-101B-BB78-00AA00383CBB}
	 IUnknown *u=0;
	 _CollectionPtr pColl;
	 IID clsid = {0xA4C46780,0x499F,0x101B,0xBB,0x78,0x00,0xAA,0x00,0x38,0x3C,0xBB};

	CoInitialize(NULL);
	
	HRESULT hr = c.CreateInstance( __uuidof( Project1::Class1 ) );

	pColl = c->getCol();

	VARIANT v;
	VariantInit(&v);
	v.bstrVal = SysAllocString(L"this is string 1!");
	v.vt = VT_BSTR;

	pColl->Add(&v);

	VARIANT v2;
	VariantInit(&v2);
	v2.bstrVal = SysAllocString(L"this is string 2!");
	v2.vt = VT_BSTR;

	pColl->Add(&v2);

	addStr(pColl, "this is my wrapper!");
	addStr(pColl, "this is my wrapper1");
	addStr(pColl, "this is my wrapper2");
	addStr(pColl, "this is my wrapper3");
	 
	c->test();

	 

	/*HRESULT hr = pColl.CreateInstance(clsid);
	if(FAILED(hr) )
	{
	MessageBox(0,"_Collection:CreateInstance:Failed","CSvcEventWrapper",0);
	}*/


     getch();



 
 

}