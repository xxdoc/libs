
#include <windows.h>
#include "msvbvm60.tlh"
#include <conio.h>

//#import "C:\\windows\system32\msvbvm60.dll" no_namespace

#import "Project1.dll" 

Project1::_Class1Ptr c;

void addStr(_CollectionPtr p , char* str){
	_variant_t vv;
	vv.SetString(str);
	p->Add( &vv.GetVARIANT() );
}

void addInt(_CollectionPtr p , int x){
	_variant_t vv;// = x; //this gave error automation type not supported in visual basic..
						  //see comutil.h it was making it an VT_INT, a long -> VT_I4
	vv.vt = VT_I4;        //I will just set it manually or can change fun prototype or cast value..
	vv.intVal  = x;
	p->Add( &vv.GetVARIANT() );
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

	addInt(pColl,47);
	addStr(pColl, "this is my wrapper!");
	addStr(pColl, "this is my wrapper1");
	addStr(pColl, "this is my wrapper2");
	addStr(pColl, "this is my wrapper3");
	
	//remember vb collections start at 1
	for(long i=1; i <= pColl->Count(); i++){
		_variant_t index = i;
		_variant_t vTmp = pColl->Item(&index);

		//automatically converts (some) types to strings for you
		//I think I read this can throw an exception in some cases..(VT_UNKNOWN?)
		_bstr_t b = vTmp; 

		printf("Item %d is type.vt=%d  CStr('%s')\n",i, vTmp.vt, (char*)b);
	}

	c->test();

	 

	/* vb6 collection is not creatable externally via clsid apparently..
	HRESULT hr = pColl.CreateInstance(clsid);
	if(FAILED(hr) )
	{
		MessageBox(0,"_Collection:CreateInstance:Failed","Oops...",0);
	}*/

    printf("\nPress any key to exit...");
    getch();
}