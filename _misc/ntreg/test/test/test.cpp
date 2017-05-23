#include <windows.h>
#include "stdstring.h"
#include "NtRegistry.h"

//http://www.codeproject.com/Articles/14508/Registry-Manipulation-Using-NT-Native-APIs

void Output(CStdString csMsg, DWORD Buttons /*MB_OK*/)
{
	MessageBox( NULL, csMsg, _T("CNtRegistry: Message..."), Buttons );
}

CStdString GetRootPath(void){ return m_csRootPath; }




//////////////////////////////////////////////////////////////////////
//
// Pretty self explanitory
//				
//////////////////////////////////////////////////////////////////////
CStdString GetRootPathFor(HKEY hRoot)
{
	// Defaults to HKCU :-)
	CStdString csRootPath = _T("");
	if (hRoot == HKEY_LOCAL_MACHINE) {
		csRootPath = _T("\\Registry\\Machine");
	}
	else if (hRoot == HKEY_CLASSES_ROOT) {
		csRootPath = _T("\\Registry\\Machine\\SOFTWARE\\Classes");
	}
	else if (hRoot == HKEY_CURRENT_CONFIG) {
		csRootPath = _T("\\Registry\\Machine\\System\\CurrentControlSet\\Hardware Profiles\\Current");
	}
	else if (hRoot == HKEY_USERS) {
		csRootPath = _T("\\Registry\\User");
	}
	else {
		csRootPath.Format(_T("\\Registry\\User\\%s"), m_csSID);
	}
	return csRootPath;
}


//////////////////////////////////////////////////////////////////////
//
// Make sure the string includes the "\registry\ ... "
//
//////////////////////////////////////////////////////////////////////
CStdString CheckRegFullPath(CStdString csPath)
{
	CStdString csFullPath = _T("");
	CStdString csTempPath = csPath;

	csTempPath.MakeLower();

	// This is so we know where we stand.
	if (csTempPath.Left(10) != _T("\\registry\\")) {
		//
		csFullPath = GetRootPath();

		// Ok, let's build the full key ...
		if (csTempPath[0] != '\\') {
			csFullPath += _T("\\");
		}

		// Append it...
		csFullPath += csPath;

		// Check if it ends in a slash...if so, take it out!!
		if (csFullPath.Right(1) == _T("\\")) {
			CStdString csTest(csFullPath);
			csFullPath = csTest.Left(csTest.GetLength() - 1);
		}
	}
	else {
		csFullPath = csPath;
	}

	// put it into a CStdString Array...
	//m_tokenEx.Split(csFullPath,_T("\\"));

	return csFullPath;
}

/****************************************************************************
**
**	Function:	IsKeyHidden
**
**  Purpose:	Call this function to check if it is a hidden subkey. 
**
**  Arguments:	(IN) - FULL path to the key to be checked.
**
**  Returns:	BOOL - Success/Failure.
**				
****************************************************************************/
BOOL IsKeyHidden(CStdString csKey)
{
	// Make sure the "FullKey" is given...
	CStdString csFullKey = CheckRegFullPath(csKey);

	ANSI_STRING asKey;
	RtlZeroMemory(&asKey,sizeof(asKey));
	RtlInitAnsiString(&asKey,csFullKey);

	UNICODE_STRING usKeyName;
	RtlZeroMemory(&usKeyName,sizeof(usKeyName));

	RtlAnsiStringToUnicodeString(&usKeyName,&asKey,TRUE);
	usKeyName.Length += 2;
	usKeyName.MaximumLength += 2;

	//
	OBJECT_ATTRIBUTES ObjectAttributes;
	InitializeObjectAttributes(&ObjectAttributes,&usKeyName,OBJ_CASE_INSENSITIVE,m_hMachineReg,NULL);

	//
	HANDLE hKey = NULL;
	m_NtStatus = NtOpenKey(&hKey, KEY_ALL_ACCESS, &ObjectAttributes);
	if(!NT_SUCCESS(m_NtStatus)) {
		return FALSE;
	}
	NtClose(hKey);
	return TRUE;
}

//////////////////////////////////////////////////////////////////////
//
//	Function:	DisplayError
//
//  Purpose:	Display the Error according to the error # passed...I don't 
//				understand just numbers ;)
//
//  Arguments:	DWORD dwLastError - Obvious...
//
//  Returns:	CStdString - Error message.
//
//////////////////////////////////////////////////////////////////////
CStdString DisplayError(DWORD dwError)
{
	LPVOID lpMessageBuffer = NULL;

	// Load the NTDLL.dll specifically
	HMODULE hNTDLL = LoadLibrary(_T("NTDLL.DLL"));

	FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM
				 |FORMAT_MESSAGE_FROM_HMODULE,hNTDLL,
				 dwError,MAKELANGID(LANG_NEUTRAL,SUBLANG_DEFAULT), // Default language
				 (LPSTR)&lpMessageBuffer,0,NULL);

	CStdString csReturn = _T("");
	CStdString csErrorMessage = (LPCTSTR)lpMessageBuffer;

	if( lpMessageBuffer != NULL ) {
		LocalFree(lpMessageBuffer);
		csReturn.Format("Error( %u [0x%8XL] ) - %s", dwError, dwError, csErrorMessage.Left(csErrorMessage.GetLength()-2));
	}
	else {
		csReturn.Format("Error( %u [0x%8XL] ) - No NT Error Data was returned??", dwError, dwError);
	}

	FreeLibrary(hNTDLL);

	return csReturn;

}


//////////////////////////////////////////////////////////////////////
//
//	Function:	GetTextualSid
//
//  Purpose:	Get the textual Security Identifier (SID) for the 
//				current user.
//
//				NOTE:  This function us called from the LookupSID function
//
//  Arguments:	(IN)  PSID		- Binary SID.
//				(OUT) LPSTR		- Buffer for Textual SID.
//				(IN)  LPDWORD	- Required/provided buffer size.
//
//  Returns:	BOOL - Success/Failure.
//
//////////////////////////////////////////////////////////////////////
BOOL GetTextualSid(PSID pSid, LPSTR szTextualSid, LPDWORD dwBufferLen) 
{
	PSID_IDENTIFIER_AUTHORITY psia;
	DWORD dwSubAuthorities;
	DWORD dwSidRev = SID_REVISION;
	DWORD dwCounter;
	DWORD dwSidSize;

	BOOL bReturn = FALSE;
    try {{
		//
		// Test if SID passed in is valid.
		if(!IsValidSid(pSid)) {
            goto leave;
		}

		// Obtain SidIdentifierAuthority.
		psia = GetSidIdentifierAuthority(pSid);

		// Obtain sidsubauthority count.
		dwSubAuthorities = *GetSidSubAuthorityCount(pSid);

		// Compute buffer length.
		// S-SID_REVISION- + identifierauthority- + subauthorities- + NULL
		dwSidSize = (15 + 12 + (12 * dwSubAuthorities) + 1) * sizeof(TCHAR);

		// Check provided buffer length.
		// If not large enough, indicate proper size and setlasterror
		if (*dwBufferLen < dwSidSize) {
			*dwBufferLen = dwSidSize;
			SetLastError(ERROR_INSUFFICIENT_BUFFER);
			m_NtStatus = ERROR_INSUFFICIENT_BUFFER;
            goto leave;
		}

		// Prepare S-SID_REVISION-.
		dwSidSize = wsprintf(szTextualSid, TEXT("S-%lu-"), dwSidRev);

		// Prepare SidIdentifierAuthority.
		if ((psia->Value[0] != 0) || (psia->Value[1] != 0)) {
			dwSidSize += wsprintf(szTextualSid + lstrlen(szTextualSid),
									TEXT("0x%02hx%02hx%02hx%02hx%02hx%02hx"),
									(USHORT) psia->Value[0],
									(USHORT) psia->Value[1],
									(USHORT) psia->Value[2],
									(USHORT) psia->Value[3],
									(USHORT) psia->Value[4],
									(USHORT) psia->Value[5]);
	   
		} 
		else {
			dwSidSize += wsprintf(szTextualSid + lstrlen(szTextualSid),
									TEXT("%lu"),
									(ULONG) (psia->Value[5]      ) +
									(ULONG) (psia->Value[4] <<  8) +
									(ULONG) (psia->Value[3] << 16) +
									(ULONG) (psia->Value[2] << 24));
		}

		// Loop through SidSubAuthorities.
		for (dwCounter = 0; dwCounter < dwSubAuthorities; dwCounter++) {
			dwSidSize += wsprintf(szTextualSid + dwSidSize, TEXT("-%lu"), *GetSidSubAuthority(pSid, dwCounter));
		}
		bReturn = TRUE;
		//
    } leave:;
    } catch(...) {}

	return bReturn;
}


//////////////////////////////////////////////////////////////////////
//
//	Function:	LookupSID
//
//  Purpose:	Lookup the textual Security Identifier (SID) for the 
//				current user.
//
//				NOTE:  This calls the GetTextualSid function above.
//
//  Arguments:	(IN/OUT)  CStdString - Location to place the SID.
//
//  Returns:	BOOL - Success/Failure.
//
//////////////////////////////////////////////////////////////////////
BOOL LookupSID (CStdString &csSID)
{
	HANDLE		hToken			= NULL;
	PTOKEN_USER	ptgUser			= NULL;
	DWORD		cbBuffer		= 0;
	LPSTR		szTextualSid	= NULL;
	DWORD		cbSid			= 36;
	TOKEN_INFORMATION_CLASS tic = TokenUser;

	CStdString csError = _T("Error");
	csSID = csError;

	// Obtain current process token.
	m_NtStatus = NtOpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &hToken);
	if(!NT_SUCCESS(m_NtStatus)) {
		Output(DisplayError(m_NtStatus), MB_OK|MB_ICONERROR);
		return FALSE;
	}

	// Obtain user identified by current process's access token.
	// Basically, query info in the token ;-)
	m_NtStatus = NtQueryInformationToken(hToken, 
										 tic,
										 ptgUser, 
										 cbBuffer, 
										 &cbBuffer);
	if(!NT_SUCCESS(m_NtStatus)) {
		//
		ptgUser = (PTOKEN_USER)RtlAllocateHeap(GetProcessHeap(), HEAP_ZERO_MEMORY, cbBuffer);
		if (!ptgUser) {
			Output(DisplayError(m_NtStatus), MB_OK|MB_ICONERROR);
			goto SIDEnd;
		}

		m_NtStatus = NtQueryInformationToken(hToken, 
											 tic,
											 ptgUser, 
											 cbBuffer, 
											 &cbBuffer); 
		if(!NT_SUCCESS(m_NtStatus)) {
			Output(DisplayError(m_NtStatus), MB_OK|MB_ICONERROR);
			goto SIDEnd;
		}
	}

	cbSid = 128;
	szTextualSid = (LPSTR) RtlAllocateHeap(GetProcessHeap(), HEAP_ZERO_MEMORY, cbSid);
	if (!szTextualSid) {
		Output(DisplayError(m_NtStatus), MB_OK|MB_ICONERROR);
		goto SIDEnd;
	}

	// Obtain the textual representation of the SID.
	if (!GetTextualSid( ptgUser->User.Sid,	// user binary Sid
						szTextualSid,		// buffer for TextualSid
						&cbSid))			// size/required buffer
	{
		m_NtStatus = GetLastError();
		Output(DisplayError(m_NtStatus), MB_OK|MB_ICONERROR);
		goto SIDEnd;
	}

	// the TextualSid representation.
	csSID.Format(_T("%s"), szTextualSid);
   

SIDEnd:

	// Free resources.
	if (hToken) {
		NtClose(hToken);
	}

	if (ptgUser) {
		RtlFreeHeap(GetProcessHeap(), 0, ptgUser);
	}

	if (szTextualSid) {
		RtlFreeHeap(GetProcessHeap(), 0, szTextualSid);
	}

	return TRUE;
}


//////////////////////////////////////////////////////////////////////
//
//	Function:	LocateNTDLLEntryPoints
//
//  Purpose:	Loads and finds the entry points we need in NTDLL.DLL 
//
//  Arguments:	IN CStdString csErr - Reason it failed (if returned FALSE).
//
//  Returns:	BOOL - Success/Failure.
//
//////////////////////////////////////////////////////////////////////
BOOL LocateNTDLLEntryPoints(CStdString& csErr)
{
	csErr = _T("");

	HINSTANCE hinstStub = GetModuleHandle(_T("ntdll.dll"));
	if(hinstStub) {
		//
		NtOpenThread = (LPNTOPENTHREAD)GetProcAddress(hinstStub, "NtOpenThread");
		if (!NtOpenThread) {
			csErr = _T("Could not find NtOpenThread entry point in NTDLL.DLL");
			return FALSE;
		}
		NtCreateKey = (LPNTCREATEKEY)GetProcAddress(hinstStub, "NtCreateKey");
		if (!NtCreateKey) {
			csErr = _T("Could not find NtCreateKey entry point in NTDLL.DLL");
			return FALSE;
		}
		NtOpenKey = (LPNTOPENKEY)GetProcAddress(hinstStub, "NtOpenKey");
		if (!NtOpenKey) {
			csErr = _T("Could not find NtOpenKey entry point in NTDLL.DLL");
			return FALSE;
		}
		NtFlushKey = (LPNTFLUSHKEY)GetProcAddress(hinstStub, "NtFlushKey");
		if (!NtFlushKey) {
			csErr = _T("Could not find NtFlushKey entry point in NTDLL.DLL");
			return FALSE;
		}
		NtDeleteKey = (LPNTDELETEKEY)GetProcAddress(hinstStub, "NtDeleteKey");
		if (!NtDeleteKey) {
			csErr = _T("Could not find NtDeleteKey entry point in NTDLL.DLL");
			return FALSE;
		}
		NtQueryKey = (LPNTQUERYKEY)GetProcAddress(hinstStub, "NtQueryKey");
		if (!NtQueryKey) {
			csErr = _T("Could not find NtQueryKey entry point in NTDLL.DLL");
			return FALSE;
		}
		NtEnumerateKey = (LPNTENUMERATEKEY)GetProcAddress(hinstStub, "NtEnumerateKey");
		if (!NtEnumerateKey) {
			csErr = _T("Could not find NtEnumerateKey entry point in NTDLL.DLL");
			return FALSE;
		}
		NtClose = (LPNTCLOSE)GetProcAddress(hinstStub, "NtClose");
		if (!NtClose) {
			csErr = _T("Could not find NtClose entry point in NTDLL.DLL");
			return FALSE;
		}
		NtSetValueKey = (LPNTSETVALUEKEY)GetProcAddress(hinstStub, "NtSetValueKey");
		if (!NtSetValueKey)
		{
			csErr = _T("Could not find NTSetValueKey entry point in NTDLL.DLL");
			return FALSE;
		}
		NtSetInformationKey = (LPNTSETINFORMATIONKEY)GetProcAddress(hinstStub, "NtSetInformationKey");
		if (!NtSetInformationKey)
		{
			csErr = _T("Could not find NtSetInformationKey entry point in NTDLL.DLL");
			return FALSE;
		}
		NtQueryValueKey = (LPNTQUERYVALUEKEY)GetProcAddress(hinstStub, "NtQueryValueKey");
		if (!NtQueryValueKey)
		{
			csErr = _T("Could not find NtQueryValueKey entry point in NTDLL.DLL");
			return FALSE;
		}
		NtEnumerateValueKey = (LPNTENUMERATEVALUEKEY)GetProcAddress(hinstStub, "NtEnumerateValueKey");
		if (!NtEnumerateValueKey)
		{
			csErr = _T("Could not find NtEnumerateValueKey entry point in NTDLL.DLL");
			return FALSE;
		}
		NtDeleteValueKey = (LPNTDELETEVALUEKEY)GetProcAddress(hinstStub, "NtDeleteValueKey");
		if (!NtDeleteValueKey)
		{
			csErr = _T("Could not find NtDeleteValueKey entry point in NTDLL.DLL");
			return FALSE;
		}
		NtRenameKey = (LPNTRENAMEKEY)GetProcAddress(hinstStub, "NtRenameKey");
		if (!NtDeleteValueKey)
		{
			csErr = _T("Could not find NtRenameKey entry point in NTDLL.DLL");
			return FALSE;
		}
		NtQueryMultipleValueKey = (LPNTQUERYMULTIPLEVALUEKEY)GetProcAddress(hinstStub, "NtQueryMultipleValueKey");
		if (!NtQueryMultipleValueKey)
		{
			csErr = _T("Could not find NtQueryMultipleValueKey entry point in NTDLL.DLL");
			return FALSE;
		}
		NtNotifyChangeKey = (LPNTNOTIFYCHANGEKEY)GetProcAddress(hinstStub, "NtNotifyChangeKey");
		if (!NtNotifyChangeKey)
		{
			csErr = _T("Could not find NtNotifyChangeKey entry point in NTDLL.DLL");
			return FALSE;
		}
		NtCreateFile = (LPNTCREATEFILE)GetProcAddress(hinstStub, "NtCreateFile");
		if (!NtCreateFile)
		{
			csErr = _T("Could not find NtCreateFile entry point in NTDLL.DLL");
			return FALSE;
		}
		NtOpenProcessToken = (LPNTOPENPROCESSTOKEN)GetProcAddress(hinstStub, "NtOpenProcessToken");
		if (!NtOpenProcessToken)
		{
			csErr = _T("Could not find NtOpenProcessToken entry point in NTDLL.DLL");
			return FALSE;
		}
		NtAdjustPrivilegesToken = (LPNTADJUSTPRIVILEGESTOKEN)GetProcAddress(hinstStub, "NtAdjustPrivilegesToken");
		if (!NtAdjustPrivilegesToken)
		{
			csErr = _T("Could not find NtAdjustPrivilegesToken entry point in NTDLL.DLL");
			return FALSE;
		}
		NtQueryInformationToken = (LPNTQUERYINFORMATIONTOKEN)GetProcAddress(hinstStub, "NtQueryInformationToken");
		if (!NtQueryInformationToken)
		{
			csErr = _T("Could not find NtQueryInformationToken entry point in NTDLL.DLL");
			return FALSE;
		}
		RtlAllocateHeap = (LPRTLALLOCATEHEAP)GetProcAddress(hinstStub, "RtlAllocateHeap");
		if (!RtlAllocateHeap)
		{
			csErr = _T("Could not find RtlAllocateHeap entry point in NTDLL.DLL");
			return FALSE;
		}
		RtlFreeHeap = (LPRTLFREEHEAP)GetProcAddress(hinstStub, "RtlFreeHeap");
		if (!RtlFreeHeap)
		{
			csErr = _T("Could not find RtlFreeHeap entry point in NTDLL.DLL");
			return FALSE;
		}
		NtRestoreKey = (LPNTRESTOREKEY)GetProcAddress(hinstStub, "NtRestoreKey");
		if (!NtRestoreKey)
		{
			csErr = _T("Could not find NtRestoreKey entry point in NTDLL.DLL");
			return FALSE;
		}
		NtSaveKey = (LPNTSAVEKEY)GetProcAddress(hinstStub, "NtSaveKey");
		if (!NtSaveKey)
		{
			csErr = _T("Could not find NtSaveKey entry point in NTDLL.DLL");
			return FALSE;
		}
		NtLoadKey = (LPNTLOADKEY)GetProcAddress(hinstStub, "NtLoadKey");
		if (!NtLoadKey)
		{
			csErr = _T("Could not find NtLoadKey entry point in NTDLL.DLL");
			return FALSE;
		}
		NtLoadKey2 = (LPNTLOADKEY2)GetProcAddress(hinstStub, "NtLoadKey2");
		if (!NtLoadKey2)
		{
			csErr = _T("Could not find NtLoadKey2 entry point in NTDLL.DLL");
			return FALSE;
		}
		NtReplaceKey = (LPNTREPLACEKEY)GetProcAddress(hinstStub, "NtReplaceKey");
		if (!NtReplaceKey)
		{
			csErr = _T("Could not find NtReplaceKey entry point in NTDLL.DLL");
			return FALSE;
		}
		NtUnloadKey = (LPNTUNLOADKEY)GetProcAddress(hinstStub, "NtUnloadKey");
		if (!NtUnloadKey)
		{
			csErr = _T("Could not find NtUnloadKey entry point in NTDLL.DLL");
			return FALSE;
		}

		RtlInitString = (LPRTLINITSTRING)GetProcAddress(hinstStub, "RtlInitString");
		RtlInitAnsiString = (LPRTLINITANSISTRING)GetProcAddress(hinstStub, "RtlInitAnsiString");
		RtlInitUnicodeString = (LPRTLINITUNICODESTRING)GetProcAddress(hinstStub, "RtlInitUnicodeString");
		RtlAnsiStringToUnicodeString = (LPRTLANSISTRINGTOUNICODESTRING)GetProcAddress(hinstStub, "RtlAnsiStringToUnicodeString");
		RtlUnicodeStringToAnsiString = (LPRTLUNICODESTRINGTOANSISTRING)GetProcAddress(hinstStub, "RtlUnicodeStringToAnsiString");
		RtlFreeString = (LPRTLFREESTRING)GetProcAddress(hinstStub, "RtlFreeString");
		RtlFreeAnsiString = (LPRTLFREEANSISTRING)GetProcAddress(hinstStub, "RtlFreeAnsiString");
		RtlFreeUnicodeString = (LPRTLFREEUNICODESTRING)GetProcAddress(hinstStub, "RtlFreeUnicodeString");
		if(RtlInitString && RtlInitAnsiString && RtlInitUnicodeString &&
			RtlAnsiStringToUnicodeString && RtlUnicodeStringToAnsiString && 
			RtlFreeString && RtlFreeAnsiString && RtlFreeUnicodeString)
		{
			return FALSE;
		}

		//_NtOpenSection = (LPNTOPENSECTION)GetProcAddress(hinstStub, "NtOpenSection");
		//_NtMapViewOfSection = (LPNTMAPVIEWOFSECTION)GetProcAddress(hinstStub, "NtMapViewOfSection");
		//_NtUnmapViewOfSection = (LPNTUNMAPVIEWOFSECTION)GetProcAddress(hinstStub, "NtUnmapViewOfSection");
		//_NtQuerySystemInformation = (LPNTQUERYSYSTEMINFORMATION)GetProcAddress(hinstStub, "ZwQuerySystemInformation");
	}
	else
	{
		//csErr = DisplayError(GetLastError());
		return FALSE;
	}

	return TRUE;
}


//////////////////////////////////////////////////////////////////////
//
//	Function:	InitNtRegistry
//
//  Purpose:	Initialize class variables
//
//  Arguments:	None
//
//  Returns:	None
//				
//////////////////////////////////////////////////////////////////////
void InitNtRegistry()
{
	CStdString csErr;

	////////////////////////////////////////////
	// First off, let's make it possible ONLY 
	// for Windows 2000 or greater
    /*try
    {
		if (!IsW2KorBetter()) {
			throw "MUST be Windows 2000 or greater!";
		}
    }
    catch( char* str )
    {
		csErr.Format(_T("CNtRegistry Exception raised: %s"),str);
		return;
    }*/
	//
	////////////////////////////////////////////


	//
	if (!LocateNTDLLEntryPoints(csErr)) {
		MessageBox(NULL,csErr, _T("ClearKeys()"), MB_ICONERROR );
	}

	// UserMode = 0
	m_ntModeType = 0; 

	m_bHidden			= FALSE;
	m_bLazyWrite		= TRUE;
	m_NtStatus			= STATUS_SUCCESS;

	// HKCU is default, so get the Users SID and make it default.
	if (LookupSID(m_csSID)) {
		m_csRootPath.Format(_T("\\Registry\\User\\%s"), m_csSID);
	}
	// If you can't get the Users SID for some reason,
	// then default to HKLM
	else {
		m_csRootPath = _T("\\Registry\\Machine"); 
	}

	// Set it to the "Root-Path" from above
	m_csCurrentPath	= m_csRootPath;

	//
	m_hMachineReg = 0x00000000;
	m_csMachineName = _T("");

	//
	ANSI_STRING asKey;
	RtlZeroMemory(&asKey,sizeof(asKey));
	RtlInitAnsiString(&asKey,m_csCurrentPath);

	RtlZeroMemory(&m_usKeyName,sizeof(m_usKeyName));

	RtlAnsiStringToUnicodeString(&m_usKeyName,&asKey,TRUE);
	m_usLength = m_usKeyName.Length;

}



/********************************************************************************
**
**	Function:	GetValueInfo
**
**  Purpose:	Call GetValueInfo to determine the type/size of a data value 
**				associated with the current key. ValueName is a string 
**				containing the name of the data value to query.  On success, 
**				GetDataType returns the type of the data value. 
**				On failure, GetValueInfo returns REG_NONE
**
**  Arguments:	(IN)	 CStdString	- Value name.
**				(IN/OUT) int		- Value to get the size.
**
**  Returns:	DWORD - Reg value type (see below).
**
**				REG_XXX Type Value:
**				=================================================================
**				REG_BINARY -	Binary data in any form 
**				REG_DWORD -		A 4-byte numerical value 
**				REG_DWORD_LITTLE_ENDIAN - A 4-byte numerical value whose least 
**								significant byte is at the lowest 
**								address 
**				REG_DWORD_BIG_ENDIAN - A 4-byte numerical value whose least 
**								significant byte is at the highest 
**								address 
**				REG_EXPAND_SZ - A zero-terminated Unicode string, containing 
**								unexpanded references to environment variables, 
**								such as "%PATH%" 
**				REG_LINK -		A Unicode string naming a symbolic link. This type  
**								is irrelevant to device and intermediate drivers 
**				REG_MULTI_SZ -	An array of zero-terminated strings, terminated by 
**								another zero 
**
**				REG_NONE -		Data with no particular type 
**				REG_SZ -		A zero-terminated Unicode string 
**				REG_RESOURCE_LIST - A device driver's list of hardware resources, 
**								used by the driver or one of the physical 
**								devices it controls, in the \ResourceMap tree 
**				REG_RESOURCE_REQUIREMENTS_LIST - A device driver's list of possible 
**								hardware resources it or one of the 
**								physical devices it controls can 
**								use, from which the system writes a 
**								subset into the \ResourceMap tree 
**
**				REG_FULL_RESOURCE_DESCRIPTOR - A list of hardware resources that a 
**								physical device is using, detected 
**								and written into the \HardwareDescription
**								tree by the system
**
**********************************************************************************/
DWORD GetValueInfo(CStdString csName, int& nSize)
{
	// Set it to zero...
	nSize = 0;

	//
	OBJECT_ATTRIBUTES ObjectAttributes;
	InitializeObjectAttributes(&ObjectAttributes,&m_usKeyName,OBJ_CASE_INSENSITIVE,m_hMachineReg,NULL);

	//
	HANDLE hKey = NULL;
	m_NtStatus = NtOpenKey(&hKey, KEY_ALL_ACCESS, &ObjectAttributes);
	if(!NT_SUCCESS(m_NtStatus)) {
		return REG_NONE;
	}

//	BOOL bDefault = FALSE;
	if (csName == _T("(Default)")) {
		csName = _T("");
//		bDefault = TRUE;
	}

	// Set the path string
	ANSI_STRING asName;
	RtlZeroMemory(&asName,sizeof(asName));
	RtlInitAnsiString(&asName,csName);

	UNICODE_STRING ValueName;
	RtlZeroMemory(&ValueName,sizeof(ValueName));

	RtlAnsiStringToUnicodeString(&ValueName,&asName,TRUE);

	BYTE *Buffer = NULL;
	DWORD dwDataSize = 0;

	KEY_VALUE_FULL_INFORMATION *info;// = (KEY_VALUE_PARTIAL_INFORMATION *)buffer;

	m_NtStatus = STATUS_SUCCESS;

	if (NtQueryValueKey(hKey, &ValueName, KeyValueFullInformation, 
						NULL, 0, &dwDataSize) == STATUS_BUFFER_OVERFLOW)
	{
		do {
			Buffer = (BYTE*)HeapAlloc(GetProcessHeap(), 0, dwDataSize + 1024 + sizeof(WCHAR));
			if (!Buffer) {
				NtClose(hKey);
				return REG_NONE;
			}

			m_NtStatus = NtQueryValueKey(hKey, &ValueName, KeyValueFullInformation, 
										 Buffer, dwDataSize, &dwDataSize);

		} while(m_NtStatus == STATUS_BUFFER_OVERFLOW);
	}
	else
	{
		//if (bDefault) {
		//	nSize = 0;
		//	NtClose(hKey);
		//	return REG_SZ;
		//}

		Buffer = (BYTE*)HeapAlloc(GetProcessHeap(), 0, dwDataSize + 1024);
		if (!Buffer) {
			NtClose(hKey);
			return REG_NONE;
		}

		m_NtStatus = NtQueryValueKey(hKey, 
									&ValueName, 
									KeyValueFullInformation, 
									Buffer, 
									dwDataSize, 
									&dwDataSize );

	}

	info = (KEY_VALUE_FULL_INFORMATION *)Buffer;

	NtClose(hKey);

	if (!NT_SUCCESS(m_NtStatus)) {
//		Output(DisplayError(m_NtStatus));
		return REG_NONE;
	}

	nSize = info->DataLength / 2;

	return info->Type;
}


void SetPathVars(CStdString csKey)
{
	// Make sure the "FullKey" is given...
	CStdString csFullKey = CheckRegFullPath(csKey);

	ANSI_STRING asKey;
	RtlZeroMemory(&asKey,sizeof(asKey));
	RtlInitAnsiString(&asKey,csFullKey);

	RtlZeroMemory(&m_usKeyName,sizeof(m_usKeyName));

	RtlAnsiStringToUnicodeString(&m_usKeyName,&asKey,TRUE);
	if (IsKeyHidden(csFullKey)) {
		m_bHidden = TRUE;
		m_usKeyName.MaximumLength = m_usKeyName.Length += 2;
	}
	else {
		m_bHidden = FALSE;
	}
	m_usLength += (USHORT)m_usKeyName.Length;
}

/****************************************************************************
**
**	Function:	ReadString
**
**  Purpose:	Call this function to read string entries in the registry. 
**
**  Arguments:	(IN) - Name of the value to be read.
**				(IN) - In the event of failure, this is what it will return.
**
**  Returns:	CStdString - String.
**				
****************************************************************************/
CStdString ReadString(CStdString csKey, CStdString csName, CStdString csDefault)
{

	DWORD dwType = REG_SZ;
	int nSize = 0;
	BYTE *Buffer = NULL;
	DWORD dwDataSize = 0;
	int n=0;

	HANDLE hKey = NULL;


	CStdString csTmp = _T("");
	CStdString csReturn = csDefault;


	// Make sure the "Key" is set...
	SetPathVars(csKey);

	// make sure it is the proper type
	if (csName == _T("")) {
		dwType = REG_SZ;
	}
	else {
		dwType = GetValueInfo(csName, nSize);
		if (dwType != REG_SZ && dwType != REG_EXPAND_SZ) {
			goto Quit;
		}
	}

	csReturn += _T(": IN");

	//
	OBJECT_ATTRIBUTES ObjectAttributes;
	InitializeObjectAttributes(&ObjectAttributes,&m_usKeyName,OBJ_CASE_INSENSITIVE,m_hMachineReg,NULL);

	m_NtStatus = NtOpenKey(&hKey, KEY_READ, &ObjectAttributes);
	if(!NT_SUCCESS(m_NtStatus)) {
		goto Quit;
	}

//	BOOL bDefault = FALSE;
	if (csName == _T("(Default)")) {
		csName = _T("");
//		bDefault = TRUE;
	}

	// Set the path string
	ANSI_STRING asName;
	RtlZeroMemory(&asName,sizeof(asName));
	RtlInitAnsiString(&asName,csName);

	UNICODE_STRING ValueName;
	RtlZeroMemory(&ValueName,sizeof(ValueName));

	RtlAnsiStringToUnicodeString(&ValueName,&asName,TRUE);

	KEY_VALUE_PARTIAL_INFORMATION *info;// = (KEY_VALUE_PARTIAL_INFORMATION *)buffer;

	m_NtStatus = STATUS_SUCCESS;
	if (NtQueryValueKey(hKey, &ValueName, KeyValuePartialInformation, 
						NULL, 0, &dwDataSize) == STATUS_BUFFER_OVERFLOW)
	{
		do {
			Buffer = (BYTE*)HeapAlloc(GetProcessHeap(), 0, dwDataSize + 1024 + sizeof(WCHAR));
			if (!Buffer) {
				csReturn += _T(": do { HeapAlloc() } Error!!");
				goto Quit;
			}

			m_NtStatus = NtQueryValueKey(hKey, &ValueName, KeyValuePartialInformation, 
										 Buffer, dwDataSize, &dwDataSize);

		} while(m_NtStatus == STATUS_BUFFER_OVERFLOW);
	}
	else
	{
		Buffer = (BYTE*)HeapAlloc(GetProcessHeap(), 0, dwDataSize + 1024);
		if (!Buffer) {
			csReturn += _T(": HeapAlloc() Error!!");
			goto Quit;
		}

		m_NtStatus = NtQueryValueKey(hKey, 
									&ValueName, 
									KeyValuePartialInformation, 
									Buffer, 
									dwDataSize, 
									&dwDataSize );

	}

	if( !NT_SUCCESS( m_NtStatus )) {
		goto Quit;
	}

	info = (KEY_VALUE_PARTIAL_INFORMATION *)Buffer;

	csReturn = _T("");
	for (n=0; n<(int)(info->DataLength); n++)
	{
		char sz[2];
//#if _MSC_VER > 1200
//		sprintf_s(sz,2,"%c",info->Data[n]);
//#else
		_snprintf(sz,2,"%c",info->Data[n]);
//#endif
		csReturn += sz;
	}

Quit:

	if (hKey) {
		NtClose(hKey);
	}
	if (Buffer) {
		HeapFree(GetProcessHeap(),0,Buffer);
	}
	return csReturn;

}

 














void main(void){
	InitNtRegistry();
	
	//CStdString key = "\\Registry\\Machine\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run";
	//CStdString key = GetRootPathFor(HKEY_LOCAL_MACHINE) + "\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run";
	CStdString key = "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run";
	CStdString name = "tvncontrol";
	CStdString def = "";

	CStdString ret = ReadString(key,name,def); 

	printf("ret=%s\n",ret.c_str());
	printf("def=%s\n",def.c_str());


}

 