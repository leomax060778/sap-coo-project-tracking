#include "stdafx.h"
#include <windows.h>
#include "unknwn.h"
#include <string>

#include <atlbase.h>
#include <atlcom.h>
#include <initguid.h>
#include <comdef.h>



#include "ProfManLoader.h"

using namespace std;


#import "libid:A53ADCA6-5BD6-475F-9C8A-5C084E3A09E3" rename_namespace("ProfMan") raw_interfaces_only no_auto_exclude//include("IPropertyBag")


RTL_CRITICAL_SECTION ProfManLoader::CritSection;
std::wstring ProfManLoader::DllLocation32Bit;
std::wstring ProfManLoader::DllLocation64Bit;
HMODULE ProfManLoader::DllHandle;
bool _profman_initialized = false;


 void ProfManLoader::CheckInitialize()
 {
	 if(!_profman_initialized)
	 {
		 _profman_initialized = true;
		 InitializeCriticalSection(&CritSection);
		 EnterCriticalSection(&CritSection);
		 DllHandle = 0;
		 LeaveCriticalSection(&CritSection);
	 }

 }

 ProfManLoader::~ProfManLoader()
 {
	 EnterCriticalSection(&CritSection);
     __try
     {
        if (DllHandle) FreeLibrary(DllHandle);
     }
     __finally
     {
       LeaveCriticalSection(&CritSection);
     }
     DeleteCriticalSection(&CritSection);
 }

 void ProfManLoader::SetDllLocation32Bit(LPWSTR Value)
 {
	 ProfManLoader::CheckInitialize(); //make sure we are initialized
	 EnterCriticalSection(&CritSection);
	  __try
     {
        DllLocation32Bit = Value;
     }
     __finally
     {
       LeaveCriticalSection(&CritSection);
     }
 }

void ProfManLoader::SetDllLocation64Bit(LPWSTR Value)
{
	 ProfManLoader::CheckInitialize(); //make sure we are initialized
	 EnterCriticalSection(&CritSection);
	 __try
     {
        DllLocation64Bit = Value;
     }
     __finally
     {
       LeaveCriticalSection(&CritSection);
     }
}

 template <typename T>
 T* __NewProfManObject(const GUID CLSID)
 {
	 CComPtr<IUnknown> unk;
	 T* res = NULL;
	 unk = ProfManLoader::NewProfManObject(CLSID);
	 ATLASSERT(unk);
	 if (unk)
	 {
		unk->QueryInterface(__uuidof(T), (void**)&res);
	 }
	 return res;
 };


 ProfMan::IProfiles* ProfManLoader::new_Profiles()
 {
	 return __NewProfManObject<ProfMan::IProfiles>(__uuidof(ProfMan::Profiles));
 }

   ProfMan::IPropertyBag* ProfManLoader::new_PropertyBag()
 {
	 return __NewProfManObject<ProfMan::IPropertyBag>(__uuidof(ProfMan::PropertyBag));
 }
  

void ProfManLogLastError(wchar_t* name)
{
	DWORD __errorCode = GetLastError();									
	wchar_t* errmsg = NULL;													
	FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, 0, 
					__errorCode, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (wchar_t*)&errmsg, 0, NULL);									
	wchar_t msg[1024];													
	swprintf(&msg[0], L"Function=%s, GetLastError()=0x%08x, msg=\"%s\"", name, __errorCode, (wchar_t*)errmsg);	
	OutputDebugStringW(&msg[0]);												
	LocalFree(errmsg);	

}


 IUnknown* ProfManLoader::NewProfManObject(const GUID CLSID)
 {
	 IUnknown* result = NULL;
	 HRESULT res;

	 ProfManLoader::CheckInitialize(); //make sure we are initialized

	 EnterCriticalSection(&CritSection);

	 if (!DllHandle)
	 {
		 if ((!DllLocation32Bit.length()) || (!DllLocation64Bit.length()))
		 {
			 //AFX_MANAGE_STATE(AfxGetStaticModuleState());
			 //HINSTANCE hInstance = AfxGetInstanceHandle();
			 HINSTANCE hInstance = 0;
		  	 wchar_t buffer[MAX_PATH];
			 int len = GetModuleFileNameW(hInstance, buffer, MAX_PATH);
			 if (len > 0)
			 {
			 	wstring path = buffer;
				size_t pos = path.find_last_of(L"/\\");
				if (pos > 0) path.resize(pos+1); else path = L"";
				if (!DllLocation32Bit.length())  
				{
					DllLocation32Bit = path; 
					DllLocation32Bit += wstring(L"ProfMan.dll");
				}
				if (!DllLocation64Bit.length())  
				{
					DllLocation64Bit = path; 
					DllLocation64Bit += wstring(L"ProfMan64.dll");
				}
			 }
		 }
		 ATLASSERT(DllLocation32Bit.length() && DllLocation64Bit.length());

		 wstring dllName;
		 (sizeof(void*) == 8)  ? dllName = DllLocation64Bit : dllName = DllLocation32Bit; 
		 DllHandle = LoadLibraryW((LPCWSTR)dllName.c_str());
		 if (!DllHandle) 
		 {
			//wchar_t error[2 * MAX_PATH];
			//wsprintf(error, L"Could not load ProfMan dll at \"%s\"\n", dllName.c_str());
			//OutputDebugStringW(error);
			ProfManLogLastError(L"LoadLibraryW");
		 }
		 ATLASSERT(DllHandle);
	}

	if (DllHandle) 
	{
		BOOL (WINAPI*DllGetClassObject)(REFCLSID,REFIID,LPVOID) =  (BOOL(WINAPI*)(REFCLSID,REFIID,LPVOID)) GetProcAddress(DllHandle, "DllGetClassObject");
		if (!DllGetClassObject) ProfManLogLastError(L"GetProcAddress");
		ATLASSERT(DllGetClassObject);

		if (DllGetClassObject) 
		{
			CComPtr<IClassFactory> classFactory = NULL;
			res = DllGetClassObject(CLSID, __uuidof(IClassFactory), &classFactory);
			ATLASSERT(S_OK == res);
			if (SUCCEEDED(res)) 
			{
					
				ATLASSERT(classFactory);

				res = classFactory->CreateInstance(NULL, __uuidof(IUnknown), (void**)&result);
				ATLASSERT(S_OK == res);

				if (FAILED(res))
		 		if (CO_E_WRONGOSFORAPP == res)
		 		{
					OutputDebugStringW(L"IClassFactory.CreateInstance returned CO_E_WRONGOSFORAPP.\nMake sure the bitness of your application and the MAPI system/Outlook match\n");
		 		}
				else
				{
					wchar_t error[128];
					wsprintf(error, L"IClassFactory.CreateInstance returned %08x\n", res);
					OutputDebugStringW(error);
				}
			} 
		}

		 
	 }
     LeaveCriticalSection(&CritSection);
	 return result;

 }