#include <windows.h>
#include <string>

#import "libid:A53ADCA6-5BD6-475F-9C8A-5C084E3A09E3" rename_namespace("ProfMan") raw_interfaces_only no_auto_exclude //include("IPropertyBag")


 class ProfManLoader
 {
 private:
     static RTL_CRITICAL_SECTION CritSection;
     static std::wstring DllLocation32Bit;
     static std::wstring DllLocation64Bit;
     static HMODULE DllHandle;

     static void CheckInitialize();
     ~ProfManLoader();

 public:
	 static IUnknown* NewProfManObject(const GUID CLSID);

	 static ProfMan::IProfiles* new_Profiles();
	 static ProfMan::IPropertyBag* new_PropertyBag();

	 //dll locations
	 static void SetDllLocation32Bit(LPWSTR Value);
	 static void SetDllLocation64Bit(LPWSTR Value);

 };



