// ATPSO.cpp : Implementation of DLL Exports.


#include "stdafx.h"
#include "resource.h"
#include "ATPSO_i.h"
#include "dllmain.h"
#include "SmartOccurrence.h"	// IJSmartOccurrence
#include "SmartOccurrence_i.c"
#include "SblEntities.h"		// IJDUserSymbolServices
#include "SblEntities_i.c"
typedef USHORT FLAVORID, *LPFLAVORID ;
#include "isymbol.h"
#include "CustomAssembly.h"		// IJCAFactory
#include "CustomAssembly_i.c"
#include "ErrorLog.h"			// IJEditErrors
#include "ErrorLog_i.c"
#include "Geom3d.h"
#include "Geom3d_i.c"
#include "Attributes.h"			// IJDAttributes
#include "Attributes_i.c"
#include "IMSDObject.h"			// IJDObject
#include "IMSDObject_i.c"
#include "Revision.h"
#include "Revision_i.c"

// Used to determine whether the DLL can be unloaded by OLE
STDAPI DllCanUnloadNow(void)
{
    return _AtlModule.DllCanUnloadNow();
}


// Returns a class factory to create an object of the requested type
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
    return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}


// DllRegisterServer - Adds entries to the system registry
STDAPI DllRegisterServer(void)
{
    // registers object, typelib and all interfaces in typelib
    HRESULT hr = _AtlModule.DllRegisterServer();
	return hr;
}


// DllUnregisterServer - Removes entries from the system registry
STDAPI DllUnregisterServer(void)
{
	HRESULT hr = _AtlModule.DllUnregisterServer();
	return hr;
}

// DllInstall - Adds/Removes entries to the system registry per user
//              per machine.	
STDAPI DllInstall(BOOL bInstall, LPCWSTR pszCmdLine)
{
    HRESULT hr = E_FAIL;
    static const wchar_t szUserSwitch[] = (L"user");

    if (pszCmdLine != NULL)
    {
    	if (_wcsnicmp(pszCmdLine, szUserSwitch, _countof(szUserSwitch)) == 0)
    	{
    		AtlSetPerUserRegistration(true);
    	}
    }

    if (bInstall)
    {	
    	hr = DllRegisterServer();
    	if (FAILED(hr))
    	{	
    		DllUnregisterServer();
    	}
    }
    else
    {
    	hr = DllUnregisterServer();
    }

    return hr;
}


