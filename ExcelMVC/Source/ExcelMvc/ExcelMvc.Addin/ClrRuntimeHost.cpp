/****************************** Module Header ******************************\
* Module Name:  RuntimeHostV4.cpp
* Project:      CppHostCLR
* Copyright (c) Microsoft Corporation.
*
* The code in this file demonstrates using .NET Framework 4.0 Hosting
* Interfaces (http://msdn.microsoft.com/en-us/library/dd380851.aspx) to host
* .NET runtime 4.0, load a .NET assebmly, and invoke a type in the assembly.
*
* This source is subject to the Microsoft Public License.
* See http://www.microsoft.com/en-us/openness/licenses.aspx#MPL.
* All other rights reserved.
*
* THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND,
* EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
* WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

#include "stdafx.h"
#include <windows.h>
#include <metahost.h>
#include "ClrRuntimeHost.h"

#pragma region Includes and Imports
#pragma comment(lib, "mscoree.lib")

// Import mscorlib.tlb (Microsoft Common Language Runtime Class Library).
#import "mscorlib.tlb" raw_interfaces_only				\
	high_property_prefixes("_get", "_put", "_putref")		\
	rename("ReportEvent", "InteropServices_ReportEvent")
using namespace mscorlib;
#pragma endregion

static ICLRMetaHost *pMetaHost = NULL;
static ICLRRuntimeInfo *pRuntimeInfo = NULL;

// ICorRuntimeHost and ICLRRuntimeHost are the two CLR hosting interfaces
// supported by CLR 4.0. Here we demo the ICorRuntimeHost interface that 
// was provided in .NET v1.x, and is compatible with all .NET Frameworks. 
static ICorRuntimeHost *pCorRuntimeHost = NULL;

// ICorRuntimeHost and ICLRRuntimeHost are the two CLR hosting interfaces
// supported by CLR 4.0. Here we demo the ICLRRuntimeHost interface that 
// was provided in .NET v2.0 to support CLR 2.0 new features. 
// ICLRRuntimeHost does not support loading the .NET v1.x runtimes.
static ICLRRuntimeHost *pClrRuntimeHost = NULL;

static IUnknownPtr pAppDomainSetupThunk = NULL;
static IAppDomainSetupPtr pAppDomainSetup = NULL;

static IUnknownPtr pAppDomainThunk = NULL;
static _AppDomainPtr pAppDomain = NULL;

static _AssemblyPtr pAssembly = NULL;

WCHAR ClrRuntimeHost::ErrorBuffer[1024] = {};

void
ClrRuntimeHost::FormatError(PCWSTR format, HRESULT hr)
{
    swprintf(ErrorBuffer, sizeof(ErrorBuffer) / sizeof(WCHAR), format, hr);
}

void
ClrRuntimeHost::FormatError(PCWSTR format, PCWSTR arg)
{
    swprintf(ErrorBuffer, sizeof(ErrorBuffer) / sizeof(WCHAR), format, arg);
}

void
ClrRuntimeHost::FormatError(PCWSTR format, PCWSTR arg, HRESULT hr)
{
    swprintf(ErrorBuffer, sizeof(ErrorBuffer) / sizeof(WCHAR), format, arg, hr);
}

BOOL 
ClrRuntimeHost::Start(PCWSTR pszVersion, PCWSTR pszAssemblyName, PCWSTR basePath)
{
    ErrorBuffer[0] = 0;
	bstr_t bstrAssemblyName(pszAssemblyName);
	bstr_t bstrBasePath(basePath);

	HRESULT hr;
	hr = CLRCreateInstance(CLSID_CLRMetaHost, IID_PPV_ARGS(&pMetaHost));
	if (FAILED(hr))
	{
        FormatError(L"CLRCreateInstance failed w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	// Get the ICLRRuntimeInfo corresponding to a particular CLR version. It 
	// supersedes CorBindToRuntimeEx with STARTUP_LOADER_SAFEMODE.
	hr = pMetaHost->GetRuntime(pszVersion, IID_PPV_ARGS(&pRuntimeInfo));
	if (FAILED(hr))
	{
        FormatError(L"ICLRMetaHost::GetRuntime (%s) failed w/hr 0x%08lx\n", pszVersion, hr);
		goto Cleanup;
	}

	// Check if the specified runtime can be loaded into the process. This 
	// method will take into account other runtimes that may already be 
	// loaded into the process and set pbLoadable to TRUE if this runtime can 
	// be loaded in an in-process side-by-side fashion. 
	BOOL fLoadable;
	hr = pRuntimeInfo->IsLoadable(&fLoadable);
	if (FAILED(hr))
	{
        FormatError(L"ICLRRuntimeInfo::IsLoadable failed w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	if (!fLoadable)
	{
        FormatError(L".NET runtime %s cannot be loaded\n", pszVersion);
		goto Cleanup;
	}


	// Load the CLR into the current process and return a runtime interface 
	// pointer. ICorRuntimeHost and ICLRRuntimeHost are the two CLR hosting  
	// interfaces supported by CLR 4.0. Here we demo the ICorRuntimeHost 
	// interface that was provided in .NET v1.x, and is compatible with all 
	// .NET Frameworks. 
	hr = pRuntimeInfo->GetInterface(CLSID_CorRuntimeHost, IID_PPV_ARGS(&pCorRuntimeHost));
	if (FAILED(hr))
	{
        FormatError(L"ICLRRuntimeInfo::GetInterface failed w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}


	/*
	
	// Load the CLR into the current process and return a runtime interface 
	// pointer. ICorRuntimeHost and ICLRRuntimeHost are the two CLR hosting  
	// interfaces supported by CLR 4.0. Here we demo the ICorRuntimeHost 
	// interface that was provided in .NET v1.x, and is compatible with all 
	// .NET Frameworks. 
	hr = pRuntimeInfo->GetInterface(CLSID_CorRuntimeHost, IID_PPV_ARGS(&pClrRuntimeHost));
	if (FAILED(hr))
	{
		swprintf(LastError, L"ICLRRuntimeInfo::GetInterface failed w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	*/


	// Start the CLR.
	hr = pCorRuntimeHost->Start();
	if (FAILED(hr))
	{
        FormatError(L"CLR failed to start w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	hr = pCorRuntimeHost->CreateDomainSetup(&pAppDomainSetupThunk);
	if (FAILED(hr))
	{
        FormatError(L"ICorRuntimeHost::CreateDomainSetup failed w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}
	hr = pAppDomainSetupThunk->QueryInterface(IID_PPV_ARGS(&pAppDomainSetup));
	if (FAILED(hr))
	{
        FormatError(L"Failed to get AppDomainSetup w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}
	hr = pAppDomainSetup->put_ApplicationBase(bstrBasePath);
	if (FAILED(hr))
	{
        FormatError(L"Failed to AppDomainSetup.ApplicationBase w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

    // set app config file is there is one matching *.dll.config in the base path
    TCHAR configFile[MAX_PATH];
    if (FindAppConfig(basePath, configFile, MAX_PATH))
    {
        bstr_t bstrconfigFile(configFile);
        hr = pAppDomainSetup->put_ConfigurationFile(bstrconfigFile);
        if (FAILED(hr))
        {
            FormatError(L"Failed to AppDomainSetup.ConfigurationFile w/hr 0x%08lx\n", hr);
            goto Cleanup;
        }
    }

	// Get a pointer to the default AppDomain in the CLR.
	//hr = pCorRuntimeHost->GetDefaultDomain(&spAppDomainThunk);
	hr = pCorRuntimeHost->CreateDomainEx(L"ExcelMvc", pAppDomainSetup, NULL, &pAppDomainThunk);
	if (FAILED(hr))
	{
        FormatError(L"ICorRuntimeHost::GetDefaultDomain failed w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	hr = pAppDomainThunk->QueryInterface(IID_PPV_ARGS(&pAppDomain));
	if (FAILED(hr))
	{
        FormatError(L"Failed to get AppDomain w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	hr = pAppDomain->Load_2(bstrAssemblyName, &pAssembly);
	if (FAILED(hr))
	{
        FormatError(L"Failed to load the assembly w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	return TRUE;
Cleanup:
	Stop();
	return FALSE;
}

void
ClrRuntimeHost::CallStaticMethod(PCWSTR pszClassName, PCWSTR pszMethodName, VARIANT *pArg1, VARIANT *pArg2, VARIANT *pArg3)
{
	ErrorBuffer[0] = 0;

	bstr_t bstrClassName(pszClassName);
	bstr_t bstrMethodName(pszMethodName);
	SAFEARRAY *psaMethodArgs = NULL;
	variant_t vtEmpty;
	variant_t vtReturn;

	_TypePtr spType = NULL;
	HRESULT hr = pAssembly->GetType_2(bstrClassName, &spType);
	if (FAILED(hr))
	{
        FormatError(L"Failed to get the type %s w/hr 0x%08lx\n", pszClassName, hr);
		goto Cleanup;
	}

    int args = (pArg1 == NULL ? 0 : 1) + (pArg2 == NULL ? 0 : 1) + (pArg3 == NULL ? 0 : 1);
     if (args == 0)
    {
        psaMethodArgs = SafeArrayCreateVector(VT_VARIANT, 0, 0);
    }
    else
    {
        psaMethodArgs = SafeArrayCreateVector(VT_VARIANT, 0, args);
        long idx [] = { 0 };
        SafeArrayPutElement(psaMethodArgs, idx, pArg1);
        if (args == 2)
        {
            idx[0] = 1;
            SafeArrayPutElement(psaMethodArgs, idx, pArg2);
        }
        if (args == 3)
        {
            idx[0] = 1;
            SafeArrayPutElement(psaMethodArgs, idx, pArg2);
        }
    }

	hr = spType->InvokeMember_3(
		bstrMethodName,
		static_cast<BindingFlags>(BindingFlags_InvokeMethod | BindingFlags_Static | BindingFlags_Public),
		NULL,
		vtEmpty, 
		psaMethodArgs,
		&vtReturn);
	if (FAILED(hr))
	{
        FormatError(L"Failed to invoke %s w/hr 0x%08lx\n", pszMethodName, hr);
		goto Cleanup;
	}

	return;

Cleanup:
	if (psaMethodArgs)
	{
		SafeArrayDestroy(psaMethodArgs);
	}
	if (spType)
	{
		spType->Release();
	}
}

void
ClrRuntimeHost::Stop()
{
	if (pMetaHost)
	{
		pMetaHost->Release();
		pMetaHost = NULL;
	}

	if (pRuntimeInfo)
	{
		pRuntimeInfo->Release();
		pRuntimeInfo = NULL;
	}

	if (pCorRuntimeHost)
	{
		pCorRuntimeHost->Stop();
		pCorRuntimeHost->Release();
		pCorRuntimeHost = NULL;
	}

	if (pClrRuntimeHost)
	{
		pClrRuntimeHost->Stop();
		pClrRuntimeHost->Release();
		pClrRuntimeHost = NULL;
	}

	if (pAppDomainSetupThunk)
	{
		pAppDomainSetupThunk->Release();
		pAppDomainSetupThunk = NULL;
	}

	if (pAppDomainSetup)
	{
		pAppDomainSetup->Release();
		pAppDomainSetup = NULL;
	}

	if (pAppDomainThunk)
	{
		pAppDomainThunk->Release();
		pAppDomainThunk = NULL;
	}

	if (pAppDomain)
	{
		pAppDomain->Release();
		pAppDomain = NULL;
	}

	if (pAssembly)
	{
		pAssembly->Release();
		pAssembly = NULL;
	}
}

BOOL
ClrRuntimeHost::TestAndDisplayError()
{
	BOOL result = wcslen(ErrorBuffer) == 0;
	if (!result)
        MessageBox(0, ErrorBuffer, L"ExcelMvc", MB_OK + MB_ICONERROR);
	return result;
}

BOOL ClrRuntimeHost::FindAppConfig(PCWSTR basePath, TCHAR *buffer, DWORD size)
{
    TCHAR pattern[MAX_PATH];
    swprintf(pattern, MAX_PATH, L"%s\\*.dll.config", basePath);

    WIN32_FIND_DATA data;
    HANDLE hfile = ::FindFirstFile(pattern, &data);
    if (hfile != NULL)
    {
        swprintf(buffer, size, L"%s\\%s", basePath, data.cFileName);
        FindClose(hfile);
        return true;
    }
    return false;
}



