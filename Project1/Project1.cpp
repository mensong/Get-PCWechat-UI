// Project1.cpp : 此文件包含 "main" 函数。程序执行将在此处开始并结束。
//

#include "pch.h"
#include <iostream>

#include<tchar.h>
#include <windows.h>
#include <oleacc.h>
#include <stdio.h>
#include <comutil.h>
#include <Shlwapi.h>
#include <string>
#include <commctrl.h>
#include <vector>
#include <set>
#include <atlcomcli.h>

//#include<initguid.h>///As with any COM interface, the system file initguid.h should be included in any source code that requires Active Accessibility.
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "oleacc.lib")
#pragma comment(lib, "comsuppw.lib")
#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib,"kernel32.lib")
//
#include "comutil.h"  
#pragma comment(lib, "comsuppw.lib")  
#pragma comment(lib,"comsuppwd.lib")  


UINT GetObjectState(IAccessible* pacc, VARIANT* pvarChild, LPTSTR lpszBuff, UINT sizeBuff)
{
    HRESULT hr;
    VARIANT varRetVal;
    //*lpszState = 0;
    memset(lpszBuff, 0, sizeBuff * sizeof(lpszBuff[0]));
    VariantInit(&varRetVal);
    hr = pacc->get_accState(*pvarChild, &varRetVal);
    if (!SUCCEEDED(hr))
        return(0);
    if (varRetVal.vt == VT_I4)
    {
        // 根据返回的状态值生成以逗号连接的字符串。
        GetStateText(varRetVal.lVal, lpszBuff, sizeBuff);
    }
    else if (varRetVal.vt == VT_BSTR)
    {
        if (varRetVal.bstrVal)
            _tcscpy_s(lpszBuff, sizeBuff, 
#ifndef _UNICODE
                _com_util::ConvertBSTRToString(varRetVal.bstrVal)
#else
                varRetVal.bstrVal
#endif
            );
    }
    VariantClear(&varRetVal);
    return _tcslen(lpszBuff);
}

UINT GetObjectName(IAccessible* pacc, VARIANT* pvarChild, LPTSTR lpszBuff, UINT sizeBuff)
{
    HRESULT hr;
    BSTR pszName;
    hr = pacc->get_accName(*pvarChild, &pszName);
    memset(lpszBuff, 0, sizeBuff * sizeof(lpszBuff[0]));
    if (pszName)_tcscpy_s(lpszBuff, sizeBuff, 
#ifndef _UNICODE
        _com_util::ConvertBSTRToString(pszName)
#else
        pszName
#endif
    );

    return _tcslen(lpszBuff);
}

UINT GetObjectValue(IAccessible* pacc, VARIANT* pvarChild, TCHAR* lpszBuff, UINT sizeBuff)
{
    HRESULT hr;
    BSTR pszValue;
    hr = pacc->get_accValue(*pvarChild, &pszValue);
    memset(lpszBuff, 0, sizeBuff * sizeof(lpszBuff[0]));
    if (pszValue)_tcscpy_s(lpszBuff, sizeBuff, 
#ifndef _UNICODE
        _com_util::ConvertBSTRToString(pszValue)
#else
        pszValue
#endif
    );

    //pacc->put_accValue(*pvarChild, CComBSTR("mensong"));

    return _tcslen(lpszBuff);
}

UINT GetObjectClass(IAccessible* pacc, LPTSTR lpszBuff, UINT sizeBuff)
{
    HWND hWnd;
    ::WindowFromAccessibleObject(pacc, &hWnd);
    memset(lpszBuff, 0, sizeBuff * sizeof(lpszBuff[0]));
    if (hWnd)::GetClassName(hWnd, lpszBuff, sizeBuff);
    return _tcslen(lpszBuff);
}

UINT GetObjectRole(IAccessible* pacc, VARIANT* pvarChild, LPTSTR lpszBuff, UINT sizeBuff)
{
    HRESULT hr;
    VARIANT varRetVal;
    //*lpszRole = 0;
    memset(lpszBuff, 0, sizeBuff * sizeof(lpszBuff[0]));
    VariantInit(&varRetVal);
    hr = pacc->get_accRole(*pvarChild, &varRetVal);
    if (!SUCCEEDED(hr))
        return(0);
    if (varRetVal.vt == VT_I4)
    {
        // 根据返回的状态值生成以逗号连接的字符串。
        GetRoleText(varRetVal.lVal, lpszBuff, sizeBuff);
    }
    else if (varRetVal.vt == VT_BSTR)
    {
        if (varRetVal.bstrVal)_tcscpy_s(lpszBuff, sizeBuff, 
#ifndef _UNICODE
            _com_util::ConvertBSTRToString(varRetVal.bstrVal)
#else
            varRetVal.bstrVal
#endif
        );
    }
    VariantClear(&varRetVal);
    return(_tcslen(lpszBuff));
}

UINT GetObjectDescription(IAccessible* pacc, VARIANT* pvarChild, LPTSTR lpszBuff, UINT buffSize)
{
	HRESULT hr;
	BSTR pszValue;
	hr = pacc->get_accDescription(*pvarChild, &pszValue);
	memset(lpszBuff, 0, buffSize * sizeof(lpszBuff[0]));
	if (pszValue)_tcscpy_s(lpszBuff, buffSize,
#ifndef _UNICODE
		_com_util::ConvertBSTRToString(pszValue)
#else
		pszValue
#endif
	);

	return _tcslen(lpszBuff);
}


void ReleaseAccessibles(std::vector<std::pair<IAccessible*, VARIANT>>& children)
{
    std::set<IAccessible*> hasReleased;
	for (size_t i = 0; i < children.size(); i++)
	{
        if (hasReleased.find(children[i].first) == hasReleased.end())
        {
			children[i].first->Release();
			hasReleased.insert(children[i].first);
        }
	}
	children.clear();
}

void PrintInfo(IAccessible* pAcc, VARIANT& child)
{
#define MAX_BUFF_LEN 10240
	TCHAR* szObjValue = new TCHAR[MAX_BUFF_LEN];
	TCHAR* szObjName = new TCHAR[MAX_BUFF_LEN];
	TCHAR* szObjRole = new TCHAR[MAX_BUFF_LEN];
	TCHAR* szObjClass = new TCHAR[MAX_BUFF_LEN];
	TCHAR* szObjState = new TCHAR[MAX_BUFF_LEN];
	//TCHAR* szObjDesc = new TCHAR[MAX_BUFF_LEN];

	// Skip invisible and unavailable objects and their children
	GetObjectState(pAcc, &child, szObjState, MAX_BUFF_LEN);
    if (_tcscmp(szObjState, _T("不可见")) != 0)
    {
		GetObjectName(pAcc, &child, szObjName, MAX_BUFF_LEN);
		GetObjectValue(pAcc, &child, szObjValue, MAX_BUFF_LEN);
		GetObjectRole(pAcc, &child, szObjRole, MAX_BUFF_LEN);
		//GetObjectRole(pAcc, &child, szObjDesc, MAX_BUFF_LEN);
		GetObjectClass(pAcc, szObjClass, MAX_BUFF_LEN);

		_tprintf(_T("szObjName=%s, szObjValue=%s, szObjState=%s, szObjClass=%s, szObjRole=%s\n"),
			szObjName, szObjValue, szObjState, szObjClass, szObjRole);
    }

	delete[] szObjValue;
	delete[] szObjName;
	delete[] szObjRole;
	delete[] szObjClass;
	delete[] szObjState;
	//delete[] szObjDesc;
}

void GetAccessibleChildren(
    IAccessible* pAccParent,    
    std::vector<std::pair<IAccessible*, VARIANT>>& children,
    bool oneLevel = true)
{
    HRESULT hr;
    long numChildren = 0;
    unsigned long numFetched;
    VARIANT varChild;
    int indexCount;
    IAccessible* pAcc = NULL;
    IEnumVARIANT* pEnum = NULL;
    IDispatch* pDisp = NULL;

    //Get the IEnumVARIANT interface
    hr = pAccParent->QueryInterface(IID_IEnumVARIANT, (PVOID*)&pEnum);
    if (pEnum)
		pEnum->Reset();

    // Get child count
    pAccParent->get_accChildCount(&numChildren);
	for (indexCount = 1; indexCount <= numChildren; indexCount++)
    {
		pAcc = NULL;
        numFetched = 0;

        // Get next child
        if (pEnum)
        {
            VariantInit(&varChild);
            hr = pEnum->Next(1, &varChild, &numFetched);
        }
        else
        {
            VariantInit(&varChild);
            varChild.vt = VT_I4;
            varChild.lVal = indexCount;
        }

        // Get IDispatch interface for the child
        if (varChild.vt == VT_I4)
        {
            pDisp = NULL;
            hr = pAccParent->get_accChild(varChild, &pDisp);
        }
        else
        {
            pDisp = varChild.pdispVal;
        }

        // Get IAccessible interface for the child
        if (pDisp)
        {
            hr = pDisp->QueryInterface(IID_IAccessible, (void**)&pAcc);
            hr = pDisp->Release();
        }

        // Get information about the child
        if (pAcc)
        {
            VariantInit(&varChild);
            varChild.vt = VT_I4;
            varChild.lVal = CHILDID_SELF;

        }
        else
        {
            pAcc = pAccParent;
        }
        
        //PrintInfo(pCAcc, varChild);
        children.push_back(std::make_pair(pAcc, varChild));

        if (!oneLevel && pAcc && pAcc != pAccParent)
        {
            // Go deeper
            GetAccessibleChildren(pAcc, children, oneLevel);
        }
    }

    // Clean up
    if (pEnum)
		pEnum->Release();
}

//
//  函数: GetWindowHWndByParentHWndAndClassName(HWND hParentHWnd, LPTSTR lpszClassName)
//  目的: 通过父句柄和类名称获取窗口句柄。
//  FindWindowEx
HWND GetWindowHWndByParentHWndAndClassName(HWND hParentHWnd, LPTSTR lpszClassName)
{
    HWND hWnd = NULL;
    HWND hTempParentHWnd = NULL;
    TCHAR szClassText[MAX_PATH] = { 0 };


    //hWnd = GetWindow(GetDesktopWindow(), GW_CHILD);
    hWnd = GetWindow(hParentHWnd, GW_CHILD);
    while (hWnd) {
        GetClassName(hWnd, szClassText, MAX_PATH);
        if (_tcsstr(szClassText, lpszClassName) != NULL) {

            hTempParentHWnd = GetParent(hWnd);

            if (hTempParentHWnd == hParentHWnd) {
                return hWnd;
            }
        }
        hWnd = GetNextWindow(hWnd, GW_HWNDNEXT);
    }
    return NULL;
}

int main()
{
    system("title WOW&color 2e");

    ::CoInitialize(NULL);

    //TCHAR _szClassName[MAX_PATH] = { NULL };    
    // ShellExecuteA(NULL, "open", "rundll32.exe", "shell32.dll,Control_RunDLL ncpa.cpl", NULL, SW_NORMAL);
    //HWND hWnd = ::FindWindow("CabinetWClass", "网络连接");
    //hWnd = ::FindWindowEx(hWnd, NULL, "ShellTabWindowClass", "网络连接");
    //hWnd = ::FindWindowEx(hWnd, NULL, "DUIViewWndClassName", NULL);
    //hWnd = ::FindWindowEx(hWnd, NULL, "DirectUIHWND", NULL);
    //hWnd = ::FindWindowEx(hWnd, NULL, "DirectUIHWND", NULL);
    //hWnd = ::FindWindowEx(hWnd, NULL, "CtrlNotifySink", NULL);
    //
    //hWnd = GetNextWindow(hWnd, GW_HWNDNEXT);
    //GetClassName(hWnd,_szClassName, MAX_PATH);
    //hWnd = GetNextWindow(hWnd, GW_HWNDNEXT);
    //    GetClassName(hWnd,_szClassName, MAX_PATH);
   //    hWnd = GetNextWindow(hWnd, GW_HWNDNEXT);
    //    GetClassName(hWnd,_szClassName, MAX_PATH);
   //    hWnd = GetNextWindow(hWnd, GW_HWNDNEXT);
    //    GetClassName(hWnd,_szClassName, MAX_PATH);
    //hWnd = ::FindWindowEx(hWnd, NULL, "SHELLDLL_DefView", "ShellView");
    // hWnd = ::FindWindowEx(hWnd, NULL, "DirectUIHWND", NULL);

    /*HWND hWnd = ::FindWindow("photoshop", NULL);
    hWnd = ::FindWindowEx(hWnd, NULL, "OWL.ApplicationBar", NULL);
    hWnd = ::FindWindowEx(hWnd, NULL, "OWL.MenuBar", NULL);*/

	//HWND hWnd = FindWindow(_T("ThunderRT6FormDC"), _T("Paste++"));
	//HWND hWnd = FindWindow(_T("#32770"), _T("Export File"));
    HWND hWnd = FindWindow(_T("WeChatMainWndForPC"), _T("微信"));
    if (!::IsWindow(hWnd))
        return FALSE;

    IAccessible* pIAcc = NULL;
    HRESULT hr = AccessibleObjectFromWindow(hWnd, OBJID_WINDOW, IID_IAccessible, (void**)&pIAcc);
    if (SUCCEEDED(hr) && pIAcc)
    {
        //if (FindAccessible(pIAcc, "本地连接", "列表项目","DirectUIHWND", &pIAccChild, &varChild))
        std::vector<std::pair<IAccessible*, VARIANT>> children;
        GetAccessibleChildren(pIAcc, children, false);

		//pIAccChild->accSelect(SELFLAG_TAKESELECTION | SELFLAG_TAKEFOCUS, varChild);
		//Sleep(2000);
		//pIAccChild->accDoDefaultAction(varChild);

		for (size_t i = 0; i < children.size(); i++)
		{
            PrintInfo(children[i].first, children[i].second);
        }

		//IAccessible* pAccChild = children[children.size() - 11].first;
		//VARIANT varChild = children[children.size() - 11].second;

		////VariantInit(&varChild);
		////varChild.vt = VT_I4;
		////varChild.lVal = CHILDID_SELF;

		//pAccChild->accSelect(
		//	SELFLAG_TAKESELECTION | SELFLAG_TAKEFOCUS,
		//	varChild);
		//pAccChild->accDoDefaultAction(varChild);

		ReleaseAccessibles(children);
    }
    ::CoUninitialize();
    system("pause");
    return 0;
}
