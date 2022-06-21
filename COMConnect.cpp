// COMConnect.cpp : Этот файл содержит функцию "main". Здесь начинается и заканчивается выполнение программы.
//


// https://www.rsdn.org/article/com/introcom.xml

// https://docs.microsoft.com/ru-ru/cpp/text/how-to-convert-between-various-string-types?view=msvc-170

#include <combaseapi.h>
#include <stdio.h>
#include <comdef.h>
#include <string>
#include <iostream>

// AutoWrap() - Automation helper function...
HRESULT AutoWrap(int autoType, VARIANT* pvResult, IDispatch* pDisp, LPOLESTR ptName, int cArgs...) {
    // Begin variable-argument list...
    va_list marker;
    va_start(marker, cArgs);

    if (!pDisp) {
        MessageBox(NULL, L"NULL IDispatch passed to AutoWrap()", L"Error", 0x10010);
        _exit(0);
    }
    
    // Variables used...
    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    DISPID dispID;
    HRESULT hr;
    wchar_t buf[200];
    //char szName[200];


    // Convert down to ANSI
    //WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);
    
    // Get DISPID for name passed...
    hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_SYSTEM_DEFAULT, &dispID);
    if (FAILED(hr)) {
        swprintf(buf, 200, L"IDispatch::GetIDsOfNames(\"%s\") failed w/err 0x%08lx", ptName, hr);
        MessageBox(NULL, buf, L"AutoWrap()", 0x10010);
        _exit(0);
        return hr;
    }

    // Allocate memory for arguments...
    VARIANT* pArgs = new VARIANT[cArgs + 1];
    // Extract arguments...
    for (int i = 0; i < cArgs; i++) {
        pArgs[i] = va_arg(marker, VARIANT);
    }

    // Build DISPPARAMS
    dp.cArgs = cArgs;
    dp.rgvarg = pArgs;

    // Handle special-case for property-puts!
    if (autoType & DISPATCH_PROPERTYPUT) {
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
    }

    // Make the call!
    hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);
    if (FAILED(hr)) {
        swprintf(buf, 200, L"IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx", pArgs[0].bstrVal, dispID, hr);
        MessageBox(NULL, buf, L"AutoWrap()", 0x10010);
        _exit(0);
        return hr;
    }
    // End variable-argument section...
    va_end(marker);

    delete[] pArgs;

    return hr;
}

int main()
{

    // Initialize COM for this thread...
    CoInitialize(NULL);

    // Get CLSID for our server...
    CLSID clsid;
    //HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
    HRESULT hr = CLSIDFromProgID(L"V83.COMConnector", &clsid);

    if (FAILED(hr)) {

        ::MessageBox(NULL, L"CLSIDFromProgID() failed", L"Error", 0x10010);
        return -1;
    }

    // Start server and get IDispatch...
    IDispatch* pApp;
    hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (void**)&pApp);
    if (FAILED(hr)) {
        ::MessageBox(NULL, L"Application not registered properly", L"Error", 0x10010);
        return -2;
    }
 
    IDispatch* pConnect;
    {
        VARIANT parm;
        parm.vt = VT_BSTR;
        const wchar_t* PC = L"File=\"C:\\1С\\DemoEnterprise20\";Usr=\"Администратор (ОрловАВ)\"";
        _bstr_t bstrtpc(PC);
        parm.bstrVal = bstrtpc;

        VARIANT result;
        VariantInit(&result);

        AutoWrap(DISPATCH_METHOD, &result, pApp, (LPOLESTR)L"Connect", 1, parm);
        pConnect = result.pdispVal;
    }

    IDispatch* pCatalogs;
    {
        VARIANT result;
        VariantInit(&result);

        AutoWrap(DISPATCH_PROPERTYGET, &result, pConnect, (LPOLESTR)L"Справочники", 0);
        pCatalogs = result.pdispVal;
    }

    IDispatch* pCatalogsCurrency;
    {
        VARIANT result;
        VariantInit(&result);

        AutoWrap(DISPATCH_PROPERTYGET, &result, pCatalogs, (LPOLESTR)L"Валюты", 0);
        pCatalogsCurrency = result.pdispVal;
    }

    IDispatch* pCurrency;
    {
        VARIANT result;
        VariantInit(&result);

        AutoWrap(DISPATCH_METHOD, &result, pCatalogsCurrency, (LPOLESTR)L"СоздатьЭлемент", 0);
        pCurrency = result.pdispVal;
    }

    {
        VARIANT parm;
        parm.vt = VT_BSTR;
        parm.bstrVal = _bstr_t(L"Тестовая");
        AutoWrap(DISPATCH_PROPERTYPUT, NULL, pCurrency, (LPOLESTR)L"Наименование", 1, parm);
    }

    AutoWrap(DISPATCH_METHOD, NULL, pCurrency, (LPOLESTR)L"Записать", 0);

    //// Make it visible (i.e. app.visible = 1)
    //{

    //    VARIANT x;
    //    x.vt = VT_I4;
    //    x.lVal = 1;
    //    AutoWrap(DISPATCH_PROPERTYPUT, NULL, pApp, (LPOLESTR)L"Visible", 1, x);
    //}

    //// Get Workbooks collection
    //IDispatch* pXlBooks;
    //{
    //    VARIANT result;
    //    VariantInit(&result);
    //    AutoWrap(DISPATCH_PROPERTYGET, &result, pApp, (LPOLESTR)L"Workbooks", 0);
    //    pXlBooks = result.pdispVal;
    //}

    //// Call Workbooks.Add() to get a new workbook...
    //IDispatch* pXlBook;
    //{
    //    VARIANT result;
    //    VariantInit(&result);
    //    AutoWrap(DISPATCH_PROPERTYGET, &result, pXlBooks, (LPOLESTR)L"Add", 0);
    //    pXlBook = result.pdispVal;
    //}

    //// Create a 15x15 safearray of variants...
    //VARIANT arr;
    //arr.vt = VT_ARRAY | VT_VARIANT;
    //{
    //    SAFEARRAYBOUND sab[2];
    //    sab[0].lLbound = 1; sab[0].cElements = 15;
    //    sab[1].lLbound = 1; sab[1].cElements = 15;
    //    arr.parray = SafeArrayCreate(VT_VARIANT, 2, sab);
    //}

    //// Fill safearray with some values...
    //for (int i = 1; i <= 15; i++) {
    //    for (int j = 1; j <= 15; j++) {
    //        // Create entry value for (i,j)
    //        VARIANT tmp;
    //        tmp.vt = VT_I4;
    //        tmp.lVal = i * j;
    //        // Add to safearray...
    //        long indices[] = { i,j };
    //        SafeArrayPutElement(arr.parray, indices, (void*)&tmp);
    //    }
    //}

    //// Get ActiveSheet object
    //IDispatch* pXlSheet;
    //{
    //    VARIANT result;
    //    VariantInit(&result);
    //    AutoWrap(DISPATCH_PROPERTYGET, &result, pApp, (LPOLESTR)L"ActiveSheet", 0);
    //    pXlSheet = result.pdispVal;
    //}

    //// Get Range object for the Range A1:O15...
    //IDispatch* pXlRange;
    //{
    //    VARIANT parm;
    //    parm.vt = VT_BSTR;
    //    parm.bstrVal = ::SysAllocString(L"A1:O15");

    //    VARIANT result;
    //    VariantInit(&result);
    //    AutoWrap(DISPATCH_PROPERTYGET, &result, pXlSheet, (LPOLESTR)L"Range", 1, parm);
    //    VariantClear(&parm);

    //    pXlRange = result.pdispVal;
    //}

    //// Set range with our safearray...
    //AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlRange, (LPOLESTR)L"Value", 1, arr);

    //// Wait for user...
    //::MessageBox(NULL, L"All done.", L"Notice", 0x10000);

    //// Set .Saved property of workbook to TRUE so we aren't prompted
    //// to save when we tell Excel to quit...
    //{
    //    VARIANT x;
    //    x.vt = VT_I4;
    //    x.lVal = 1;
    //    AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlBook, (LPOLESTR)L"Saved", 1, x);
    //}

    //// Tell Excel to quit (i.e. App.Quit)
    //AutoWrap(DISPATCH_METHOD, NULL, pApp, (LPOLESTR)L"Quit", 0);

    //// Release references...
    //pXlRange->Release();
    //pXlSheet->Release();
    //pXlBook->Release();
    //pXlBooks->Release();
    pApp->Release();
    //VariantClear(&arr);

    // Uninitialize COM for this thread...
    CoUninitialize();
}

