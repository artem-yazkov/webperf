#include <atlbase.h>
#include <atlcom.h>
#include <atlconv.h>

#include <mshtml.h>
#include <exdisp.h>
#include <exdispid.h>

#include <windows.h>
#include <cstdio>
#include <conio.h>

#pragma comment(lib, "Winmm.lib")

/*
 *
 */
class IENavigator :
	public IDispEventImpl<1, IENavigator, &DIID_DWebBrowserEvents2, &LIBID_SHDocVw, 1, 1>
{
public:
	BEGIN_SINK_MAP(IENavigator)
		SINK_ENTRY_EX(1, DIID_DWebBrowserEvents2, DISPID_DOCUMENTCOMPLETE, OnDocumentComplete)
	END_SINK_MAP()
public:
	IENavigator();
    ~IENavigator();
    
    void Navigate(const char *url);
private:
	IWebBrowser2* pBrowser;
	void __stdcall OnDocumentComplete(IDispatch* pDisp, VARIANT* URL);
    void WaitForComplete();
    void WaitForCompleteSink();
};

IENavigator::IENavigator()
{
    CoInitialize(NULL);
	HRESULT hr = CoCreateInstance(CLSID_InternetExplorer, NULL, CLSCTX_LOCAL_SERVER,
                                  IID_IWebBrowser2, (void**)&pBrowser);
    //hr = DispEventAdvise(pBrowser);
}

IENavigator::~IENavigator()
{
    //DispEventUnadvise(pBrowser);
    pBrowser->Release();
    CoUninitialize();
}

void
IENavigator::OnDocumentComplete(IDispatch* pDisp, VARIANT* URL)
{
    printf("OnDocumentComplete\n");
}

void
IENavigator::WaitForComplete()
{
    CHandle event(CreateEvent(NULL, TRUE, FALSE, NULL));
    while (true) {
        /* Wait some messages for one millisecond */
        const DWORD nWaitResult = MsgWaitForMultipleObjects(1, &event.m_h, FALSE, 1, QS_ALLINPUT | QS_ALLPOSTMESSAGE);
        if (nWaitResult != WAIT_TIMEOUT) {
            /* Dispatch all messages from queue */
            MSG Message;
            while (PeekMessage(&Message, NULL, WM_NULL, WM_NULL, PM_REMOVE)) {
                TranslateMessage(&Message);
                DispatchMessage(&Message);
            }
        }
        READYSTATE rs;
        if (SUCCEEDED(pBrowser->get_ReadyState(&rs)) && (rs == READYSTATE_COMPLETE)) {
            return;
        }
    }
}

void
IENavigator::WaitForCompleteSink()
{
    CHandle event(CreateEvent(NULL, TRUE, FALSE, NULL));
    while (true) {
        const DWORD nWaitResult = MsgWaitForMultipleObjects(1, &event.m_h, FALSE, 10000, QS_ALLINPUT | QS_ALLPOSTMESSAGE);
        if (nWaitResult == WAIT_TIMEOUT) {
            break;
        }
        MSG Message;
        while (PeekMessage(&Message, NULL, WM_NULL, WM_NULL, PM_REMOVE)) {
            TranslateMessage(&Message);
            DispatchMessage(&Message);
        }
        READYSTATE rs;
        HRESULT    hr = pBrowser->get_ReadyState(&rs);
        if (SUCCEEDED(hr) && (rs == READYSTATE_COMPLETE)) {
            return;
        }
    }
}

void
IENavigator::Navigate(const char *url)
{
    VARIANT vEmpty;
    VariantInit(&vEmpty);

    BSTR bstrURL = SysAllocString(CA2W(url));
    
    DWORD startTime = timeGetTime();
    HRESULT hr = pBrowser->Navigate(bstrURL, &vEmpty, &vEmpty, &vEmpty, &vEmpty);

    if (SUCCEEDED(hr)) {
        WaitForComplete();
        DWORD stopTime = timeGetTime();
        printf("%s: %u ms\n", url, stopTime - startTime);
        pBrowser->put_Visible(VARIANT_TRUE);
    } else {
        pBrowser->Quit();
    }
    SysFreeString(bstrURL);
}

int main()
{
    IENavigator navigator;
    navigator.Navigate("http://microsoft.com");
    printf("ready\n");
    _getch();

    return 0;
}