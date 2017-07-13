#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#include <windows.h>
#include <shobjidl.h>
#include <shlguid.h>

TCHAR szClassName[] = TEXT("Window");

BOOL GetTargetFile(LPCWSTR lpszLinkFilePath, LPWSTR lpszTargetFilePath)
{
	BOOL bReturn = FALSE;
	IShellLink *pShellLink = 0;
	CoInitialize(NULL);
	if (S_OK == CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (void **)&pShellLink))
	{
		IPersistFile *pPersistFile = 0;
		if (S_OK == pShellLink->QueryInterface(IID_IPersistFile, (void **)&pPersistFile))
		{
			if (S_OK == pPersistFile->Load(lpszLinkFilePath, STGM_READ))
			{
				if (S_OK == pShellLink->GetPath(lpszTargetFilePath, MAX_PATH, 0, SLGP_UNCPRIORITY))
				{
					bReturn = TRUE;
				}
			}
			pPersistFile->Release();
		}
		pShellLink->Release();
	}
	CoUninitialize();
	return bReturn;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
	{
	case WM_CREATE:
		DragAcceptFiles(hWnd, TRUE);
		break;
	case WM_DROPFILES:
		{
			const UINT iFileNum = DragQueryFile((HDROP)wParam, -1, NULL, 0);
			for (UINT i = 0; i < iFileNum; ++i)
			{
				TCHAR szFilePath[MAX_PATH];
				DragQueryFile((HDROP)wParam, i, szFilePath, MAX_PATH);
				TCHAR szTargetFilePath[MAX_PATH];
				if (GetTargetFile(szFilePath, szTargetFilePath))
				{
					lstrcpy(szFilePath, szTargetFilePath);
				}
				MessageBox(hWnd, szFilePath, TEXT("確認"), 0);
			}
			DragFinish((HDROP)wParam);
		}
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		(HBRUSH)(COLOR_WINDOW + 1),
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("ドラッグ＆ドロップされたショートカット先のファイルのパスを取得"),
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return (int)msg.wParam;
}
