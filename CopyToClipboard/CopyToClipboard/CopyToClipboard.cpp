#include<windows.h>

int WINAPI WinMain(
	HINSTANCE hInstance,
	HINSTANCE hPrevInstance,
	char *lpCmdLine,
	int nCmdShow)
{
	if (!lpCmdLine) return 1;
	auto length = strlen(lpCmdLine);
	if (!length) return 2;
	auto _pBuf = GlobalAlloc(GMEM_SHARE | GMEM_MOVEABLE, (length + 1) * sizeof(char));
	if (!_pBuf) return 3;
	auto pBuf = (TCHAR *)GlobalLock(_pBuf);
	if (!pBuf)
	{
		GlobalFree(_pBuf);
		return 4;
	}

	strcpy(pBuf, lpCmdLine);

	GlobalUnlock(pBuf);
	if (!OpenClipboard(NULL)) return 3;
	EmptyClipboard();
	SetClipboardData(CF_TEXT, pBuf);
	CloseClipboard();

	return 0;
}
