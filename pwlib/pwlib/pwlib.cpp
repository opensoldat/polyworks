// pwlib.cpp : Defines the exported functions for the DLL application.
//

#include "stdafx.h"
#include "pwlib.h"


// Converts a gif image to a bmp so it can be used in PW.
// --------------------------------
// var src: File name of gif image to convert.
// var dest: File name of the new bmp image.
// return: -1 if successfull, 0-8 on error.
PWLIB_API int _stdcall PwGifToBmp(LPCSTR src, LPCSTR dest) {
	bool error;
	int wideLength, result = -1;
	LPWSTR wideSrc;
	ULONG_PTR gdiToken;
	GdiplusStartupInput gdiStartupInput;
	Bitmap* bitmap;

	// Convert source file name to wide char format.
	wideLength = 2*MultiByteToWideChar(CP_ACP, 0, src, -1, 0, 0);
	if(!wideLength)
		return 0;
	wideSrc = (LPWSTR)HeapAlloc(GetProcessHeap(), 0, wideLength);
	if(wideSrc)
		error = !MultiByteToWideChar(CP_ACP, 0, src, -1, wideSrc, wideLength);
	else
		return 1;
	if(error) {
		result = 2;
		goto GifToBmpEnd0;
	}

	// Startup GDI+.
	error = GdiplusStartup(&gdiToken, &gdiStartupInput, 0) != Ok;
	if(error) {
		result = 3;
		goto GifToBmpEnd0;
	}

	// Convert the image.
	bitmap = new Bitmap(wideSrc);
	if(!bitmap) {
		result = 4;
		goto GifToBmpEnd1;
	}
	result = SaveTGA(bitmap, dest);
	if(result != -1)
		result += 5;

	// Clean up.
	delete bitmap;
GifToBmpEnd1:
	GdiplusShutdown(gdiToken);
GifToBmpEnd0:
	HeapFree(GetProcessHeap(), 0, wideSrc);

	// Return.
	return result;
}


/*
 *      _Helper functions_
 */


int SaveTGA(Bitmap* bitmap, LPCSTR fileName) {
	bool error;
	int result = -1;
	HANDLE tgaFile;
	Rect bmpRect;
	BitmapData bmpData;
	DWORD dwNotUsed;
	UINT i, j;
	LPBYTE offData;

	// Lock the bitmap bits for reading.
	bmpRect.X = 0;
	bmpRect.Y = 0;
	bmpRect.Width = bitmap->GetWidth();
	bmpRect.Height = bitmap->GetHeight();
	error = bitmap->LockBits(&bmpRect, ImageLockModeRead, PixelFormat32bppARGB, &bmpData) != Ok;
	if(error)
		return 0;

	// TGA header.
	BYTE tgaHeaderTop[12] = {0, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0};
	WORD tgaHeaderWidth = bmpData.Width;
	WORD tgaHeaderHeight = bmpData.Height;
	WORD tgaHeaderBpp = 32;
	BYTE tgaHeaderDesc = 8;

	// Create the file.
	tgaFile = CreateFileA(fileName, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0);
	if(tgaFile == INVALID_HANDLE_VALUE) {
		result = 1;
		goto SaveTGAEnd0;
	}

	// Write the header to the file.
	error = !WriteFile(tgaFile, (LPCVOID)&tgaHeaderTop, 12, &dwNotUsed, 0);
	error |= !WriteFile(tgaFile, (LPCVOID)&tgaHeaderWidth, 2, &dwNotUsed, 0);
	error |= !WriteFile(tgaFile, (LPCVOID)&tgaHeaderHeight, 2, &dwNotUsed, 0);
	error |= !WriteFile(tgaFile, (LPCVOID)&tgaHeaderBpp, 2, &dwNotUsed, 0);
	error |= !WriteFile(tgaFile, (LPCVOID)&tgaHeaderDesc, 4, &dwNotUsed, 0);
	if(error) {
		result = 2;
		goto SaveTGAEnd1;
	}

	// Write the bitmap data to the file.
	error = 0;
	for(i = 0; i < bmpData.Height; ++i) {
		for(j = 0; j < bmpData.Width; ++j) {
			offData = (LPBYTE)bmpData.Scan0;
			offData += 4*((bmpData.Height - i - 1)*bmpData.Width + j);
			error |= !WriteFile(tgaFile, (LPCVOID)offData, 4, &dwNotUsed, 0);
		}
	}
	if(error)
		result = 3;

	// Clean up.
SaveTGAEnd1:
	CloseHandle(tgaFile);
SaveTGAEnd0:
	bitmap->UnlockBits(&bmpData);

	// Return.
	return result;
}