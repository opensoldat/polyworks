// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the PWLIB_EXPORTS
// symbol defined on the command line. this symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// PWLIB_API functions as being imported from a DLL, whereas this DLL sees symbols
// defined with this macro as being exported.
#ifdef PWLIB_EXPORTS
#define PWLIB_API __declspec(dllexport)
#else
#define PWLIB_API extern "C" __declspec(dllimport)
#endif


// Exported functions.
PWLIB_API int _stdcall PwGifToBmp(LPCSTR src, LPCSTR dest);

// Helper functions.
int SaveTGA(Bitmap* bitmap, LPCSTR fileName);