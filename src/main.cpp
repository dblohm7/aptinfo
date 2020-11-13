#include <string>

#include <stdio.h>

#include <windows.h>
#include <objbase.h>
#include <locationapi.h>
#include <mmdeviceapi.h>

#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "locationapi.lib")
#pragma comment(lib, "mmdevapi.lib")

#define GET_CLSID_STR_3(x) #x
#define GET_CLSID_STR_2(x) GET_CLSID_STR_3(x)
#define GET_CLSID_STR GET_CLSID_STR_2(CHECK_CLSID)

#define GET_CLSID_2(x) x
#define GET_CLSID GET_CLSID_2(CHECK_CLSID)

static const int kGuidLenWithBracesNul = 39;
static const size_t kThreadingModelBufChars = 10; // ArrayLength(L"Apartment")

int wmain(int argc, wchar_t* argv[]) {
  wchar_t guidBuf[kGuidLenWithBracesNul];
  if (::StringFromGUID2(GET_CLSID, guidBuf, kGuidLenWithBracesNul) != kGuidLenWithBracesNul) {
    wprintf(L"Failed to convert \"%S\" to a string\n", GET_CLSID_STR);
    return 1;
  }

  wprintf(L"CHECK_CLSID is %S: %s\n", GET_CLSID_STR, guidBuf);

  std::wstring subKey(L"CLSID\\");
  subKey += guidBuf;
  subKey += L"\\InprocServer32";

  wchar_t threadingModelBuf[kThreadingModelBufChars] = {};

  DWORD numBytes = sizeof(threadingModelBuf);
  LONG result = ::RegGetValueW(HKEY_CLASSES_ROOT, subKey.c_str(),
                               L"ThreadingModel", RRF_RT_REG_SZ, nullptr,
                               threadingModelBuf, &numBytes);
  if (result != ERROR_SUCCESS) {
    wprintf(L"Query failed with code %ld. Not an inproc server?\n", result);
    return 1;
  }

  wprintf(L"Threading model is \"%s\"\n", threadingModelBuf);
  return 0;
}

