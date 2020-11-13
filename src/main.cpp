#include <string>

#include <stdio.h>

#include <windows.h>
#include <objbase.h>

#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "ole32.lib")


template <typename T, size_t N>
static inline constexpr size_t ArrayLength(T (&aArr)[N])
{
  return N;
}

static constexpr int kGuidLenWithBracesInclNul = 39;
static constexpr int kGuidLenWithBracesExclNul = kGuidLenWithBracesInclNul - 1;
static constexpr size_t kThreadingModelBufChars = ArrayLength(L"Apartment");

static wchar_t gStrClsid[kGuidLenWithBracesInclNul];
static wchar_t gStrIid[kGuidLenWithBracesInclNul];
static CLSID gClsid;
static IID gIid;
static bool gHasIid = false;

static void
Usage(const wchar_t* aArgv0, const wchar_t* aMsg = nullptr)
{
  if (aMsg) {
    fwprintf_s(stderr, L"Error: %s\n\n", aMsg);
  }

  fwprintf_s(stderr, L"Usage: %s <ProgID or CLSID> [IID]\n\n", aArgv0);
  fwprintf_s(stderr, L"    CLSID and IID must be specified in registry format,\n    including dashes and curly braces\n");
}

static bool
ParseArgv(const int argc, wchar_t* argv[])
{
  if (argc < 2) {
    Usage(argv[0]);
    return false;
  }

  // First, get the CLSID.
  if (argv[1][0] == L'{' && wcslen(argv[1]) == kGuidLenWithBracesExclNul) {
    // Assume it's a guid.
    if (wcscpy_s(gStrClsid, argv[1])) {
      return false;
    }

    if (FAILED(::CLSIDFromString(gStrClsid, &gClsid))) {
      Usage(argv[0], L"Failed to parse CLSID");
      return false;
    }

  } else {
    // Assume it's a ProgID
    if (FAILED(::CLSIDFromProgID(argv[1], &gClsid))) {
      Usage(argv[0], L"Invalid ProgID");
      return false;
    }

    wprintf_s(L"Found ProgID \"%s\"\n", argv[1]);

    if (!::StringFromGUID2(gClsid, gStrClsid, ArrayLength(gStrClsid))) {
      return false;
    }
  }

  wprintf_s(L"Using CLSID %s\n", gStrClsid);

  if (argc == 2) {
    // No optional IID provided, continue
    return true;
  }

  // IID sanity checks
  if (wcslen(argv[2]) != kGuidLenWithBracesExclNul || argv[2][0] != L'{') {
    Usage(argv[0], L"Invalid IID");
    return false;
  }

  // Assume it's a guid.
  if (wcscpy_s(gStrIid, argv[2])) {
    return false;
  }

  if (FAILED(::IIDFromString(gStrIid, &gIid))) {
    Usage(argv[0], L"Failed to parse IID");
    return false;
  }

  wprintf_s(L"Using IID %s\n", gStrIid);

  gHasIid = true;
  return true;
}

int
wmain(int argc, wchar_t* argv[]) {
  if (!ParseArgv(argc, argv)) {
    return 1;
  }

  std::wstring subKey(L"CLSID\\");
  subKey += gStrClsid;
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

