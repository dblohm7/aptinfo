#include <string>
#include <variant>

#include <stdio.h>

#include <windows.h>
#include <objbase.h>

#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "ole32.lib")

template <typename T, size_t N>
static inline constexpr size_t ArrayLength(T (&aArr)[N]) {
  return N;
}

static constexpr int kGuidLenWithBracesInclNul = 39;
static constexpr int kGuidLenWithBracesExclNul = kGuidLenWithBracesInclNul - 1;
static constexpr size_t kThreadingModelBufChars = ArrayLength(L"Apartment");
static const CLSID CLSID_FreeThreadedMarshaler = {0x0000033A,
                                                  0x0000,
                                                  0x0000,
                                                  {
                                                      0xC0,
                                                      0x00,
                                                      0x00,
                                                      0x00,
                                                      0x00,
                                                      0x00,
                                                      0x00,
                                                      0x46,
                                                  }};

static wchar_t gStrClsid[kGuidLenWithBracesInclNul];
static wchar_t gStrIid[kGuidLenWithBracesInclNul];
static CLSID gClsid;
static IID gIid;
static bool gHasIid = false;

static void Usage(const wchar_t* aArgv0, const wchar_t* aMsg = nullptr) {
  if (aMsg) {
    fwprintf_s(stderr, L"Error: %s\n\n", aMsg);
  }

  fwprintf_s(stderr, L"Usage: %s <ProgID or CLSID> [IID]\n\n", aArgv0);
  fwprintf_s(stderr,
             L"    CLSID and IID must be specified in registry format,\n    "
             L"including dashes and curly braces\n");
}

static bool ParseArgv(const int argc, wchar_t* argv[]) {
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

enum class ThreadingModel { STA, MTA, Both, Neutral };

static std::variant<ThreadingModel, LSTATUS> GetClassThreadingModel(
    const wchar_t* aStrClsid) {
  std::wstring subKeyInprocServer(L"CLSID\\");
  subKeyInprocServer += aStrClsid;
  subKeyInprocServer += L"\\InprocServer32";

  wchar_t threadingModelBuf[kThreadingModelBufChars] = {};
  DWORD numBytes = sizeof(threadingModelBuf);
  LSTATUS result = ::RegGetValueW(HKEY_CLASSES_ROOT, subKeyInprocServer.c_str(),
                                  L"ThreadingModel", RRF_RT_REG_SZ, nullptr,
                                  threadingModelBuf, &numBytes);
  if (result != ERROR_SUCCESS) {
    return result;
  }

  if (!_wcsicmp(threadingModelBuf, L"Apartment")) {
    return ThreadingModel::STA;
  }

  if (!_wcsicmp(threadingModelBuf, L"Free")) {
    return ThreadingModel::MTA;
  }

  if (!_wcsicmp(threadingModelBuf, L"Both")) {
    return ThreadingModel::Both;
  }

  if (!_wcsicmp(threadingModelBuf, L"Neutral")) {
    return ThreadingModel::Neutral;
  }

  return static_cast<LSTATUS>(ERROR_UNIDENTIFIED_ERROR);
}

static const wchar_t* GetThreadingModelDescription(
    const ThreadingModel aThdModel) {
  switch (aThdModel) {
    case ThreadingModel::STA:
      return L"single-threaded:\nProxying is required to access from any other "
             L"apartment.";
    case ThreadingModel::MTA:
      return L"multi-threaded:\nProxying is required to access from any "
             L"single-threaded apartment.";
    case ThreadingModel::Both:
      return L"either single-threaded or multi-threaded, but not "
             L"simultaneously";
    case ThreadingModel::Neutral:
      return L"thread-neutral:\nThis class may be invoked by any thread within "
             L"any apartment.";
    default:
      return nullptr;
  };
}

class Apartment {
 public:
  explicit Apartment(const ThreadingModel aThdModel)
      : mHr(::CoInitializeEx(nullptr, (aThdModel == ThreadingModel::STA)
                                          ? COINIT_APARTMENTTHREADED
                                          : COINIT_MULTITHREADED)) {}

  ~Apartment() {
    if (FAILED(mHr)) {
      return;
    }

    ::CoUninitialize();
  }

  explicit operator bool() const { return SUCCEEDED(mHr); }

  Apartment(const Apartment&) = delete;
  Apartment(Apartment&&) = delete;
  Apartment& operator=(const Apartment&) = delete;
  Apartment& operator=(Apartment&&) = delete;

 private:
  const HRESULT mHr;
};

static ThreadingModel CheckServerObjectCapabilities(
    REFCLSID aClsid, const ThreadingModel aThdModel,
    const IID* aIid = nullptr) {
  if (aThdModel == ThreadingModel::Neutral) {
    // We're already neutral, these additional checks are unnecessary.
    return aThdModel;
  }

  Apartment apt(aThdModel);
  if (!apt) {
    return aThdModel;
  }

  IUnknown* punk = nullptr;
  HRESULT hr =
      ::CoCreateInstance(aClsid, nullptr, CLSCTX_INPROC_SERVER, IID_IUnknown,
                         reinterpret_cast<void**>(&punk));
  if (FAILED(hr)) {
    return aThdModel;
  }

  wprintf_s(L"Querying for IAgileObject...");
  IUnknown* agile = nullptr;
  hr = punk->QueryInterface(IID_IAgileObject, reinterpret_cast<void**>(&agile));
  if (SUCCEEDED(hr)) {
    wprintf_s(L" found.\n");
    agile->Release();
    punk->Release();
    return ThreadingModel::Neutral;
  }

  wprintf_s(L" not found.\n");

  if (!aIid) {
    wprintf_s(L"Warning: IID required to query for free-threaded marshaler.\n");
    // We need an IID to do any further checks
    return aThdModel;
  }

  wprintf_s(L"Querying for IMarshal...");

  // Check for the free-threaded marshaler
  IMarshal* marshal = nullptr;
  hr = punk->QueryInterface(IID_IMarshal, reinterpret_cast<void**>(&marshal));
  if (FAILED(hr)) {
    wprintf_s(L" not found.\n");
    punk->Release();
    return aThdModel;
  }

  wprintf_s(L" found.\nChecking whether object uses the free-threaded marshaler...");

  CLSID unmarshalClass;
  hr = marshal->GetUnmarshalClass(*aIid, nullptr, MSHCTX_INPROC, nullptr,
                                  MSHLFLAGS_NORMAL, &unmarshalClass);

  // We're done with these. Release unconditionally.
  marshal->Release();
  punk->Release();

  if (FAILED(hr)) {
    wprintf_s(L" failed.\n");
    return aThdModel;
  }

  if (unmarshalClass == CLSID_FreeThreadedMarshaler) {
    wprintf_s(L" yes.\n");
    return ThreadingModel::Neutral;
  }

  wprintf_s(L" no.\n");
  return aThdModel;
}

int wmain(int argc, wchar_t* argv[]) {
  if (!ParseArgv(argc, argv)) {
    return 1;
  }

  std::variant<ThreadingModel, LSTATUS> inprocModel =
      GetClassThreadingModel(gStrClsid);
  if (std::holds_alternative<ThreadingModel>(inprocModel)) {
    wprintf_s(L"Server threading model is %s.\n",
              GetThreadingModelDescription(CheckServerObjectCapabilities(
                  gClsid, std::get<ThreadingModel>(inprocModel),
                  gHasIid ? &gIid : nullptr)));
    return 0;
  }

  LSTATUS result = std::get<LSTATUS>(inprocModel);
  if (result == ERROR_FILE_NOT_FOUND) {
    wprintf_s(L"Class is not a registered in-process server.\n");
  } else {
    fwprintf_s(stderr, L"InprocServer32 query failed with code %ld.\n", result);
    return 1;
  }

  wprintf_s(L"Attempting to resolve as a local server...\n");

  std::wstring subKeyLocalServer(L"CLSID\\");
  subKeyLocalServer += gStrClsid;
  subKeyLocalServer += L"\\LocalServer32";

  HKEY regKeyLocalServer;
  result = ::RegOpenKeyExW(HKEY_CLASSES_ROOT, subKeyLocalServer.c_str(), 0,
                           KEY_READ, &regKeyLocalServer);
  if (result == ERROR_FILE_NOT_FOUND) {
    fwprintf_s(stderr, L"CLSID is not a registered local server.\n");
    return 1;
  } else if (result != ERROR_SUCCESS) {
    fwprintf_s(stderr, L"LocalServer32 query failed with code %ld.\n", result);
    return 1;
  }

  ::RegCloseKey(regKeyLocalServer);

  wprintf_s(
      L"CLSID is a local server. Its threading model will be determined by\n");
  wprintf_s(L"    the threading model of its proxy/stub class.\n");

  if (!gHasIid) {
    fwprintf_s(stderr, L"An IID must be provided to proceed any further.\n");
    return 1;
  }

  std::wstring subKeyProxyStubClsid(L"Interface\\");
  subKeyProxyStubClsid += gStrIid;
  subKeyProxyStubClsid += L"\\ProxyStubClsid32";

  wchar_t proxyStubClsidBuf[kGuidLenWithBracesInclNul] = {};
  DWORD numBytes = sizeof(proxyStubClsidBuf);
  result =
      ::RegGetValueW(HKEY_CLASSES_ROOT, subKeyProxyStubClsid.c_str(), nullptr,
                     RRF_RT_REG_SZ, nullptr, proxyStubClsidBuf, &numBytes);
  if (result != ERROR_SUCCESS) {
    fwprintf_s(stderr, L"Could not resolve IID's proxy/stub CLSID.\n");
    return 1;
  }

  std::variant<ThreadingModel, LSTATUS> proxyModel =
      GetClassThreadingModel(proxyStubClsidBuf);
  if (std::holds_alternative<LSTATUS>(proxyModel)) {
    fwprintf_s(stderr, L"Could not resolve proxy/stub threading model.\n");
    return 1;
  }

  wprintf_s(L"Proxy threading model is %s.\n",
            GetThreadingModelDescription(std::get<ThreadingModel>(proxyModel)));
  return 0;
}
