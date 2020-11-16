#include <optional>
#include <string>
#include <variant>

#include <stdio.h>
#include <comdef.h>

#include <windows.h>
#include <objbase.h>

#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "ole32.lib")

// TODO ASK: If there's an AppID, try both inproc and local

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
static const wchar_t* gProgID;
static std::optional<CLSID> gClsid;
static std::optional<IID> gIid;
static bool gVerbose;

static void Usage(const wchar_t* aArgv0, const wchar_t* aMsg = nullptr) {
  if (aMsg) {
    fwprintf_s(stderr, L"Error: %s\n\n", aMsg);
  }

  fwprintf_s(stderr, L"Usage: %s [-v] <ProgID or CLSID> [IID]\n\n", aArgv0);
  fwprintf_s(stderr,
             L"    CLSID and IID must be specified in registry format,\n    "
             L"including dashes and curly braces\n");
}

static bool ParseArgv(const int argc, wchar_t* argv[]) {
  if (argc < 2) {
    Usage(argv[0]);
    return false;
  }

  for (size_t i = 1; i < argc; ++i) {
    if (argv[i][0] == L'-' || argv[i][0] == L'/') {
      if (argv[i][1] == L'v') {
        gVerbose = true;
      }
    } else if (argv[i][0] == '{' &&
               wcslen(argv[i]) == kGuidLenWithBracesExclNul) {
      GUID guid;

      if (!gClsid.has_value()) {
        if (FAILED(::CLSIDFromString(argv[i], &guid))) {
          Usage(argv[i], L"Failed to parse CLSID");
          return false;
        }

        wcscpy_s(gStrClsid, argv[i]);
        gClsid.emplace(guid);
      } else {
        if (FAILED(::IIDFromString(argv[i], &guid))) {
          Usage(argv[i], L"Failed to parse IID");
          return false;
        }

        wcscpy_s(gStrIid, argv[i]);
        gIid.emplace(guid);
      }
    } else if (!gClsid.has_value()) {
      // ProgID?
      GUID guid;

      if (FAILED(::CLSIDFromProgID(argv[i], &guid))) {
        Usage(argv[0], L"Invalid ProgID");
        return false;
      }

      gClsid.emplace(guid);

      if (!::StringFromGUID2(guid, gStrClsid, ArrayLength(gStrClsid))) {
        return false;
      }

      gProgID = argv[i];
    }
  }

  if (!gClsid) {
    return false;
  }

  if (gVerbose) {
    wprintf_s(L"Using CLSID %s\n", gStrClsid);
    if (gProgID) {
      wprintf_s(L"Obtained from ProgID \"%s\"\n", gProgID);
    }

    if (gIid) {
      wprintf_s(L"Using IID %s\n", gStrIid);
    }
  }

  return true;
}

enum class ClassType {
  Server,
  Proxy,
};

enum class ThreadingModel {
  STA,
  MTA,
  Both,
  Neutral,
};

enum class Provenance {
  Registry,
  FreeThreadedMarshaler,
  Manifest,  // <-- unsupported by us (no public API), but still possible
  AgileObject,
};

class ComClassThreadInfo {
 public:
  constexpr ComClassThreadInfo(const ThreadingModel aThdModel,
                               const Provenance aProvenance)
      : mThreadingModel7(aThdModel),
        mProvenance7(aProvenance),
        mThreadingModel8(aThdModel),
        mProvenance8(aProvenance) {}

  constexpr ComClassThreadInfo(const ThreadingModel aThdModel7,
                               const Provenance aProvenance7,
                               const ThreadingModel aThdModel8,
                               const Provenance aProvenance8)
      : mThreadingModel7(aThdModel7),
        mProvenance7(aProvenance7),
        mThreadingModel8(aThdModel8),
        mProvenance8(aProvenance8) {}

  std::wstring GetDescription(const ClassType aClassType);

  ComClassThreadInfo CheckObjectCapabilities(
      REFCLSID, const std::optional<IID>& aOptIid) const;

 private:
  static const wchar_t* GetThreadingModelDescription(
      const ThreadingModel aThdModel);
  static const wchar_t* GetProvenanceDescription(const Provenance aProvenance);

 private:
  const ThreadingModel mThreadingModel7;
  const Provenance mProvenance7;
  const ThreadingModel mThreadingModel8;
  const Provenance mProvenance8;
};

const wchar_t* ComClassThreadInfo::GetThreadingModelDescription(
    const ThreadingModel aThdModel) {
  switch (aThdModel) {
    case ThreadingModel::STA:
      return L"single-threaded:\nProxying is required to access from any other "
             L"apartment.\n";
    case ThreadingModel::MTA:
      return L"multi-threaded:\nProxying is required to access from any "
             L"single-threaded apartment.\n";
    case ThreadingModel::Both:
      return L"either single-threaded or multi-threaded,\n    but are "
             L"mutually-exclusive.\n";
    case ThreadingModel::Neutral:
      return L"thread-neutral:\nThis class may be invoked by any thread within "
             L"any apartment.\n";
    default:
      return nullptr;
  };
}

const wchar_t* ComClassThreadInfo::GetProvenanceDescription(
    const Provenance aProvenance) {
  switch (aProvenance) {
    case Provenance::Registry:
      return L"The object's threading model was determined via the system "
             L"registry.\n";
    case Provenance::FreeThreadedMarshaler:
      return L"The object aggregates the free-threaded marshaler.\n";
    case Provenance::Manifest:
      return L"The object's threading model was determined via a manifest.\n";
    case Provenance::AgileObject:
      return L"The object supports the IAgileObject interface.\n";
    default:
      return nullptr;
  }
}

std::wstring ComClassThreadInfo::GetDescription(const ClassType aClassType) {
  std::wstring result(aClassType == ClassType::Server ? L"Server " : L"Proxy ");
  if (mThreadingModel7 == mThreadingModel8) {
    result += L"threading model is ";
    result += GetThreadingModelDescription(mThreadingModel7);
    result += GetProvenanceDescription(mProvenance7);
    return result;
  }

  result += L"has different threading models depending on OS version.\n\n";
  result += L"On Windows 7, the threading model is ";
  result += GetThreadingModelDescription(mThreadingModel7);
  result += GetProvenanceDescription(mProvenance7);
  result += L"\n\nOn Windows 8 and newer, the threading model is ";
  result += GetThreadingModelDescription(mThreadingModel8);
  result += GetProvenanceDescription(mProvenance8);
  return result;
}

static std::variant<ComClassThreadInfo, LSTATUS> GetClassThreadingModel(
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
    return ComClassThreadInfo{ThreadingModel::STA, Provenance::Registry};
  }

  if (!_wcsicmp(threadingModelBuf, L"Free")) {
    return ComClassThreadInfo{ThreadingModel::MTA, Provenance::Registry};
  }

  if (!_wcsicmp(threadingModelBuf, L"Both")) {
    return ComClassThreadInfo{ThreadingModel::Both, Provenance::Registry};
  }

  if (!_wcsicmp(threadingModelBuf, L"Neutral")) {
    return ComClassThreadInfo{ThreadingModel::Neutral, Provenance::Registry};
  }

  return static_cast<LSTATUS>(ERROR_UNIDENTIFIED_ERROR);
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

ComClassThreadInfo ComClassThreadInfo::CheckObjectCapabilities(
    REFCLSID aClsid, const std::optional<IID>& aOptIid) const {
  if (mThreadingModel7 == ThreadingModel::Neutral) {
    // We're already neutral, these additional checks are unnecessary.
    return *this;
  }

  ThreadingModel thdModel7 = mThreadingModel7;
  Provenance prov7 = mProvenance7;
  ThreadingModel thdModel8 = mThreadingModel7;
  Provenance prov8 = mProvenance7;

  Apartment apt(mThreadingModel7);
  if (!apt) {
    return *this;
  }

  if (gVerbose) {
    wprintf_s(L"Creating object...\n");
  }

  IUnknownPtr punk;
  HRESULT hr =
      ::CoCreateInstance(aClsid, nullptr, CLSCTX_INPROC_SERVER, IID_IUnknown,
                         reinterpret_cast<void**>(&punk));
  if (FAILED(hr)) {
    return *this;
  }

  if (gVerbose) {
    wprintf_s(L"Querying for IAgileObject...");
  }

  IUnknownPtr agile;
  hr = punk->QueryInterface(IID_IAgileObject, reinterpret_cast<void**>(&agile));
  if (SUCCEEDED(hr)) {
    if (gVerbose) {
      wprintf_s(L" found.\n");
    }

    thdModel8 = ThreadingModel::Neutral;
    prov8 = Provenance::AgileObject;
  } else if (gVerbose) {
    wprintf_s(L" not found.\n");
  }

  if (!aOptIid.has_value()) {
    wprintf_s(L"WARNING: IID required to query for free-threaded marshaler.\n");
    wprintf_s(L"         Results might be incomplete!\n");
    // We need an IID to do any further checks
    return ComClassThreadInfo{thdModel7, prov7, thdModel8, prov8};
  }

  if (gVerbose) {
    wprintf_s(L"Querying for IMarshal...");
  }

  // Check for the free-threaded marshaler
  IMarshalPtr marshal;
  hr = punk->QueryInterface(IID_IMarshal, reinterpret_cast<void**>(&marshal));
  if (FAILED(hr)) {
    if (gVerbose) {
      wprintf_s(L" not found.\n");
    }

    return ComClassThreadInfo{thdModel7, prov7, thdModel8, prov8};
  } else if (gVerbose) {
    wprintf_s(
        L" found.\nChecking whether object uses the free-threaded "
        L"marshaler...");
  }

  CLSID unmarshalClass;
  hr = marshal->GetUnmarshalClass(aOptIid.value(), nullptr, MSHCTX_INPROC,
                                  nullptr, MSHLFLAGS_NORMAL, &unmarshalClass);
  if (FAILED(hr)) {
    if (gVerbose) {
      wprintf_s(L" failed.\n");
    }

    return ComClassThreadInfo{thdModel7, prov7, thdModel8, prov8};
  }

  if (unmarshalClass == CLSID_FreeThreadedMarshaler) {
    if (gVerbose) {
      wprintf_s(L" yes.\n");
    }

    thdModel7 = ThreadingModel::Neutral;
    prov7 = Provenance::FreeThreadedMarshaler;
    if (thdModel8 != ThreadingModel::Neutral) {
      thdModel8 = ThreadingModel::Neutral;
      prov8 = Provenance::FreeThreadedMarshaler;
    }
  } else if (gVerbose) {
    wprintf_s(L" no.\n");
  }

  return ComClassThreadInfo{thdModel7, prov7, thdModel8, prov8};
}

int wmain(int argc, wchar_t* argv[]) {
  if (!ParseArgv(argc, argv)) {
    return 1;
  }

  std::variant<ComClassThreadInfo, LSTATUS> inprocModel =
      GetClassThreadingModel(gStrClsid);
  if (std::holds_alternative<ComClassThreadInfo>(inprocModel)) {
    std::wstring output = std::get<ComClassThreadInfo>(inprocModel)
                              .CheckObjectCapabilities(gClsid.value(), gIid)
                              .GetDescription(ClassType::Server);
    wprintf_s(L"%s", output.c_str());
    return 0;
  }

  LSTATUS result = std::get<LSTATUS>(inprocModel);
  if (result == ERROR_FILE_NOT_FOUND) {
    if (gVerbose) {
      wprintf_s(L"Class is not a registered in-process server.\n");
    }
  } else {
    fwprintf_s(stderr, L"InprocServer32 query failed with code %ld.\n", result);
    return 1;
  }

  if (gVerbose) {
    wprintf_s(L"Attempting to resolve as a local server...\n");
  }

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

  if (gVerbose) {
    wprintf_s(
        L"CLSID is a local server. Its threading model will be determined "
        L"by\n");
    wprintf_s(L"    the threading model of its proxy/stub class.\n");
  }

  if (!gIid.has_value()) {
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

  std::variant<ComClassThreadInfo, LSTATUS> proxyModel =
      GetClassThreadingModel(proxyStubClsidBuf);
  if (std::holds_alternative<LSTATUS>(proxyModel)) {
    fwprintf_s(stderr, L"Could not resolve proxy/stub threading model.\n");
    return 1;
  }

  CLSID proxyStubClsid;
  if (FAILED(::CLSIDFromString(proxyStubClsidBuf, &proxyStubClsid))) {
    fwprintf_s(stderr, L"Could not parse proxy/stub CLSID.\n");
    return 1;
  }

  std::wstring output = std::get<ComClassThreadInfo>(proxyModel)
                            .CheckObjectCapabilities(proxyStubClsid, gIid)
                            .GetDescription(ClassType::Proxy);
  wprintf_s(L"%s", output.c_str());
  return 0;
}
