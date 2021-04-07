/* -*- Mode: C++; tab-width: 8; indent-tabs-mode: nil; c-basic-offset: 2 -*- */
/* vim: set ts=8 sts=2 et sw=2 tw=80: */
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

#include <optional>
#include <string>
#include <variant>

#include <comdef.h>
#include <stdio.h>

#include <objbase.h>
#include <windows.h>

_COM_SMARTPTR_TYPEDEF(IAgileObject, IID_IAgileObject);

template <typename T, size_t N>
static inline constexpr size_t ArrayLength(T (&aArr)[N]) {
  return N;
}

static constexpr int kGuidLenWithBracesInclNul = 39;
static constexpr int kGuidLenWithBracesExclNul = kGuidLenWithBracesInclNul - 1;
static constexpr size_t kThreadingModelBufCharLen = ArrayLength(L"Apartment");
static const CLSID CLSID_FreeThreadedMarshaler = {
    0x0000033A,
    0x0000,
    0x0000,
    {0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}};
static const CLSID CLSID_UniversalMarshaler = {
    0x00020424,
    0x0000,
    0x0000,
    {0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}};

static wchar_t gStrClsid[kGuidLenWithBracesInclNul];
static wchar_t gStrIid[kGuidLenWithBracesInclNul];
static const wchar_t *gProgID;
static std::optional<CLSID> gClsid;
static std::optional<IID> gIid;
static bool gDescriptive;
static bool gVerbose;

static void Usage(const wchar_t *aArgv0, const wchar_t *aMsg = nullptr) {
  if (aMsg) {
    fwprintf_s(stderr, L"Error: %s\n\n", aMsg);
  }

  fwprintf_s(stderr, L"Usage: %s [-d] [-v] <ProgID or CLSID> [IID]\n\n",
             aArgv0);
  fwprintf_s(stderr, L"Where:\n\n");
  fwprintf_s(stderr,
             L"\t-d\tDescriptive mode: include additional descriptive text in "
             L"output\n");
  fwprintf_s(stderr, L"\t-v\tVerbose mode (implies -d)\n");
  fwprintf_s(stderr,
             L"\n\tIID is optional, but omitting it may result in incomplete "
             L"output.\n");
  fwprintf_s(stderr, L"\n\tCLSID and IID must be specified in registry "
                     L"format,\n\tincluding dashes and curly braces.\n\t"
                     L"For example: {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}\n");
}

static bool ParseArgv(const int argc, wchar_t *argv[]) {
  if (argc < 2) {
    Usage(argv[0]);
    return false;
  }

  for (int i = 1; i < argc; ++i) {
    if (argv[i][0] == L'-' || argv[i][0] == L'/') {
      if (argv[i][1] == L'd') {
        gDescriptive = true;
      } else if (argv[i][1] == L'v') {
        gDescriptive = true;
        gVerbose = true;
      }
    } else if (argv[i][0] == L'{' &&
               wcslen(argv[i]) == kGuidLenWithBracesExclNul) {
      GUID guid;

      if (!gClsid.has_value()) {
        if (FAILED(::CLSIDFromString(argv[i], &guid))) {
          Usage(argv[0], L"Failed to parse CLSID");
          return false;
        }

        wcscpy_s(gStrClsid, argv[i]);
        gClsid.emplace(guid);
      } else {
        if (FAILED(::IIDFromString(argv[i], &guid))) {
          Usage(argv[0], L"Failed to parse IID");
          return false;
        }

        wcscpy_s(gStrIid, argv[i]);
        gIid.emplace(guid);
      }
    } else if (!gClsid.has_value()) {
      // ProgID?
      CLSID clsid;

      if (FAILED(::CLSIDFromProgID(argv[i], &clsid))) {
        Usage(argv[0], L"Invalid ProgID");
        return false;
      }

      gClsid.emplace(clsid);

      if (!::StringFromGUID2(clsid, gStrClsid,
                             static_cast<int>(ArrayLength(gStrClsid)))) {
        Usage(argv[0], L"Failed converting CLSID to string");
        return false;
      }

      gProgID = argv[i];
    }
  }

  if (!gClsid) {
    Usage(argv[0], L"You must provide either a CLSID or a ProgID.");
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
  Manifest, // <-- unsupported by us (no public API), but still possible
  AgileObject,
};

class ComClassThreadInfo final {
public:
  constexpr ComClassThreadInfo(const ThreadingModel aThdModel,
                               const Provenance aProvenance)
      : mThreadingModel7(aThdModel), mProvenance7(aProvenance),
        mThreadingModel8(aThdModel), mProvenance8(aProvenance) {}

  constexpr ComClassThreadInfo(const ThreadingModel aThdModel7,
                               const Provenance aProvenance7,
                               const ThreadingModel aThdModel8,
                               const Provenance aProvenance8)
      : mThreadingModel7(aThdModel7), mProvenance7(aProvenance7),
        mThreadingModel8(aThdModel8), mProvenance8(aProvenance8) {}

  std::wstring GetDescription(const ClassType aClassType) const;

  ComClassThreadInfo
  CheckObjectCapabilities(REFCLSID, const std::optional<IID> &aOptIid) const;

  ComClassThreadInfo(const ComClassThreadInfo &) = default;
  ComClassThreadInfo(ComClassThreadInfo &&) = default;
  ComClassThreadInfo &operator=(const ComClassThreadInfo &) = delete;
  ComClassThreadInfo &operator=(ComClassThreadInfo &&) = delete;

private:
  static std::wstring
  GetThreadingModelDescription(const ThreadingModel aThdModel);
  static const wchar_t *GetProvenanceDescription(const Provenance aProvenance);

private:
  const ThreadingModel mThreadingModel7;
  const Provenance mProvenance7;
  const ThreadingModel mThreadingModel8;
  const Provenance mProvenance8;
};

std::wstring ComClassThreadInfo::GetThreadingModelDescription(
    const ThreadingModel aThdModel) {
  std::wstring result;

  switch (aThdModel) {
  case ThreadingModel::STA:
    result = L"Single-threaded";
    if (gDescriptive) {
      result += L":\n\tProxying is required to access from any other apartment";
    }
    break;
  case ThreadingModel::MTA:
    result = L"Multi-threaded";
    if (gDescriptive) {
      result += L":\n\tProxying is required to access from any single-threaded "
                L"apartment";
    }

    break;
  case ThreadingModel::Both:
    result = L"Both";
    if (gDescriptive) {
      result += L":\n\tEither single-threaded or multi-threaded, but "
                L"mutually-exclusive";
    }

    break;
  case ThreadingModel::Neutral:
    result = L"Thread-neutral";
    if (gDescriptive) {
      result +=
          L":\n\tThis object may be invoked by any thread residing in any "
          L"apartment";
    }

    break;
  default:
    return result;
  };

  result += L".\n";
  return result;
}

const wchar_t *
ComClassThreadInfo::GetProvenanceDescription(const Provenance aProvenance) {
  switch (aProvenance) {
  case Provenance::Registry:
    return L"System registry.\n";
  case Provenance::FreeThreadedMarshaler:
    return L"Free-threaded marshaler.\n";
  case Provenance::Manifest:
    return L"Manifest.\n";
  case Provenance::AgileObject:
    return L"IAgileObject.\n";
  default:
    return nullptr;
  }
}

std::wstring
ComClassThreadInfo::GetDescription(const ClassType aClassType) const {
  std::wstring result(aClassType == ClassType::Server ? L"Server " : L"Proxy ");
  if (mThreadingModel7 == mThreadingModel8) {
    result += L"threading model: ";
    result += GetThreadingModelDescription(mThreadingModel7);
    result += L"Provenance: ";
    result += GetProvenanceDescription(mProvenance7);
    return result;
  }

  result += L"has different threading models depending on the application's "
            L"supported OS version.\n\n";
  result += L"For applications indicating compatibility with Windows 7 or "
            L"older,\nthe threading model is ";
  result += GetThreadingModelDescription(mThreadingModel7);
  result += L"Provenance: ";
  result += GetProvenanceDescription(mProvenance7);
  result += L"\n\nFor applications indicating compatibility with Windows 8 or "
            L"newer,\nthe threading model is ";
  result += GetThreadingModelDescription(mThreadingModel8);
  result += L"Provenance: ";
  result += GetProvenanceDescription(mProvenance8);
  return result;
}

static bool HasDllSurrogate(const wchar_t *aStrClsid) {
  if (gVerbose) {
    wprintf_s(L"Checking for DLL surrogate... ");
  }

  std::wstring subKeyClsid(L"CLSID\\");
  subKeyClsid += aStrClsid;

  wchar_t appIdBuf[kGuidLenWithBracesInclNul] = {};
  DWORD numBytes = sizeof(appIdBuf);
  LSTATUS result =
      ::RegGetValueW(HKEY_CLASSES_ROOT, subKeyClsid.c_str(), L"AppID",
                     RRF_RT_REG_SZ, nullptr, appIdBuf, &numBytes);
  if (result != ERROR_SUCCESS) {
    if (gVerbose) {
      wprintf_s(L"No AppID.\n");
    }

    return false;
  }

  std::wstring subKeyAppid(L"AppID\\");
  subKeyAppid += appIdBuf;

  wchar_t surrogate[MAX_PATH + 1] = {};
  numBytes = sizeof(surrogate);
  result =
      ::RegGetValueW(HKEY_CLASSES_ROOT, subKeyAppid.c_str(), L"DllSurrogate",
                     RRF_RT_REG_SZ, nullptr, surrogate, &numBytes);
  if (result != ERROR_SUCCESS) {
    if (gVerbose) {
      wprintf_s(L"AppID does not have DllSurrogate value.\n");
    }

    return false;
  }

  if (gVerbose) {
    std::wstring strSurrogate;
    if (surrogate[0]) {
      strSurrogate = L"\"";
      strSurrogate += surrogate;
      strSurrogate += L"\"";
    } else {
      strSurrogate = L"System default (dllhost.exe)";
    }

    wprintf_s(L"%s.\n", strSurrogate.c_str());
  }

  return true;
}

static std::variant<ComClassThreadInfo, LSTATUS>
GetClassThreadingModel(const wchar_t *aStrClsid) {
  std::wstring subKeyClsid(L"CLSID\\");
  subKeyClsid += aStrClsid;

  std::wstring subKeyInprocServer(subKeyClsid);
  subKeyInprocServer += L"\\InprocServer32";

  wchar_t threadingModelBuf[kThreadingModelBufCharLen] = {};
  DWORD numBytes = sizeof(threadingModelBuf);
  LSTATUS result = ::RegGetValueW(HKEY_CLASSES_ROOT, subKeyInprocServer.c_str(),
                                  L"ThreadingModel", RRF_RT_REG_SZ, nullptr,
                                  threadingModelBuf, &numBytes);
  if (result == ERROR_FILE_NOT_FOUND) {
    // Check if the subkey exists, at least. If not, the class is not
    // registered at all. If it is registered, then we're just missing the
    // ThreadingModel registry value.
    HKEY regKeyInprocServer;
    LSTATUS subkeyOpened =
        ::RegOpenKeyEx(HKEY_CLASSES_ROOT, subKeyInprocServer.c_str(), 0,
                       KEY_READ, &regKeyInprocServer);
    if (subkeyOpened != ERROR_SUCCESS) {
      return result;
    }

    // The class *is* registered, just missing its ThreadingModel. Close the key
    // and fall-through for additional processing.
    ::RegCloseKey(regKeyInprocServer);
  } else if (result != ERROR_SUCCESS) {
    return result;
  }

  if (gVerbose) {
    wchar_t serverDllPath[MAX_PATH + 1] = {};
    numBytes = sizeof(serverDllPath) - sizeof(wchar_t);
    LSTATUS pathResult =
        ::RegGetValueW(HKEY_CLASSES_ROOT, subKeyInprocServer.c_str(), nullptr,
                       RRF_RT_REG_SZ, nullptr, serverDllPath, &numBytes);
    if (pathResult == ERROR_SUCCESS) {
      wprintf_s(L"Path to server DLL: \"%s\"\n", serverDllPath);
    } else {
      wprintf_s(L"Failed to retrieve path to server DLL, code %ld!\n",
                pathResult);
    }
  }

  // Empty or non-existent ThreadingModel implies STA.
  if (result == ERROR_FILE_NOT_FOUND || threadingModelBuf[0] == 0 ||
      !_wcsicmp(threadingModelBuf, L"Apartment")) {
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

class Apartment final {
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
  HRESULT GetHResult() const { return mHr; }

  Apartment(const Apartment &) = delete;
  Apartment(Apartment &&) = delete;
  Apartment &operator=(const Apartment &) = delete;
  Apartment &operator=(Apartment &&) = delete;

private:
  const HRESULT mHr;
};

ComClassThreadInfo ComClassThreadInfo::CheckObjectCapabilities(
    REFCLSID aClsid, const std::optional<IID> &aOptIid) const {
  if (mThreadingModel7 == ThreadingModel::Neutral) {
    // We're already neutral, these additional checks are unnecessary.
    return *this;
  }

  ThreadingModel thdModel7 = mThreadingModel7;
  Provenance prov7 = mProvenance7;
  ThreadingModel thdModel8 = mThreadingModel7;
  Provenance prov8 = mProvenance7;

  if (gVerbose) {
    wprintf_s(L"Entering apartment... ");
  }

  Apartment apt(mThreadingModel7);
  if (!apt) {
    if (gVerbose) {
      wprintf_s(L"Failed with HRESULT 0x%08lX.\n", apt.GetHResult());
    }

    wprintf_s(L"WARNING: Could not enter a test apartment. Results might be "
              L"incomplete!\n");
    return *this;
  }

  if (gVerbose) {
    wprintf_s(L"OK.\nCreating object... ");
  }

  IUnknownPtr punk;
  HRESULT hr = punk.CreateInstance(aClsid, nullptr, CLSCTX_INPROC_SERVER);
  if (FAILED(hr)) {
    if (gVerbose) {
      wprintf_s(L"Failed with HRESULT 0x%08lX.\n", hr);
    }

    wprintf_s(L"WARNING: Could not create a test instance. Results might be "
              L"incomplete!\n");
    return *this;
  } else if (gVerbose) {
    wprintf_s(L"OK.\n");
  }

  if (gVerbose) {
    wprintf_s(L"Querying for IAgileObject... ");
  }

  IAgileObjectPtr agile;
  hr = punk.QueryInterface(IID_IAgileObject, &agile);
  if (SUCCEEDED(hr)) {
    if (gVerbose) {
      wprintf_s(L"Found.\n");
    }

    thdModel8 = ThreadingModel::Neutral;
    prov8 = Provenance::AgileObject;
  } else if (gVerbose) {
    if (hr == E_NOINTERFACE) {
      wprintf_s(L"Not found.\n");
    } else {
      wprintf_s(L"Failed with HRESULT 0x%08lX.\n", hr);
    }
  }

  // We need an IID to do any further checks
  if (!aOptIid.has_value()) {
    wprintf_s(L"WARNING: IID required to query for free-threaded "
              L"marshaler.\n\tResults might be incomplete!\n");
    return ComClassThreadInfo{thdModel7, prov7, thdModel8, prov8};
  }

  if (gVerbose) {
    wprintf_s(L"Querying for IMarshal... ");
  }

  // Check for the free-threaded marshaler
  IMarshalPtr marshal;
  hr = punk.QueryInterface(IID_IMarshal, &marshal);
  if (FAILED(hr)) {
    if (gVerbose) {
      if (hr == E_NOINTERFACE) {
        wprintf_s(L"Not found.\n");
      } else {
        wprintf_s(L"Failed with HRESULT 0x%08lX.\n", hr);
      }
    }

    return ComClassThreadInfo{thdModel7, prov7, thdModel8, prov8};
  } else if (gVerbose) {
    wprintf_s(L"Found.\nChecking whether object aggregates the free-threaded "
              L"marshaler... ");
  }

  CLSID unmarshalClass;
  hr = marshal->GetUnmarshalClass(aOptIid.value(), nullptr, MSHCTX_INPROC,
                                  nullptr, MSHLFLAGS_NORMAL, &unmarshalClass);
  if (FAILED(hr)) {
    if (gVerbose) {
      wprintf_s(L"Failed with HRESULT 0x%08lX.\n", hr);
    }

    return ComClassThreadInfo{thdModel7, prov7, thdModel8, prov8};
  }

  if (unmarshalClass == CLSID_FreeThreadedMarshaler) {
    if (gVerbose) {
      wprintf_s(L"Yes.\n");
    }

    thdModel7 = ThreadingModel::Neutral;
    prov7 = Provenance::FreeThreadedMarshaler;
    if (thdModel8 != ThreadingModel::Neutral) {
      thdModel8 = ThreadingModel::Neutral;
      prov8 = Provenance::FreeThreadedMarshaler;
    }
  } else if (gVerbose) {
    wprintf_s(L"No.\n");
  }

  return ComClassThreadInfo{thdModel7, prov7, thdModel8, prov8};
}

static int CheckProxyForInterface(const wchar_t *aStrIid) {
  if (gVerbose) {
    wprintf_s(L"Checking interface's proxy/stub class...\n");
  }

  std::wstring subKeyProxyStubClsid(L"Interface\\");
  subKeyProxyStubClsid += aStrIid;
  subKeyProxyStubClsid += L"\\ProxyStubClsid32";

  wchar_t proxyStubClsidBuf[kGuidLenWithBracesInclNul] = {};
  DWORD numBytes = sizeof(proxyStubClsidBuf);
  LSTATUS result =
      ::RegGetValueW(HKEY_CLASSES_ROOT, subKeyProxyStubClsid.c_str(), nullptr,
                     RRF_RT_REG_SZ, nullptr, proxyStubClsidBuf, &numBytes);
  if (result != ERROR_SUCCESS) {
    fwprintf_s(stderr, L"Could not resolve IID's proxy/stub CLSID.\n");
    return 1;
  }

  CLSID proxyStubClsid;
  if (FAILED(::CLSIDFromString(proxyStubClsidBuf, &proxyStubClsid))) {
    fwprintf_s(stderr, L"Could not parse proxy/stub CLSID.\n");
    return 1;
  }

  if (gVerbose) {
    if (proxyStubClsid == CLSID_UniversalMarshaler) {
      wprintf_s(
          L"This interface uses OLE Automation for proxy/stub marshaling.\n");
    } else {
      wchar_t strProxyStubClsid[kGuidLenWithBracesInclNul] = {};
      if (::StringFromGUID2(proxyStubClsid, strProxyStubClsid,
                            kGuidLenWithBracesInclNul)) {
        wprintf_s(L"CLSID for proxy/stub: %s\n", strProxyStubClsid);
      }
    }
  }

  std::variant<ComClassThreadInfo, LSTATUS> proxyModel =
      GetClassThreadingModel(proxyStubClsidBuf);
  if (std::holds_alternative<LSTATUS>(proxyModel)) {
    fwprintf_s(stderr, L"Could not resolve proxy/stub threading model.\n");
    return 1;
  }

  std::wstring output =
      std::get<ComClassThreadInfo>(proxyModel).GetDescription(ClassType::Proxy);
  wprintf_s(L"%s", output.c_str());
  return 0;
}

int wmain(int argc, wchar_t *argv[]) {
  if (!ParseArgv(argc, argv)) {
    return 1;
  }

  std::variant<ComClassThreadInfo, LSTATUS> inprocModel =
      GetClassThreadingModel(gStrClsid);
  if (std::holds_alternative<ComClassThreadInfo>(inprocModel)) {
    std::wstring output = std::get<ComClassThreadInfo>(inprocModel)
                              .CheckObjectCapabilities(gClsid.value(), gIid)
                              .GetDescription(ClassType::Server);
    wprintf_s(L"When instantiating in-process (via CLSCTX_INPROC_SERVER):\n%s",
              output.c_str());
    if (!HasDllSurrogate(gStrClsid) && !gIid.has_value()) {
      return 0;
    }

    wprintf_s(
        L"\nThis class may optionally be instantiated out-of-process\n\t(via "
        L"CLSCTX_LOCAL_SERVER):\n");
    if (gVerbose) {
      wprintf_s(L"In this case its threading model will be determined by the "
                L"threading model of\n\tits proxy/stub class.\n");
    }

    if (gIid.has_value()) {
      // Ignore return value since we already have some success
      CheckProxyForInterface(gStrIid);
      return 0;
    }

    fwprintf_s(stderr, L"An IID must be provided to proceed any further.\n");
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
    fwprintf_s(stderr,
               L"CLSID is not a persistently-registered local server.\n");
    // Try querying for a class object that might have been registered at
    // runtime. We need to enter an apartment before calling CoGetClassObject.
    if (gVerbose) {
      wprintf_s(L"Entering apartment...\n");
    }

    Apartment apt(ThreadingModel::MTA);
    if (!apt) {
      if (gVerbose) {
        wprintf_s(L"Failed with HRESULT 0x%08lX.\n", apt.GetHResult());
      }

      wprintf_s(L"WARNING: Could not enter a test apartment. Results might be "
                L"incomplete!\n");
      return 1;
    }

    if (gVerbose) {
      wprintf_s(L"Attempting to resolve via CoGetClassObject... ");
    }

    IClassFactoryPtr classFactory;
    HRESULT hr = ::CoGetClassObject(
        gClsid.value(), CLSCTX_LOCAL_SERVER, nullptr, IID_IClassFactory,
        reinterpret_cast<void **>(
            static_cast<IClassFactory **>(&classFactory)));
    if (FAILED(hr)) {
      if (gVerbose) {
        wprintf_s(L"\nFailed with HRESULT 0x%08lX.\n", hr);
      }

      fwprintf_s(stderr,
                 L"CLSID is not a temporarily-registered local server.\n");
      return 1;
    }

    if (gVerbose) {
      wprintf_s(L"OK.\n");
    }
  } else if (result != ERROR_SUCCESS) {
    fwprintf_s(stderr, L"LocalServer32 query failed with code %ld.\n", result);
    return 1;
  } else {
    ::RegCloseKey(regKeyLocalServer);
  }

  wprintf_s(L"When instantiating out-of-process (via CLSCTX_LOCAL_SERVER):\n");

  if (gVerbose) {
    wprintf_s(
        L"Its threading model will be determined by the threading model of\n\t"
        L"its proxy/stub class.\n");
  }

  if (!gIid.has_value()) {
    fwprintf_s(stderr,
               L"ERROR: An IID must be provided to proceed any further.\n");
    return 1;
  }

  return CheckProxyForInterface(gStrIid);
}
