#include <Windows.h>
#include "../../API/RainmeterAPI.h"

#define _WIN32_DCOM

#include <comutil.h>
#include <wbemidl.h>
#include <wrl/client.h>
#include <wrl/wrappers/corewrappers.h>
#include <memory>
#include <thread>
#include <functional>

// Overview: This is a blank canvas on which to build your plugin.

// Note: GetString, ExecuteBang and an unnamed function for use as a section
// variable have been commented out. Uncomment any functions as needed. For more
// information, see the SDK docs:
// https://docs.rainmeter.net/developers/plugin/cpp/

struct ComRuntime {
  ComRuntime() {
    _com_util::CheckError(CoInitializeEx(nullptr, COINIT_MULTITHREADED));
  }
  ~ComRuntime() { CoUninitialize(); }
};

class Monitor {
 public:
  Monitor()
      : monitor_count_(0),
        cpu_temp_(0),
        gpu_temp_(0),
        cpu_fan_(0),
        gpu_fan_(0) {}

  void Add() {
    auto lock = lock_.LockExclusive();
    monitor_count_++;
    if (monitor_count_ == 1) {
      poll_thread_ =
          std::make_unique<std::thread>(std::bind(&Monitor::PollLoop, this));
    }
  }

  void Remove() {
    auto lock = lock_.LockExclusive();
    monitor_count_--;
    if (monitor_count_ == 0) {
      if (poll_thread_->joinable()) {
        poll_thread_->join();
      }
      poll_thread_.reset(nullptr);
    }
  }

  int CpuTemp() const { return cpu_temp_; }
  int GpuTemp() const { return gpu_temp_; }
  int CpuFan() const { return cpu_fan_; }
  int GpuFan() const { return gpu_fan_; }

 private:
  void PollLoop() {
    com_runtime_ = std::make_unique<ComRuntime>();

    _com_util::CheckError(CoInitializeSecurity(
        nullptr,                    // Security descriptor
        -1,                         // COM negotiates authentication service
        nullptr,                    // Authentication services
        nullptr,                    // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,  // Default authentication level for proxies
        RPC_C_IMP_LEVEL_IMPERSONATE,  // Default Impersonation level for proxies
        nullptr,                      // Authentication info
        EOAC_NONE,  // Additional capabilities of the client or server
        nullptr));  // Reserved

    _com_util::CheckError(
        CoCreateInstance(CLSID_WbemLocator, nullptr, CLSCTX_INPROC_SERVER,
                         IID_PPV_ARGS(locator_.ReleaseAndGetAddressOf())));

    _com_util::CheckError(locator_->ConnectServer(
        _bstr_t(L"root\\WMI"),                // namespace
        nullptr,                              // User name
        nullptr,                              // User password
        nullptr,                              // Locale
        NULL,                                 // Security flags
        nullptr,                              // Authority
        nullptr,                              // Context object
        service_.ReleaseAndGetAddressOf()));  // IWbemServices proxy

    _com_util::CheckError(
        CoSetProxyBlanket(service_.Get(), RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE,
                          nullptr, RPC_C_AUTHN_LEVEL_CALL,
                          RPC_C_IMP_LEVEL_IMPERSONATE, nullptr, EOAC_NONE));

    while (monitor_count_ > 0) {
      UpdateMeasures();
      std::this_thread::sleep_for(std::chrono::milliseconds(1000));
    }

    com_runtime_.reset(nullptr);
  }

  void UpdateMeasures() {
    Microsoft::WRL::ComPtr<IEnumWbemClassObject> wbem_enum;
    Microsoft::WRL::ComPtr<IWbemClassObject> wbem_object;
    ULONG wbem_count = 0;
    variant_t wbem_variant;

    _com_util::CheckError(service_->ExecQuery(
        _bstr_t("WQL"), _bstr_t("SELECT CPU FROM MSI_CPU"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, nullptr,
        wbem_enum.ReleaseAndGetAddressOf()));

    _com_util::CheckError(wbem_enum->Skip(WBEM_INFINITE, 1));
    _com_util::CheckError(wbem_enum->Next(
        WBEM_INFINITE, 1, wbem_object.ReleaseAndGetAddressOf(), &wbem_count));
    _com_util::CheckError(
        wbem_object->Get(L"CPU", 0, &wbem_variant, nullptr, nullptr));

    cpu_temp_ = wbem_variant.uiVal;

    _com_util::CheckError(service_->ExecQuery(
        _bstr_t("WQL"), _bstr_t("SELECT VGA FROM MSI_VGA"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, nullptr,
        wbem_enum.ReleaseAndGetAddressOf()));

    _com_util::CheckError(wbem_enum->Skip(WBEM_INFINITE, 1));
    _com_util::CheckError(wbem_enum->Next(
        WBEM_INFINITE, 1, wbem_object.ReleaseAndGetAddressOf(), &wbem_count));
    _com_util::CheckError(
        wbem_object->Get(L"VGA", 0, &wbem_variant, nullptr, nullptr));

    gpu_temp_ = wbem_variant.uiVal;

    _com_util::CheckError(service_->ExecQuery(
        _bstr_t("WQL"), _bstr_t("SELECT AP FROM MSI_AP"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, nullptr,
        wbem_enum.ReleaseAndGetAddressOf()));

    _com_util::CheckError(wbem_enum->Skip(WBEM_INFINITE, 2));
    _com_util::CheckError(wbem_enum->Next(
        WBEM_INFINITE, 1, wbem_object.ReleaseAndGetAddressOf(), &wbem_count));
    _com_util::CheckError(
        wbem_object->Get(L"AP", 0, &wbem_variant, nullptr, nullptr));
    const USHORT cpu_fan_low8 = wbem_variant.uiVal;

    _com_util::CheckError(wbem_enum->Next(
        WBEM_INFINITE, 1, wbem_object.ReleaseAndGetAddressOf(), &wbem_count));
    _com_util::CheckError(
        wbem_object->Get(L"AP", 0, &wbem_variant, nullptr, nullptr));
    const USHORT cpu_fan_high8 = wbem_variant.uiVal;

    if (cpu_fan_low8 != 0 || cpu_fan_high8 != 0) {
      cpu_fan_ = static_cast<int>(
          60000000.0 /
          (static_cast<double>(((cpu_fan_high8 << 8) + cpu_fan_low8) * 2) *
           62.5));
    } else {
      cpu_fan_ = 0;
    }

    _com_util::CheckError(wbem_enum->Next(
        WBEM_INFINITE, 1, wbem_object.ReleaseAndGetAddressOf(), &wbem_count));
    _com_util::CheckError(
        wbem_object->Get(L"AP", 0, &wbem_variant, nullptr, nullptr));
    const USHORT gpu_fan_low8 = wbem_variant.uiVal;

    _com_util::CheckError(wbem_enum->Next(
        WBEM_INFINITE, 1, wbem_object.ReleaseAndGetAddressOf(), &wbem_count));
    _com_util::CheckError(
        wbem_object->Get(L"AP", 0, &wbem_variant, nullptr, nullptr));
    const USHORT gpu_fan_high8 = wbem_variant.uiVal;

    if (gpu_fan_low8 != 0 || gpu_fan_high8 != 0) {
      gpu_fan_ = static_cast<int>(
          60000000.0 /
          (static_cast<double>(((gpu_fan_high8 << 8) + gpu_fan_low8) * 2) *
           62.5));
    } else {
      gpu_fan_ = 0;
    }
  }

 private:
  std::unique_ptr<ComRuntime> com_runtime_;
  Microsoft::WRL::ComPtr<IWbemLocator> locator_;
  Microsoft::WRL::ComPtr<IWbemServices> service_;

  Microsoft::WRL::Wrappers::SRWLock lock_;
  int monitor_count_;
  std::unique_ptr<std::thread> poll_thread_;

  int cpu_temp_;
  int gpu_temp_;
  int cpu_fan_;
  int gpu_fan_;
};

Monitor global_monitor;

struct Measure {
  Measure() {
    get_latest_value_ = [] { return 0; };
    global_monitor.Add();
  }
  ~Measure() { global_monitor.Remove(); }

  void Reload(LPCWSTR type) {
    if (_wcsicmp(type, L"CpuTemp") == 0) {
      get_latest_value_ = [] { return global_monitor.CpuTemp(); };
    } else if (_wcsicmp(type, L"GpuTemp") == 0) {
      get_latest_value_ = [] { return global_monitor.GpuTemp(); };
    } else if (_wcsicmp(type, L"CpuFan") == 0) {
      get_latest_value_ = [] { return global_monitor.CpuFan(); };
    } else if (_wcsicmp(type, L"GpuFan") == 0) {
      get_latest_value_ = [] { return global_monitor.GpuFan(); };
    } else {
      // error
      get_latest_value_ = [] { return 0; };
    }
  }

  int LatestValue() const { return get_latest_value_(); }

 private:
  std::function<int()> get_latest_value_;
};

PLUGIN_EXPORT void Initialize(void** data, void* rm) {
  Measure* measure = new Measure;
  *data = measure;
}

PLUGIN_EXPORT void Finalize(void* data) {
  Measure* measure = (Measure*)data;
  delete measure;
}

PLUGIN_EXPORT void Reload(void* data, void* rm, double* maxValue) {
  Measure* measure = (Measure*)data;
  LPCWSTR type = RmReadString(rm, L"MeasureType", L"");
  measure->Reload(type);
}

PLUGIN_EXPORT double Update(void* data) {
  Measure* measure = (Measure*)data;
  return measure->LatestValue();
}

// PLUGIN_EXPORT LPCWSTR GetString(void* data)
//{
//	Measure* measure = (Measure*)data;
//	return L"";
//}

// PLUGIN_EXPORT void ExecuteBang(void* data, LPCWSTR args)
//{
//	Measure* measure = (Measure*)data;
//}

// PLUGIN_EXPORT LPCWSTR (void* data, const int argc, const WCHAR* argv[])
//{
//	Measure* measure = (Measure*)data;
//	return nullptr;
//}
