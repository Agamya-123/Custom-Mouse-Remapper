#define WIN32_LEAN_AND_MEAN
#include <winsock2.h>
#include <ws2tcpip.h>
#include <windows.h>
#include <shellapi.h>
#include <commdlg.h>

#include <algorithm>
#include <atomic>
#include <cctype>
#include <cstdio>
#include <cstdlib>
#include <fstream>
#include <sstream>
#include <string>
#include <thread>
#include <unordered_map>
#include <vector>

namespace {

enum class ActionType { None, Keys, Run, Open, Text, Macro };

struct Action {
  ActionType type = ActionType::None;
  std::vector<WORD> keys;
  std::string payload;
};

struct Config {
  Action button4;
  Action button5;
  bool suspendInFullscreen = true;
  int dpi = 800;
};

Config g_config;
HWND g_mainWindow = nullptr;
HWND g_settingsWindow = nullptr;
NOTIFYICONDATAA g_tray = {};
std::string g_configPath;
std::thread g_statusServerThread;
std::atomic<bool> g_statusServerStop{false};
std::atomic<SOCKET> g_statusListenSocket{INVALID_SOCKET};

std::atomic<unsigned long long> g_rawInputEvents{0};
std::atomic<unsigned int> g_pollRateHz{0};
std::atomic<unsigned int> g_mouseButtons{0};
std::atomic<unsigned int> g_mouseSampleRate{0};
std::atomic<unsigned long long> g_dpiChangeEvents{0};
std::atomic<unsigned long long> g_lastRateWindowTick{0};
std::atomic<unsigned long long> g_lastRateWindowCount{0};

// Core Globals
std::atomic<bool> g_macroRecording{false};
std::atomic<bool> g_appIsReady{false};

// Sticky Keys for Alt-Tab cycling
bool g_isAltHeld = false;
const UINT_PTR TIMER_ID_ALT_RELEASE = 1001;

bool ParseAction(const std::string &rawValue, Action &action);
bool SaveConfig(const std::string &path, const Config &cfg);
std::string ActionToConfigValue(const Action &action);
bool ExecuteMacro(const std::string &payload);

constexpr UINT WM_TRAYICON = WM_APP + 1;
constexpr UINT ID_TRAY_SETTINGS = 1001;
constexpr UINT ID_TRAY_RELOAD = 1002;
constexpr UINT ID_TRAY_EXIT = 1003;
constexpr UINT ID_TRAY_UI_REMAPPER = 1004;
constexpr UINT ID_TRAY_UI_PROFILES = 1005;
constexpr UINT ID_TRAY_UI_SETTINGS = 1006;

constexpr int IDC_B4_TYPE = 2001;
constexpr int IDC_B4_VALUE = 2002;
constexpr int IDC_B4_BROWSE = 2003;
constexpr int IDC_B5_TYPE = 2011;
constexpr int IDC_B5_VALUE = 2012;
constexpr int IDC_B5_BROWSE = 2013;
constexpr int IDC_FULLSCREEN = 2021;
constexpr int IDC_HELP_TEXT = 2022;
constexpr int IDC_SAVE = 2031;
constexpr int IDC_CANCEL = 2032;

constexpr COLORREF UI_BG = RGB(12, 22, 36);
constexpr COLORREF UI_CARD = RGB(17, 33, 54);
constexpr COLORREF UI_INPUT = RGB(24, 42, 68);
constexpr COLORREF UI_TEXT = RGB(232, 242, 255);
constexpr COLORREF UI_TEXT_DIM = RGB(154, 178, 206);

HFONT g_fontUi = nullptr;
HFONT g_fontUiBold = nullptr;
HFONT g_fontUiTitle = nullptr;
HFONT g_fontUiMono = nullptr;
HBRUSH g_brushBg = nullptr;
HBRUSH g_brushCard = nullptr;
HBRUSH g_brushInput = nullptr;

// Removed UI resources for Win32 settings window

void SetControlFont(HWND hwnd, HFONT font) {
  SendMessageA(hwnd, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
}

std::string Trim(const std::string &s) {
  const auto start = s.find_first_not_of(" \t\r\n");
  if (start == std::string::npos) {
    return "";
  }

  const auto end = s.find_last_not_of(" \t\r\n");
  return s.substr(start, end - start + 1);
}

std::string ToUpper(std::string s) {
  std::transform(s.begin(), s.end(), s.begin(), [](unsigned char c) {
    return static_cast<char>(std::toupper(c));
  });
  return s;
}

std::vector<std::string> Split(const std::string &s, char delim) {
  std::vector<std::string> out;
  std::stringstream ss(s);
  std::string part;
  while (std::getline(ss, part, delim)) {
    out.push_back(part);
  }
  return out;
}

std::wstring Utf8ToWide(const std::string &input) {
  if (input.empty()) {
    return L"";
  }

  int len = MultiByteToWideChar(CP_UTF8, 0, input.c_str(), -1, nullptr, 0);
  if (len <= 0) {
    return L"";
  }

  std::wstring wide(static_cast<size_t>(len), L'\0');
  if (MultiByteToWideChar(CP_UTF8, 0, input.c_str(), -1, wide.data(), len) !=
      len) {
    return L"";
  }

  if (!wide.empty() && wide.back() == L'\0') {
    wide.pop_back();
  }

  return wide;
}

WORD KeyNameToVk(const std::string &name) {
  static const std::unordered_map<std::string, WORD> keyMap = {
      {"CTRL", VK_CONTROL},
      {"CONTROL", VK_CONTROL},
      {"SHIFT", VK_SHIFT},
      {"ALT", VK_MENU},
      {"WIN", VK_LWIN},
      {"WINDOWS", VK_LWIN},
      {"TAB", VK_TAB},
      {"ENTER", VK_RETURN},
      {"RETURN", VK_RETURN},
      {"ESC", VK_ESCAPE},
      {"ESCAPE", VK_ESCAPE},
      {"SPACE", VK_SPACE},
      {"BACKSPACE", VK_BACK},
      {"DELETE", VK_DELETE},
      {"DEL", VK_DELETE},
      {"INSERT", VK_INSERT},
      {"INS", VK_INSERT},
      {"HOME", VK_HOME},
      {"END", VK_END},
      {"PGUP", VK_PRIOR},
      {"PAGEUP", VK_PRIOR},
      {"PGDN", VK_NEXT},
      {"PAGEDOWN", VK_NEXT},
      {"UP", VK_UP},
      {"DOWN", VK_DOWN},
      {"LEFT", VK_LEFT},
      {"RIGHT", VK_RIGHT},
      {"CAPSLOCK", VK_CAPITAL},
      {"PRINTSCREEN", VK_SNAPSHOT},
      {"PRTSC", VK_SNAPSHOT},
      {"VOLUMEUP", VK_VOLUME_UP},
      {"VOLUMEDOWN", VK_VOLUME_DOWN},
      {"VOLUMEMUTE", VK_VOLUME_MUTE},
      {"PLAYPAUSE", VK_MEDIA_PLAY_PAUSE},
      {"NEXTTRACK", VK_MEDIA_NEXT_TRACK},
      {"PREVTRACK", VK_MEDIA_PREV_TRACK},
      {"BROWSERBACK", VK_BROWSER_BACK},
      {"BROWSERFORWARD", VK_BROWSER_FORWARD}};

  std::string upper = ToUpper(Trim(name));
  if (upper.empty()) {
    return 0;
  }

  auto it = keyMap.find(upper);
  if (it != keyMap.end()) {
    return it->second;
  }

  if (upper.size() == 1) {
    char c = upper[0];
    if ((c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9')) {
      return static_cast<WORD>(c);
    }
  }

  if (upper.size() >= 2 && upper[0] == 'F') {
    int fn = std::atoi(upper.c_str() + 1);
    if (fn >= 1 && fn <= 24) {
      return static_cast<WORD>(VK_F1 + (fn - 1));
    }
  }

  return 0;
}

bool ParseKeys(const std::string &value, std::vector<WORD> &outKeys) {
  outKeys.clear();
  auto parts = Split(value, '+');
  for (const std::string &raw : parts) {
    WORD vk = KeyNameToVk(raw);
    if (vk == 0) {
      return false;
    }
    outKeys.push_back(vk);
  }
  return !outKeys.empty();
}

bool SendKeyCombo(const std::vector<WORD> &keys) {
  if (keys.empty()) {
    return false;
  }

  std::vector<INPUT> inputs;
  inputs.reserve(keys.size() * 2);

  for (WORD vk : keys) {
    INPUT in = {};
    in.type = INPUT_KEYBOARD;
    in.ki.wVk = vk;
    inputs.push_back(in);
  }

  for (auto it = keys.rbegin(); it != keys.rend(); ++it) {
    INPUT in = {};
    in.type = INPUT_KEYBOARD;
    in.ki.wVk = *it;
    in.ki.dwFlags = KEYEVENTF_KEYUP;
    inputs.push_back(in);
  }

  return SendInput(static_cast<UINT>(inputs.size()), inputs.data(),
                   sizeof(INPUT)) == inputs.size();
}

void ReleaseStickyAlt() {
  if (g_isAltHeld) {
    INPUT in = {};
    in.type = INPUT_KEYBOARD;
    in.ki.wVk = VK_MENU;
    in.ki.dwFlags = KEYEVENTF_KEYUP;
    SendInput(1, &in, sizeof(INPUT));
    g_isAltHeld = false;
    if (g_mainWindow) {
      KillTimer(g_mainWindow, TIMER_ID_ALT_RELEASE);
    }
  }
}

bool HandleAltTab() {
  // Use Alt+Esc for "Infinite Cycling"
  // This cycles through all apps without opening the switcher UI
  // and doesn't get stuck toggling between just 2 apps.
  INPUT inputs[4] = {};
  
  // Alt Down
  inputs[0].type = INPUT_KEYBOARD;
  inputs[0].ki.wVk = VK_MENU;
  
  // Esc Down
  inputs[1].type = INPUT_KEYBOARD;
  inputs[1].ki.wVk = VK_ESCAPE;
  
  // Esc Up
  inputs[2].type = INPUT_KEYBOARD;
  inputs[2].ki.wVk = VK_ESCAPE;
  inputs[2].ki.dwFlags = KEYEVENTF_KEYUP;
  
  // Alt Up
  inputs[3].type = INPUT_KEYBOARD;
  inputs[3].ki.wVk = VK_MENU;
  inputs[3].ki.dwFlags = KEYEVENTF_KEYUP;

  SendInput(4, inputs, sizeof(INPUT));
  return true;
}

bool IsAltTabCombo(const std::vector<WORD> &keys) {
  if (keys.size() != 2)
    return false;
  bool hAlt = false, hTab = false;
  for (WORD k : keys) {
    if (k == VK_MENU || k == VK_LMENU || k == VK_RMENU)
      hAlt = true;
    if (k == VK_TAB)
      hTab = true;
  }
  return hAlt && hTab;
}

bool SendUnicodeText(const std::string &text) {
  std::wstring wide = Utf8ToWide(text);
  if (wide.empty()) {
    return false;
  }

  std::vector<INPUT> inputs;
  inputs.reserve(wide.size() * 2);

  for (wchar_t ch : wide) {
    INPUT down = {};
    down.type = INPUT_KEYBOARD;
    down.ki.dwFlags = KEYEVENTF_UNICODE;
    down.ki.wScan = ch;

    INPUT up = down;
    up.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;

    inputs.push_back(down);
    inputs.push_back(up);
  }

  return SendInput(static_cast<UINT>(inputs.size()), inputs.data(),
                   sizeof(INPUT)) == inputs.size();
}

bool RunCommand(const std::string &commandLine) {
  std::wstring cmd = Utf8ToWide(commandLine);
  if (cmd.empty()) {
    return false;
  }

  std::vector<wchar_t> mutableCmd(cmd.begin(), cmd.end());
  mutableCmd.push_back(L'\0');

  STARTUPINFOW si = {};
  si.cb = sizeof(si);
  PROCESS_INFORMATION pi = {};

  BOOL ok = CreateProcessW(nullptr, mutableCmd.data(), nullptr, nullptr, FALSE,
                           0, nullptr, nullptr, &si, &pi);

  if (!ok) {
    return false;
  }

  CloseHandle(pi.hThread);
  CloseHandle(pi.hProcess);
  return true;
}

bool OpenTarget(const std::string &target) {
  std::wstring wideTarget = Utf8ToWide(target);
  if (wideTarget.empty()) {
    return false;
  }

  HINSTANCE res = ShellExecuteW(nullptr, L"open", wideTarget.c_str(), nullptr,
                                nullptr, SW_SHOWNORMAL);
  return reinterpret_cast<INT_PTR>(res) > 32;
}

void SetLaunchOnStartup(bool enable) {
  HKEY hKey;
  const char *runKey = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";
  const char *appName = "NexusMouseRemapper";

  if (RegOpenKeyExA(HKEY_CURRENT_USER, runKey, 0, KEY_SET_VALUE, &hKey) ==
      ERROR_SUCCESS) {
    if (enable) {
      char path[MAX_PATH];
      GetModuleFileNameA(nullptr, path, MAX_PATH);
      
      // Enclose path in quotes for Windows to handle spaces correctly
      char quotedPath[MAX_PATH + 3];
      sprintf(quotedPath, "\"%s\"", path);

      RegSetValueExA(hKey, appName, 0, REG_SZ,
                     reinterpret_cast<const BYTE *>(quotedPath),
                     static_cast<DWORD>(strlen(quotedPath) + 1));
    } else {
      RegDeleteValueA(hKey, appName);
    }
    RegCloseKey(hKey);
  }
}

bool IsLaunchOnStartupEnabled() {
  HKEY hKey;
  const char *runKey = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";
  const char *appName = "NexusMouseRemapper";
  bool enabled = false;

  if (RegOpenKeyExA(HKEY_CURRENT_USER, runKey, 0, KEY_QUERY_VALUE, &hKey) ==
      ERROR_SUCCESS) {
    if (RegQueryValueExA(hKey, appName, nullptr, nullptr, nullptr, nullptr) ==
        ERROR_SUCCESS) {
      enabled = true;
    }
    RegCloseKey(hKey);
  }
  return enabled;
}

bool ExecuteMacro(const std::string &payload) {
  // Format: VK_CODE:STATE,DELAY,VK_CODE:STATE...
  // States: D (Down), U (Up), P (Press/Both)
  auto parts = Split(payload, ',');
  for (const auto &p : parts) {
    if (p.empty()) continue;
    if (std::isdigit(p[0])) {
      Sleep(std::atoi(p.c_str()));
    } else {
      auto sub = Split(p, ':');
      if (sub.size() < 1) continue;
      WORD vk = KeyNameToVk(sub[0]);
      if (vk == 0) continue;
      char state = (sub.size() > 1) ? ToUpper(sub[1])[0] : 'P';
      
      INPUT in = {};
      in.type = INPUT_KEYBOARD;
      in.ki.wVk = vk;
      if (state == 'U') {
        in.ki.dwFlags = KEYEVENTF_KEYUP;
        SendInput(1, &in, sizeof(INPUT));
      } else if (state == 'D') {
        SendInput(1, &in, sizeof(INPUT));
      } else {
        // Press (Down + Up)
        SendInput(1, &in, sizeof(INPUT));
        Sleep(10);
        in.ki.dwFlags = KEYEVENTF_KEYUP;
        SendInput(1, &in, sizeof(INPUT));
      }
    }
  }
  return true;
}

bool ExecuteAction(const Action &action) {
  // If doing non-alt-tab action, release Alt if it was stuck
  if (g_isAltHeld &&
      (action.type != ActionType::Keys || !IsAltTabCombo(action.keys))) {
    ReleaseStickyAlt();
  }

  switch (action.type) {
  case ActionType::Keys:
    if (IsAltTabCombo(action.keys)) {
      return HandleAltTab();
    }
    return SendKeyCombo(action.keys);
  case ActionType::Run:
    return RunCommand(action.payload);
  case ActionType::Open:
    return OpenTarget(action.payload);
  case ActionType::Text:
    return SendUnicodeText(action.payload);
  case ActionType::Macro:
    return ExecuteMacro(action.payload);
  case ActionType::None:
  default:
    return false;
  }
}

std::string GetConfigPath() {
  char buffer[MAX_PATH] = {};
  GetModuleFileNameA(nullptr, buffer, MAX_PATH);
  std::string exePath(buffer);

  const auto slash = exePath.find_last_of("\\/");
  const std::string dir =
      (slash == std::string::npos) ? "." : exePath.substr(0, slash);
  return dir + "\\mouse_remap.ini";
}

std::string GetExeDir() {
  char buffer[MAX_PATH] = {};
  GetModuleFileNameA(nullptr, buffer, MAX_PATH);
  std::string exePath(buffer);
  const auto slash = exePath.find_last_of("\\/");
  return (slash == std::string::npos) ? "." : exePath.substr(0, slash);
}

std::string ToFileUrl(const std::string &path) {
  std::string out = "file:///";
  for (char c : path) {
    if (c == '\\') {
      out += '/';
    } else if (c == ' ') {
      out += "%20";
    } else {
      out += c;
    }
  }
  return out;
}

std::string GetEdgePath() {
  char path[MAX_PATH] = {};
  DWORD size = sizeof(path);
  if (RegGetValueA(
          HKEY_LOCAL_MACHINE,
          "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\msedge.exe",
          nullptr, RRF_RT_REG_SZ, nullptr, path, &size) == ERROR_SUCCESS) {
    return std::string(path);
  }
  return "msedge.exe";
}

void OpenStitchPage(const char *page) {
  const std::string path = GetExeDir() + "\\ui\\" + page;
  const std::string fileUrl = ToFileUrl(path);
  const std::string edge = GetEdgePath();
  const std::string args = "--app=\"" + fileUrl + "\" --window-size=440,900";

  HINSTANCE res = ShellExecuteA(nullptr, "open", edge.c_str(), args.c_str(),
                                nullptr, SW_SHOWNORMAL);
  if (reinterpret_cast<INT_PTR>(res) <= 32) {
    // Fallback to default browser
    ShellExecuteA(nullptr, "open", fileUrl.c_str(), nullptr, nullptr,
                  SW_SHOWNORMAL);
  }
}

std::string JsonEscape(const std::string &s) {
  std::string out;
  out.reserve(s.size() + 16);
  for (char c : s) {
    switch (c) {
    case '\\':
      out += "\\\\";
      break;
    case '"':
      out += "\\\"";
      break;
    case '\n':
      out += "\\n";
      break;
    case '\r':
      out += "\\r";
      break;
    case '\t':
      out += "\\t";
      break;
    default:
      out += c;
      break;
    }
  }
  return out;
}

void UpdatePollingRateWindow() {
  const unsigned long long now = GetTickCount64();
  const unsigned long long lastTick = g_lastRateWindowTick.load();
  const unsigned long long count = g_rawInputEvents.load();

  if (lastTick == 0) {
    g_lastRateWindowTick.store(now);
    g_lastRateWindowCount.store(count);
    return;
  }

  if (now - lastTick >= 1000) {
    const unsigned long long prevCount = g_lastRateWindowCount.load();
    const unsigned long long delta =
        (count >= prevCount) ? (count - prevCount) : 0;
    g_pollRateHz.store(static_cast<unsigned int>(delta));
    g_lastRateWindowTick.store(now);
    g_lastRateWindowCount.store(count);
  }
}

void UpdateMouseDeviceInfo(HANDLE hDevice) {
  if (!hDevice) {
    return;
  }

  RID_DEVICE_INFO info = {};
  info.cbSize = sizeof(info);
  UINT size = sizeof(info);
  if (GetRawInputDeviceInfoA(hDevice, RIDI_DEVICEINFO, &info, &size) ==
      static_cast<UINT>(-1)) {
    return;
  }

  if (info.dwType == RIM_TYPEMOUSE) {
    g_mouseButtons.store(info.mouse.dwNumberOfButtons);
    g_mouseSampleRate.store(info.mouse.dwSampleRate);
  }
}

void UpdateTelemetryFromRawInput(const RAWINPUT *raw) {
  if (!raw || raw->header.dwType != RIM_TYPEMOUSE) {
    return;
  }

  g_rawInputEvents.fetch_add(1);
  UpdatePollingRateWindow();
  UpdateMouseDeviceInfo(raw->header.hDevice);
}

std::string BuildStatusJson() {
  std::ostringstream ss;
  ss << "{";
  ss << "\"poll_rate_hz\":" << g_pollRateHz.load() << ",";
  ss << "\"mouse_buttons\":" << g_mouseButtons.load() << ",";
  ss << "\"status\":\"active\",";
  ss << "\"config_path\":\"" << JsonEscape(g_configPath) << "\"";
  ss << "}";
  return ss.str();
}

std::string BuildConfigJson() {
  std::ostringstream ss;
  ss << "{";
  ss << "\"button4\":\"" << JsonEscape(ActionToConfigValue(g_config.button4))
     << "\",";
  ss << "\"button5\":\"" << JsonEscape(ActionToConfigValue(g_config.button5))
     << "\",";
  ss << "\"suspend_fullscreen\":"
     << (g_config.suspendInFullscreen ? "true" : "false") << ",";
  ss << "\"dpi\":" << g_config.dpi << ",";
  ss << "\"launch_on_startup\":"
     << (IsLaunchOnStartupEnabled() ? "true" : "false");
  ss << "}";
  return ss.str();
}

std::string JsonUnescape(const std::string &s) {
  std::string out;
  out.reserve(s.size());
  bool esc = false;
  for (char c : s) {
    if (!esc) {
      if (c == '\\') {
        esc = true;
      } else {
        out.push_back(c);
      }
      continue;
    }

    switch (c) {
    case '\\':
      out.push_back('\\');
      break;
    case '"':
      out.push_back('"');
      break;
    case 'n':
      out.push_back('\n');
      break;
    case 'r':
      out.push_back('\r');
      break;
    case 't':
      out.push_back('\t');
      break;
    default:
      out.push_back(c);
      break;
    }
    esc = false;
  }
  return out;
}

bool ExtractJsonString(const std::string &body, const std::string &key,
                       std::string &out) {
  const std::string needle = "\"" + key + "\"";
  const size_t k = body.find(needle);
  if (k == std::string::npos) {
    return false;
  }
  size_t colon = body.find(':', k + needle.size());
  if (colon == std::string::npos) {
    return false;
  }
  size_t q1 = body.find('"', colon + 1);
  if (q1 == std::string::npos) {
    return false;
  }
  size_t i = q1 + 1;
  bool esc = false;
  for (; i < body.size(); ++i) {
    char c = body[i];
    if (esc) {
      esc = false;
      continue;
    }
    if (c == '\\') {
      esc = true;
      continue;
    }
    if (c == '"') {
      break;
    }
  }
  if (i >= body.size()) {
    return false;
  }
  out = JsonUnescape(body.substr(q1 + 1, i - q1 - 1));
  return true;
}

bool ExtractJsonBool(const std::string &body, const std::string &key,
                     bool &out) {
  const std::string needle = "\"" + key + "\"";
  const size_t k = body.find(needle);
  if (k == std::string::npos) {
    return false;
  }
  size_t colon = body.find(':', k + needle.size());
  if (colon == std::string::npos) {
    return false;
  }
  size_t v = body.find_first_not_of(" \t\r\n", colon + 1);
  if (v == std::string::npos) {
    return false;
  }
  if (body.compare(v, 4, "true") == 0) {
    out = true;
    return true;
  }
  if (body.compare(v, 5, "false") == 0) {
    out = false;
    return true;
  }
  return false;
}

bool ExtractJsonInt(const std::string &body, const std::string &key, int &out) {
  const std::string needle = "\"" + key + "\"";
  const size_t k = body.find(needle);
  if (k == std::string::npos) {
    return false;
  }
  size_t colon = body.find(':', k + needle.size());
  if (colon == std::string::npos) {
    return false;
  }
  size_t v = body.find_first_not_of(" \t\r\n", colon + 1);
  if (v == std::string::npos) {
    return false;
  }
  out = std::atoi(body.c_str() + v);
  return true;
}

bool ApplyConfigJson(const std::string &body, std::string &error) {
  std::string b4;
  std::string b5;
  bool fullscreen = g_config.suspendInFullscreen;

  if (!ExtractJsonString(body, "button4", b4) ||
      !ExtractJsonString(body, "button5", b5)) {
    error = "missing button mapping";
    return false;
  }
  (void)ExtractJsonBool(body, "suspend_fullscreen", fullscreen);

  int dpi = g_config.dpi;
  (void)ExtractJsonInt(body, "dpi", dpi);

  Action a4;
  Action a5;
  if (!ParseAction(b4, a4) || !ParseAction(b5, a5)) {
    error = "invalid action syntax";
    return false;
  }

  Config next = g_config;
  next.button4 = a4;
  next.button5 = a5;
  next.suspendInFullscreen = fullscreen;
  next.dpi = dpi;

  bool startup = false;
  if (ExtractJsonBool(body, "launch_on_startup", startup)) {
    SetLaunchOnStartup(startup);
  }

  if (!SaveConfig(g_configPath, next)) {
    error = "failed to save config";
    return false;
  }

  g_config = next;
  return true;
}

void HandleStatusClient(SOCKET client) {
  std::string r;
  r.reserve(4096);
  char chunk[2048] = {};
  int contentLength = -1;
  size_t headerEnd = std::string::npos;

  while (true) {
    const int got = recv(client, chunk, sizeof(chunk), 0);
    if (got <= 0) {
      break;
    }
    r.append(chunk, chunk + got);

    if (headerEnd == std::string::npos) {
      headerEnd = r.find("\r\n\r\n");
      if (headerEnd != std::string::npos) {
        const std::string h = r.substr(0, headerEnd);
        const std::string key = "Content-Length:";
        size_t p = h.find(key);
        if (p != std::string::npos) {
          p += key.size();
          while (p < h.size() && (h[p] == ' ' || h[p] == '\t')) {
            ++p;
          }
          contentLength = std::atoi(h.c_str() + p);
        } else {
          contentLength = 0;
        }
      }
    }

    if (headerEnd != std::string::npos && contentLength >= 0) {
      const size_t haveBody = r.size() - (headerEnd + 4);
      if (haveBody >= static_cast<size_t>(contentLength)) {
        break;
      }
    }
  }

  if (r.empty()) {
    return;
  }
  std::string body;
  std::string code = "200 OK";
  std::string type = "application/json";

  if (r.rfind("GET /status", 0) == 0) {
    body = BuildStatusJson();
  } else if (r.rfind("GET /config", 0) == 0) {
    body = BuildConfigJson();
  } else if (r.rfind("POST /config", 0) == 0) {
    const size_t sep = r.find("\r\n\r\n");
    std::string reqBody = (sep == std::string::npos) ? "" : r.substr(sep + 4);
    std::string err;
    if (ApplyConfigJson(reqBody, err)) {
      body = "{\"ok\":true}";
    } else {
      code = "400 Bad Request";
      body = "{\"ok\":false,\"error\":\"" + JsonEscape(err) + "\"}";
    }
  } else if (r.rfind("GET /browse", 0) == 0) {
    char path[MAX_PATH] = {};
    OPENFILENAMEA ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = nullptr;
    ofn.lpstrFile = path;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrFilter = "Applications (*.exe)\0*.exe\0All Files (*.*)\0*.*\0";
    ofn.nFilterIndex = 1;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_NOCHANGEDIR;

    if (GetOpenFileNameA(&ofn)) {
      body = "{\"ok\":true, \"path\":\"" + JsonEscape(path) + "\"}";
    } else {
      body = "{\"ok\":false}";
    }
  } else if (r.rfind("OPTIONS ", 0) == 0) {
    body = "";
    type = "text/plain";
  } else {
    code = "404 Not Found";
    type = "application/json";
    body = "{\"error\":\"not_found\"}";
  }

  std::ostringstream resp;
  resp << "HTTP/1.1 " << code << "\r\n";
  resp << "Content-Type: " << type << "\r\n";
  resp << "Access-Control-Allow-Origin: *\r\n";
  resp << "Access-Control-Allow-Methods: GET, POST, OPTIONS\r\n";
  resp << "Access-Control-Allow-Headers: *\r\n";
  resp << "Connection: close\r\n";
  resp << "Content-Length: " << body.size() << "\r\n\r\n";
  resp << body;
  const std::string out = resp.str();
  send(client, out.c_str(), static_cast<int>(out.size()), 0);
}

void StatusServerThreadProc() {
  WSADATA wsa = {};
  if (WSAStartup(MAKEWORD(2, 2), &wsa) != 0) {
    return;
  }

  SOCKET listenSock = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
  if (listenSock == INVALID_SOCKET) {
    WSACleanup();
    return;
  }
  g_statusListenSocket.store(listenSock);

  sockaddr_in addr = {};
  addr.sin_family = AF_INET;
  addr.sin_port = htons(48621);
  inet_pton(AF_INET, "127.0.0.1", &addr.sin_addr);

  int yes = 1;
  setsockopt(listenSock, SOL_SOCKET, SO_REUSEADDR,
             reinterpret_cast<const char *>(&yes), sizeof(yes));
  if (bind(listenSock, reinterpret_cast<sockaddr *>(&addr), sizeof(addr)) ==
          SOCKET_ERROR ||
      listen(listenSock, 4) == SOCKET_ERROR) {
    closesocket(listenSock);
    g_statusListenSocket.store(INVALID_SOCKET);
    WSACleanup();
    return;
  }

  while (!g_statusServerStop.load()) {
    fd_set set;
    FD_ZERO(&set);
    FD_SET(listenSock, &set);
    timeval tv = {0, 250000};
    const int sel = select(0, &set, nullptr, nullptr, &tv);
    if (sel <= 0) {
      continue;
    }

    SOCKET client = accept(listenSock, nullptr, nullptr);
    if (client == INVALID_SOCKET) {
      continue;
    }
    HandleStatusClient(client);
    closesocket(client);
  }

  closesocket(listenSock);
  g_statusListenSocket.store(INVALID_SOCKET);
  WSACleanup();
}

void StartStatusServer() {
  g_statusServerStop.store(false);
  g_statusServerThread = std::thread(StatusServerThreadProc);
}

void StopStatusServer() {
  g_statusServerStop.store(true);
  const SOCKET s = g_statusListenSocket.load();
  if (s != INVALID_SOCKET) {
    closesocket(s);
  }
  if (g_statusServerThread.joinable()) {
    g_statusServerThread.join();
  }
}

std::string ActionTypeToString(ActionType type) {
  switch (type) {
  case ActionType::Keys:
    return "keys";
  case ActionType::Run:
    return "run";
  case ActionType::Open:
    return "open";
  case ActionType::Text:
    return "text";
  case ActionType::Macro:
    return "macro";
  default:
    return "none";
  }
}

std::string KeysToString(const std::vector<WORD> &keys) {
  std::unordered_map<WORD, std::string> reverse = {
      {VK_CONTROL, "CTRL"},
      {VK_SHIFT, "SHIFT"},
      {VK_MENU, "ALT"},
      {VK_LWIN, "WIN"},
      {VK_TAB, "TAB"},
      {VK_RETURN, "ENTER"},
      {VK_ESCAPE, "ESC"},
      {VK_SPACE, "SPACE"},
      {VK_BACK, "BACKSPACE"},
      {VK_DELETE, "DELETE"},
      {VK_INSERT, "INSERT"},
      {VK_HOME, "HOME"},
      {VK_END, "END"},
      {VK_PRIOR, "PGUP"},
      {VK_NEXT, "PGDN"},
      {VK_UP, "UP"},
      {VK_DOWN, "DOWN"},
      {VK_LEFT, "LEFT"},
      {VK_RIGHT, "RIGHT"},
      {VK_VOLUME_UP, "VOLUMEUP"},
      {VK_VOLUME_DOWN, "VOLUMEDOWN"},
      {VK_VOLUME_MUTE, "VOLUMEMUTE"},
      {VK_MEDIA_PLAY_PAUSE, "PLAYPAUSE"},
      {VK_MEDIA_NEXT_TRACK, "NEXTTRACK"},
      {VK_MEDIA_PREV_TRACK, "PREVTRACK"}};

  std::string out;
  for (size_t i = 0; i < keys.size(); ++i) {
    WORD vk = keys[i];
    std::string part;
    auto it = reverse.find(vk);
    if (it != reverse.end()) {
      part = it->second;
    } else if ((vk >= 'A' && vk <= 'Z') || (vk >= '0' && vk <= '9')) {
      part.push_back(static_cast<char>(vk));
    } else if (vk >= VK_F1 && vk <= VK_F24) {
      part = "F" + std::to_string(vk - VK_F1 + 1);
    } else {
      part = std::to_string(vk);
    }

    if (!out.empty()) {
      out += "+";
    }
    out += part;
  }

  return out;
}

std::string ActionToConfigValue(const Action &action) {
  if (action.type == ActionType::None) {
    return "none:";
  }

  if (action.type == ActionType::Keys) {
    return "keys:" + KeysToString(action.keys);
  }

  return ActionTypeToString(action.type) + ":" + action.payload;
}

bool ParseAction(const std::string &rawValue, Action &action) {
  action = {};
  std::string value = Trim(rawValue);
  if (value.empty()) {
    return false;
  }

  const auto colon = value.find(':');
  if (colon == std::string::npos) {
    return false;
  }

  std::string type = ToUpper(Trim(value.substr(0, colon)));
  std::string payload = Trim(value.substr(colon + 1));

  if (type == "NONE") {
    action.type = ActionType::None;
    return true;
  }

  if (type == "KEYS") {
    std::vector<WORD> keys;
    if (!ParseKeys(payload, keys)) {
      return false;
    }
    action.type = ActionType::Keys;
    action.keys = std::move(keys);
    return true;
  }

  if (type == "RUN") {
    action.type = ActionType::Run;
    action.payload = payload;
    return !action.payload.empty();
  }

  if (type == "OPEN") {
    action.type = ActionType::Open;
    action.payload = payload;
    return !action.payload.empty();
  }

  if (type == "TEXT") {
    action.type = ActionType::Text;
    action.payload = payload;
    return !action.payload.empty();
  }

  if (type == "MACRO") {
    action.type = ActionType::Macro;
    action.payload = payload;
    return !action.payload.empty();
  }

  return false;
}

void WriteDefaultConfigIfMissing(const std::string &path) {
  std::ifstream probe(path);
  if (probe.good()) {
    return;
  }

  std::ofstream out(path);
  out << "# Mouse side button remap config\n";
  out << "# button4 / button5 syntax: <type>:<value>\n";
  out << "# types: none, keys, run, open, text\n";
  out << "# suspend_fullscreen=true disables remap when a fullscreen window is "
         "active\n\n";
  out << "button4=keys:CTRL+C\n";
  out << "button5=keys:ALT+TAB\n";
  out << "suspend_fullscreen=true\n";
}

Config LoadConfig(const std::string &path) {
  Config cfg;

  std::ifstream in(path);
  if (!in.is_open()) {
    return cfg;
  }

  std::string line;
  while (std::getline(in, line)) {
    std::string t = Trim(line);
    if (t.empty() || t[0] == '#') {
      continue;
    }

    const auto eq = t.find('=');
    if (eq == std::string::npos) {
      continue;
    }

    std::string key = ToUpper(Trim(t.substr(0, eq)));
    std::string value = Trim(t.substr(eq + 1));

    if (key == "BUTTON4") {
      Action action;
      if (ParseAction(value, action)) {
        cfg.button4 = action;
      }
    } else if (key == "BUTTON5") {
      Action action;
      if (ParseAction(value, action)) {
        cfg.button5 = action;
      }
    } else if (key == "SUSPEND_FULLSCREEN") {
      std::string b = ToUpper(value);
      cfg.suspendInFullscreen =
          (b == "1" || b == "TRUE" || b == "YES" || b == "ON");
    } else if (key == "DPI") {
      cfg.dpi = std::atoi(value.c_str());
    }
  }

  return cfg;
}

bool SaveConfig(const std::string &path, const Config &cfg) {
  std::ofstream out(path, std::ios::trunc);
  if (!out.is_open()) {
    return false;
  }

  out << "# Mouse side button remap config\n";
  out << "button4=" << ActionToConfigValue(cfg.button4) << "\n";
  out << "button5=" << ActionToConfigValue(cfg.button5) << "\n";
  out << "suspend_fullscreen=" << (cfg.suspendInFullscreen ? "true" : "false")
      << "\n";
  out << "dpi=" << cfg.dpi << "\n";
  return true;
}

bool IsFullscreenForegroundWindow() {
  HWND fg = GetForegroundWindow();
  if (!fg || fg == g_mainWindow || fg == g_settingsWindow) {
    return false;
  }

  RECT wr = {};
  if (!GetWindowRect(fg, &wr)) {
    return false;
  }

  HMONITOR monitor = MonitorFromWindow(fg, MONITOR_DEFAULTTONEAREST);
  MONITORINFO mi = {};
  mi.cbSize = sizeof(mi);
  if (!GetMonitorInfoA(monitor, &mi)) {
    return false;
  }

  const int tol = 2;
  return abs(wr.left - mi.rcMonitor.left) <= tol &&
         abs(wr.top - mi.rcMonitor.top) <= tol &&
         abs(wr.right - mi.rcMonitor.right) <= tol &&
         abs(wr.bottom - mi.rcMonitor.bottom) <= tol;
}

void HandleMouse(USHORT flags) {
  static DWORD lastBtn4Tick = 0;
  static DWORD lastBtn5Tick = 0;
  const DWORD now = GetTickCount();
  const DWORD debounceMs = 120;

  if (g_config.suspendInFullscreen && IsFullscreenForegroundWindow()) {
    return;
  }

  if ((flags & RI_MOUSE_BUTTON_4_DOWN) && (now - lastBtn4Tick > debounceMs)) {
    lastBtn4Tick = now;
    ExecuteAction(g_config.button4);
  }

  if ((flags & RI_MOUSE_BUTTON_5_DOWN) && (now - lastBtn5Tick > debounceMs)) {
    lastBtn5Tick = now;
    ExecuteAction(g_config.button5);
  }
}

ActionType IndexToActionType(int idx) {
  switch (idx) {
  case 1:
    return ActionType::Keys;
  case 2:
    return ActionType::Run;
  case 3:
    return ActionType::Open;
  case 4:
    return ActionType::Text;
  case 5:
    return ActionType::Macro;
  default:
    return ActionType::None;
  }
}

int ActionTypeToIndex(ActionType type) {
  switch (type) {
  case ActionType::Keys:
    return 1;
  case ActionType::Run:
    return 2;
  case ActionType::Open:
    return 3;
  case ActionType::Text:
    return 4;
  default:
    return 0;
  }
}

void SetWindowTextFromString(HWND hwnd, const std::string &text) {
  SetWindowTextA(hwnd, text.c_str());
}

std::string GetWindowTextToString(HWND hwnd) {
  int len = GetWindowTextLengthA(hwnd);
  std::string s(static_cast<size_t>(len), '\0');
  if (len > 0) {
    GetWindowTextA(hwnd, s.data(), len + 1);
  }
  return s;
}

std::string BuildValueForEditor(const Action &action) {
  if (action.type == ActionType::Keys) {
    return KeysToString(action.keys);
  }
  return action.payload;
}

Action BuildActionFromEditor(ActionType type, const std::string &value) {
  Action action;
  action.type = type;

  if (type == ActionType::Keys) {
    ParseKeys(value, action.keys);
  } else if (type == ActionType::Run || type == ActionType::Open ||
             type == ActionType::Text) {
    action.payload = value;
  }

  return action;
}

// Removed settings controls logic

void ShowSettingsWindow();

void ShowTrayMenu(HWND hwnd) {
  HMENU menu = CreatePopupMenu();
  AppendMenuA(menu, MF_STRING, ID_TRAY_UI_REMAPPER, "Open Remapper");
  AppendMenuA(menu, MF_STRING, ID_TRAY_UI_PROFILES, "Profiles");
  AppendMenuA(menu, MF_STRING, ID_TRAY_UI_SETTINGS, "Settings");
  AppendMenuA(menu, MF_SEPARATOR, 0, nullptr);
  AppendMenuA(menu, MF_STRING, ID_TRAY_RELOAD, "Reload Config");
  AppendMenuA(menu, MF_STRING, ID_TRAY_EXIT, "Exit");

  POINT pt;
  GetCursorPos(&pt);
  SetForegroundWindow(hwnd);
  TrackPopupMenu(menu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, nullptr);
  DestroyMenu(menu);
}

// Win32 settings window removed in favor of web UI

bool AddTrayIcon(HWND hwnd) {
  g_tray.cbSize = sizeof(g_tray);
  g_tray.hWnd = hwnd;
  g_tray.uID = 1;
  g_tray.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
  g_tray.uCallbackMessage = WM_TRAYICON;
  g_tray.hIcon = LoadIcon(nullptr, IDI_APPLICATION);
  lstrcpynA(g_tray.szTip, "Mouse Remapper", sizeof(g_tray.szTip));
  return Shell_NotifyIconA(NIM_ADD, &g_tray) == TRUE;
}

void RemoveTrayIcon() { Shell_NotifyIconA(NIM_DELETE, &g_tray); }

LRESULT CALLBACK MainProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
  switch (msg) {
  case WM_CREATE: {
    RAWINPUTDEVICE rid = {};
    rid.usUsagePage = 0x01;
    rid.usUsage = 0x02;
    rid.dwFlags = RIDEV_INPUTSINK;
    rid.hwndTarget = hwnd;
    if (!RegisterRawInputDevices(&rid, 1, sizeof(rid))) {
      MessageBoxA(hwnd, "Failed to register raw input.", "Mouse Remapper",
                  MB_ICONERROR | MB_OK);
      PostQuitMessage(1);
      return 0;
    }
    if (!AddTrayIcon(hwnd)) {
      MessageBoxA(hwnd, "Failed to create tray icon.", "Mouse Remapper",
                  MB_ICONWARNING | MB_OK);
    }
    return 0;
  }
  case WM_INPUT: {
    UINT size = 0;
    if (GetRawInputData(reinterpret_cast<HRAWINPUT>(lParam), RID_INPUT, nullptr,
                        &size, sizeof(RAWINPUTHEADER)) != 0) {
      return 0;
    }

    std::vector<BYTE> buffer(size);
    if (GetRawInputData(reinterpret_cast<HRAWINPUT>(lParam), RID_INPUT,
                        buffer.data(), &size, sizeof(RAWINPUTHEADER)) != size) {
      return 0;
    }

    const RAWINPUT *raw = reinterpret_cast<const RAWINPUT *>(buffer.data());
    if (raw->header.dwType == RIM_TYPEMOUSE) {
      UpdateTelemetryFromRawInput(raw);
      HandleMouse(raw->data.mouse.usButtonFlags);
    }
    return 0;
  }
  case WM_TRAYICON:
    if (LOWORD(lParam) == WM_RBUTTONUP || LOWORD(lParam) == WM_CONTEXTMENU) {
      ShowTrayMenu(hwnd);
    } else if (LOWORD(lParam) == WM_LBUTTONDBLCLK) {
      OpenStitchPage("remapper.html");
    }
    return 0;
  case WM_COMMAND:
    switch (LOWORD(wParam)) {
    case ID_TRAY_UI_REMAPPER:
      OpenStitchPage("remapper.html");
      return 0;
    case ID_TRAY_UI_PROFILES:
      OpenStitchPage("profiles.html");
      return 0;
    case ID_TRAY_UI_SETTINGS:
      OpenStitchPage("settings.html");
      return 0;
    case ID_TRAY_SETTINGS:
      OpenStitchPage("settings.html");
      return 0;
    case ID_TRAY_RELOAD:
      g_config = LoadConfig(g_configPath);
      return 0;
    case ID_TRAY_EXIT:
      DestroyWindow(hwnd);
      return 0;
    default:
      return DefWindowProcA(hwnd, msg, wParam, lParam);
    }
  case WM_TIMER:
    if (wParam == TIMER_ID_ALT_RELEASE) {
      ReleaseStickyAlt();
    }
    return 0;
  case WM_DESTROY:
    if (g_settingsWindow) {
      DestroyWindow(g_settingsWindow);
      g_settingsWindow = nullptr;
    }
    StopStatusServer();
    RemoveTrayIcon();
    PostQuitMessage(0);
    return 0;
  default:
    return DefWindowProcA(hwnd, msg, wParam, lParam);
  }
}

} // namespace

int WINAPI WinMain(HINSTANCE instance, HINSTANCE, LPSTR, int) {
  g_configPath = GetConfigPath();
  WriteDefaultConfigIfMissing(g_configPath);
  g_config = LoadConfig(g_configPath);
  StartStatusServer();

  WNDCLASSA wc = {};
  wc.lpfnWndProc = MainProc;
  wc.hInstance = instance;
  wc.lpszClassName = "MouseRemapperMainClass";
  wc.hCursor = LoadCursor(nullptr, IDC_ARROW);

  if (!RegisterClassA(&wc)) {
    return 1;
  }

  g_mainWindow = CreateWindowExA(
      0, wc.lpszClassName, "Mouse Remapper Hidden Window", WS_OVERLAPPEDWINDOW,
      0, 0, 0, 0, nullptr, nullptr, instance, nullptr);

  if (!g_mainWindow) {
    MessageBoxA(nullptr, "Failed to create app window.", "Mouse Remapper",
                MB_ICONERROR | MB_OK);
    return 1;
  }

  OpenStitchPage("remapper.html");

  MSG msg = {};
  while (GetMessageA(&msg, nullptr, 0, 0) > 0) {
    TranslateMessage(&msg);
    DispatchMessageA(&msg);
  }

  return 0;
}
