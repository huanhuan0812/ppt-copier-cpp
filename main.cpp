#include <windows.h>
#include <dbt.h>
#include <shlobj.h>
#include <shellapi.h>
#include <psapi.h>
#include <iostream>
#include <string>
#include <vector>
#include <set>
#include <map>
#include <memory>
#include <thread>
#include <mutex>
#include <chrono>
#include <fstream>
#include <sstream>
#include <ctime>
#include <algorithm>
#include <filesystem>

#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "psapi.lib")

namespace fs = std::filesystem;

// 常量定义
const UINT WM_NOTIFY_CALLBACK = WM_USER + 100;
const int WM_TRAYICON = WM_USER + 101;
const UINT WM_DEVICECHANGE = 0x0219;
const DWORD DBT_DEVICEARRIVAL = 0x8000;
const DWORD DBT_DEVICEREMOVECOMPLETE = 0x8004;

// PPT扩展名
const std::set<std::wstring> PPT_EXTENSIONS = {
    L".ppt", L".pptx", L".pps", L".ppsx"
};

// 日志级别
enum LogLevel {
    LOG_DEBUG,
    LOG_INFO,
    LOG_WARNING,
    LOG_ERROR
};

// 日志管理器
class Logger {
private:
    std::wofstream logFile;
    std::mutex logMutex;
    std::wstring logDir;
    std::wstring logFilePath;
    LogLevel consoleLevel;
    LogLevel fileLevel;

    std::wstring GetCurrentTime() {
        auto now = std::chrono::system_clock::now();
        auto time_t_now = std::chrono::system_clock::to_time_t(now);
        struct tm tm_now;
        localtime_s(&tm_now, &time_t_now);
        
        wchar_t buffer[64];
        wcsftime(buffer, 64, L"%Y-%m-%d %H:%M:%S", &tm_now);
        return std::wstring(buffer);
    }

    std::wstring GetCurrentDate() {
        auto now = std::chrono::system_clock::now();
        auto time_t_now = std::chrono::system_clock::to_time_t(now);
        struct tm tm_now;
        localtime_s(&tm_now, &time_t_now);
        
        wchar_t buffer[32];
        wcsftime(buffer, 32, L"%Y%m%d", &tm_now);
        return std::wstring(buffer);
    }

    void RotateLogFile() {
        if (logFile.is_open()) {
            logFile.close();
        }
        
        std::wstring newLogPath = logDir + L"\\ppt_monitor_" + GetCurrentDate() + L".log";
        
        // 检查文件大小，如果超过10MB则轮转
        if (fs::exists(newLogPath) && fs::file_size(newLogPath) > 10 * 1024 * 1024) {
            for (int i = 4; i >= 1; --i) {
                std::wstring oldFile = logDir + L"\\ppt_monitor_" + GetCurrentDate() + L".log." + std::to_wstring(i);
                std::wstring newFile = logDir + L"\\ppt_monitor_" + GetCurrentDate() + L".log." + std::to_wstring(i + 1);
                if (fs::exists(oldFile)) {
                    fs::rename(oldFile, newFile);
                }
            }
            fs::rename(newLogPath, logDir + L"\\ppt_monitor_" + GetCurrentDate() + L".log.1");
        }
        
        logFile.open(newLogPath, std::ios::app);
        logFilePath = newLogPath;
    }

public:
    Logger() : consoleLevel(LOG_INFO), fileLevel(LOG_DEBUG) {
        // 获取程序目录
        wchar_t exePath[MAX_PATH];
        GetModuleFileNameW(NULL, exePath, MAX_PATH);
        std::wstring exeDir = fs::path(exePath).parent_path().wstring();
        
        logDir = exeDir + L"\\logs";
        
        try {
            if (!fs::exists(logDir)) {
                fs::create_directories(logDir);
            }
            RotateLogFile();
        }
        catch (...) {
            // 忽略日志目录创建失败
        }
    }

    ~Logger() {
        if (logFile.is_open()) {
            logFile.close();
        }
    }

    void Log(LogLevel level, const std::wstring& message) {
        std::lock_guard<std::mutex> lock(logMutex);
        
        std::wstring levelStr;
        switch (level) {
        case LOG_DEBUG: levelStr = L"DEBUG"; break;
        case LOG_INFO: levelStr = L"INFO"; break;
        case LOG_WARNING: levelStr = L"WARNING"; break;
        case LOG_ERROR: levelStr = L"ERROR"; break;
        }
        
        std::wstring logLine = GetCurrentTime() + L" - " + levelStr + L" - " + message;
        
        // 输出到控制台
        if (level >= consoleLevel) {
            std::wcout << logLine << std::endl;
        }
        
        // 输出到文件
        if (level >= fileLevel && logFile.is_open()) {
            logFile << logLine << std::endl;
            logFile.flush();
        }
    }

    void Debug(const std::wstring& msg) { Log(LOG_DEBUG, msg); }
    void Info(const std::wstring& msg) { Log(LOG_INFO, msg); }
    void Warning(const std::wstring& msg) { Log(LOG_WARNING, msg); }
    void Error(const std::wstring& msg) { Log(LOG_ERROR, msg); }
};

// 配置管理器
class ConfigManager {
private:
    std::wstring configFile;
    std::wstring backupDir;
    int maxRetentionDays;
    int scanIntervalSeconds;
    Logger& logger;

    void CreateDefaultConfig() {
        backupDir = L"./";
        maxRetentionDays = 30;
        scanIntervalSeconds = 30;
        SaveConfig();
    }

    void ParseConfig(const std::wstring& content) {
        std::wistringstream iss(content);
        std::wstring line;
        
        while (std::getline(iss, line)) {
            // 移除首尾空格
            size_t start = line.find_first_not_of(L" \t\r\n");
            if (start == std::wstring::npos) continue;
            size_t end = line.find_last_not_of(L" \t\r\n");
            line = line.substr(start, end - start + 1);
            
            if (line.empty() || line[0] == L'#') continue;
            
            size_t eqPos = line.find(L'=');
            if (eqPos == std::wstring::npos) continue;
            
            std::wstring key = line.substr(0, eqPos);
            std::wstring value = line.substr(eqPos + 1);
            
            // 移除值中的空格
            value.erase(0, value.find_first_not_of(L" \t"));
            value.erase(value.find_last_not_of(L" \t") + 1);
            
            if (key == L"backup_dir") {
                backupDir = value;
            }
            else if (key == L"max_retention_days") {
                maxRetentionDays = std::stoi(value);
            }
            else if (key == L"scan_interval_seconds") {
                scanIntervalSeconds = std::stoi(value);
            }
        }
    }

public:
    ConfigManager(Logger& log) : logger(log), maxRetentionDays(30), scanIntervalSeconds(30) {
        wchar_t exePath[MAX_PATH];
        GetModuleFileNameW(NULL, exePath, MAX_PATH);
        configFile = fs::path(exePath).parent_path().wstring() + L"\\ppt_monitor.ini";
        
        if (fs::exists(configFile)) {
            std::wifstream file(configFile);
            if (file.is_open()) {
                std::wstringstream buffer;
                buffer << file.rdbuf();
                ParseConfig(buffer.str());
                logger.Info(L"成功加载配置文件: " + configFile);
            }
        }
        else {
            logger.Info(L"配置文件不存在，创建默认配置");
            CreateDefaultConfig();
        }
    }

    void SaveConfig() {
        std::wofstream file(configFile);
        if (file.is_open()) {
            file << L"# PowerPoint PPT文件备份监控配置文件\n";
            file << L"# 备份目录\n";
            file << L"backup_dir = " << backupDir << L"\n";
            file << L"# 最大保留天数\n";
            file << L"max_retention_days = " << maxRetentionDays << L"\n";
            file << L"# 设备扫描间隔（秒）\n";
            file << L"scan_interval_seconds = " << scanIntervalSeconds << L"\n";
            file.close();
        }
    }

    std::wstring GetBackupDir() const { return backupDir; }
    int GetMaxRetentionDays() const { return maxRetentionDays; }
    int GetScanInterval() const { return scanIntervalSeconds; }
    
    void SetBackupDir(const std::wstring& dir) {
        backupDir = dir;
        SaveConfig();
    }
    
    void SetMaxRetentionDays(int days) {
        maxRetentionDays = days;
        SaveConfig();
    }
    
    void SetScanInterval(int seconds) {
        scanIntervalSeconds = seconds;
        SaveConfig();
    }
};

// 持久化文件管理器
class PersistentFileManager {
private:
    std::wstring stateFile;
    std::wstring today;
    std::map<std::wstring, time_t> processedFiles;
    std::mutex mutex;
    Logger& logger;

    void LoadState() {
        if (fs::exists(stateFile)) {
            std::wifstream file(stateFile);
            if (file.is_open()) {
                std::wstring line;
                std::getline(file, line); // 读取日期行
                
                if (line.find(L"date:") == 0) {
                    std::wstring storedDate = line.substr(5);
                    if (storedDate == today) {
                        while (std::getline(file, line)) {
                            size_t colonPos = line.find(L':');
                            if (colonPos != std::wstring::npos) {
                                std::wstring filePath = line.substr(0, colonPos);
                                time_t mtime = std::stoll(line.substr(colonPos + 1));
                                processedFiles[filePath] = mtime;
                            }
                        }
                        logger.Info(L"加载状态成功，今日已处理 " + std::to_wstring(processedFiles.size()) + L" 个文件");
                    }
                    else {
                        logger.Info(L"状态文件日期不匹配，已重置");
                    }
                }
            }
        }
    }

public:
    PersistentFileManager(const std::wstring& baseBackupDir, Logger& log) : logger(log) {
        stateFile = baseBackupDir + L"\\monitor_state.txt";
        auto now = std::chrono::system_clock::now();
        auto time_t_now = std::chrono::system_clock::to_time_t(now);
        struct tm tm_now;
        localtime_s(&tm_now, &time_t_now);
        
        wchar_t buffer[16];
        wcsftime(buffer, 16, L"%Y-%m-%d", &tm_now);
        today = buffer;
        
        LoadState();
    }

    void SaveStateImmediately() {
        std::lock_guard<std::mutex> lock(mutex);
        std::wofstream file(stateFile);
        if (file.is_open()) {
            file << L"date:" << today << L"\n";
            for (const auto& pair : processedFiles) {
                file << pair.first << L":" << pair.second << L"\n";
            }
        }
    }

    void AddProcessedFile(const std::wstring& filePath, time_t mtime) {
        std::lock_guard<std::mutex> lock(mutex);
        processedFiles[filePath] = mtime;
        SaveStateImmediately();
    }

    bool IsAlreadyProcessed(const std::wstring& filePath) {
        std::lock_guard<std::mutex> lock(mutex);
        return processedFiles.find(filePath) != processedFiles.end();
    }

    time_t GetFileMtime(const std::wstring& filePath) {
        std::lock_guard<std::mutex> lock(mutex);
        auto it = processedFiles.find(filePath);
        if (it != processedFiles.end()) {
            return it->second;
        }
        return 0;
    }

    void CleanupOldState() {
        auto now = std::chrono::system_clock::now();
        auto time_t_now = std::chrono::system_clock::to_time_t(now);
        struct tm tm_now;
        localtime_s(&tm_now, &time_t_now);
        
        wchar_t buffer[16];
        wcsftime(buffer, 16, L"%Y-%m-%d", &tm_now);
        std::wstring currentDate(buffer);
        
        if (currentDate != today) {
            std::lock_guard<std::mutex> lock(mutex);
            processedFiles.clear();
            today = currentDate;
            SaveStateImmediately();
            logger.Info(L"跨天重置状态，新日期: " + today);
        }
    }

    size_t GetProcessedCount() const {
        return processedFiles.size();
    }
};

// 设备事件监听器
class DeviceEventMonitor {
private:
    HWND hwnd;
    std::unique_ptr<std::thread> eventThread;
    bool running;
    std::function<void(const std::vector<std::wstring>&, bool)> callback;
    Logger& logger;

    static LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
        DeviceEventMonitor* pThis = nullptr;
        
        if (msg == WM_NCCREATE) {
            CREATESTRUCT* pCreate = reinterpret_cast<CREATESTRUCT*>(lParam);
            pThis = reinterpret_cast<DeviceEventMonitor*>(pCreate->lpCreateParams);
            SetWindowLongPtr(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(pThis));
        }
        else {
            pThis = reinterpret_cast<DeviceEventMonitor*>(GetWindowLongPtr(hWnd, GWLP_USERDATA));
        }
        
        if (pThis) {
            return pThis->HandleMessage(hWnd, msg, wParam, lParam);
        }
        
        return DefWindowProc(hWnd, msg, wParam, lParam);
    }

    LRESULT HandleMessage(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
        if (msg == WM_DEVICECHANGE) {
            if (wParam == DBT_DEVICEARRIVAL || wParam == DBT_DEVICEREMOVECOMPLETE) {
                PDEV_BROADCAST_HDR lpdb = (PDEV_BROADCAST_HDR)lParam;
                if (lpdb && lpdb->dbch_devicetype == DBT_DEVTYP_VOLUME) {
                    PDEV_BROADCAST_VOLUME lpdbv = (PDEV_BROADCAST_VOLUME)lpdb;
                    DWORD unitMask = lpdbv->dbcv_unitmask;
                    
                    std::vector<std::wstring> drives;
                    for (int i = 0; i < 26; i++) {
                        if (unitMask & (1 << i)) {
                            wchar_t drive[3] = { static_cast<wchar_t>(L'A' + i), L':', 0 };
                            drives.push_back(std::wstring(drive));
                        }
                    }
                    
                    bool isInsert = (wParam == DBT_DEVICEARRIVAL);
                    logger.Info(isInsert ? L"检测到设备插入" : L"检测到设备移除");
                    
                    if (callback) {
                        callback(drives, isInsert);
                    }
                }
            }
            return TRUE;
        }
        else if (msg == WM_DESTROY) {
            PostQuitMessage(0);
            return 0;
        }
        
        return DefWindowProc(hWnd, msg, wParam, lParam);
    }

    void RunMessageLoop() {
        MSG msg;
        while (running && GetMessage(&msg, NULL, 0, 0)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

public:
    DeviceEventMonitor(Logger& log) : hwnd(nullptr), running(false), logger(log) {}

    ~DeviceEventMonitor() {
        StopListening();
    }

    void StartListening(std::function<void(const std::vector<std::wstring>&, bool)> cb) {
        callback = cb;
        
        WNDCLASSEX wc = {};
        wc.cbSize = sizeof(WNDCLASSEX);
        wc.lpfnWndProc = WndProc;
        wc.hInstance = GetModuleHandle(NULL);
        wc.lpszClassName = L"PPTMonitorDeviceListener";
        
        RegisterClassEx(&wc);
        
        hwnd = CreateWindowEx(
            0,
            L"PPTMonitorDeviceListener",
            L"PPTMonitorDeviceListenerWindow",
            0,
            0, 0, 0, 0,
            NULL, NULL,
            GetModuleHandle(NULL),
            this
        );
        
        if (hwnd) {
            running = true;
            
            // 获取初始驱动器
            std::vector<std::wstring> initialDrives = GetRemovableDrives();
            if (!initialDrives.empty() && callback) {
                callback(initialDrives, true);
            }
            
            // 注册设备通知
            DEV_BROADCAST_DEVICEINTERFACE dbdi = {};
            dbdi.dbcc_size = sizeof(DEV_BROADCAST_DEVICEINTERFACE);
            dbdi.dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE;
            
            HDEVNOTIFY hDevNotify = RegisterDeviceNotification(
                hwnd,
                &dbdi,
                DEVICE_NOTIFY_WINDOW_HANDLE
            );
            
            eventThread = std::make_unique<std::thread>(&DeviceEventMonitor::RunMessageLoop, this);
            logger.Info(L"设备事件监听器已启动");
        }
    }

    void StopListening() {
        running = false;
        if (hwnd) {
            PostMessage(hwnd, WM_QUIT, 0, 0);
            if (eventThread && eventThread->joinable()) {
                eventThread->join();
            }
            DestroyWindow(hwnd);
            hwnd = nullptr;
        }
        logger.Info(L"设备事件监听器已停止");
    }

    std::vector<std::wstring> GetRemovableDrives() {
        std::vector<std::wstring> drives;
        DWORD drivesMask = GetLogicalDrives();
        
        for (int i = 0; i < 26; i++) {
            if (drivesMask & (1 << i)) {
                wchar_t drive[4] = { static_cast<wchar_t>(L'A' + i), L':', L'\\', 0 };
                UINT driveType = GetDriveTypeW(drive);
                
                if (driveType == DRIVE_REMOVABLE || driveType == DRIVE_CDROM) {
                    wchar_t driveLetter[3] = { static_cast<wchar_t>(L'A' + i), L':', 0 };
                    drives.push_back(std::wstring(driveLetter));
                }
            }
        }
        
        return drives;
    }
};

// PPT监控器
class PPTMonitor {
private:
    Logger& logger;
    ConfigManager configManager;
    PersistentFileManager fileManager;
    DeviceEventMonitor deviceMonitor;
    
    std::set<std::wstring> connectedDrives;
    std::set<std::wstring> removableDrivesCache;
    std::unique_ptr<std::thread> scanThread;
    std::unique_ptr<std::thread> monitorThread;
    bool running;
    bool scanThreadRunning;
    
    std::wstring baseBackupDir;
    int maxRetentionDays;
    int scanInterval;
    
    std::mutex drivesMutex;

    bool IsRemovableDrive(const std::wstring& path) {
        if (path.empty()) return false;
        
        std::wstring driveLetter = path.substr(0, 2);
        std::transform(driveLetter.begin(), driveLetter.end(), driveLetter.begin(), ::towupper);
        
        std::lock_guard<std::mutex> lock(drivesMutex);
        return removableDrivesCache.find(driveLetter) != removableDrivesCache.end();
    }

    bool IsValidPPTFile(const std::wstring& filePath) {
        if (filePath.empty()) return false;
        
        // 检查是否是临时文件
        std::wstring filename = fs::path(filePath).filename().wstring();
        if (filename.find(L"~$") == 0) return false;
        
        // 检查隐藏属性
        DWORD attrs = GetFileAttributesW(filePath.c_str());
        if (attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_HIDDEN)) {
            return false;
        }
        
        // 检查扩展名
        std::wstring ext = fs::path(filePath).extension().wstring();
        std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
        
        return PPT_EXTENSIONS.find(ext) != PPT_EXTENSIONS.end();
    }

    bool HasFileChanged(const std::wstring& filePath) {
        try {
            struct _stat fileStat;
            if (_wstat(filePath.c_str(), &fileStat) != 0) {
                return true;
            }
            
            fileManager.CleanupOldState();
            
            if (fileManager.IsAlreadyProcessed(filePath)) {
                time_t storedMtime = fileManager.GetFileMtime(filePath);
                if (abs(fileStat.st_mtime - storedMtime) < 1) {
                    return false;
                }
                else {
                    fileManager.AddProcessedFile(filePath, fileStat.st_mtime);
                    return true;
                }
            }
            else {
                fileManager.AddProcessedFile(filePath, fileStat.st_mtime);
                return true;
            }
        }
        catch (...) {
            return true;
        }
    }

    std::wstring GetDateFolderPath() {
        auto now = std::chrono::system_clock::now();
        auto time_t_now = std::chrono::system_clock::to_time_t(now);
        struct tm tm_now;
        localtime_s(&tm_now, &time_t_now);
        
        wchar_t buffer[16];
        wcsftime(buffer, 16, L"%Y-%m-%d", &tm_now);
        std::wstring dateFolder = baseBackupDir + L"\\" + std::wstring(buffer);
        
        if (!fs::exists(dateFolder)) {
            fs::create_directories(dateFolder);
        }
        
        return dateFolder;
    }

    bool CopyPPTFile(const std::wstring& sourcePath) {
        try {
            if (!fs::exists(sourcePath)) {
                logger.Warning(L"源文件不存在: " + sourcePath);
                return false;
            }
            
            std::wstring dateFolder = GetDateFolderPath();
            std::wstring filename = fs::path(sourcePath).filename().wstring();
            std::wstring destPath = dateFolder + L"\\" + filename;
            
            int counter = 1;
            std::wstring originalDestPath = destPath;
            
            while (fs::exists(destPath)) {
                std::wstring stem = fs::path(originalDestPath).stem().wstring();
                std::wstring ext = fs::path(originalDestPath).extension().wstring();
                destPath = dateFolder + L"\\" + stem + L"_" + std::to_wstring(counter) + ext;
                counter++;
            }
            
            fs::copy(sourcePath, destPath, fs::copy_options::overwrite_existing);
            logger.Info(L"已备份PPT文件: " + sourcePath + L" -> " + destPath);
            
            struct _stat fileStat;
            if (_wstat(sourcePath.c_str(), &fileStat) == 0) {
                fileManager.AddProcessedFile(sourcePath, fileStat.st_mtime);
            }
            
            return true;
        }
        catch (const std::exception& e) {
            logger.Error(L"备份失败: " + std::wstring(e.what(), e.what() + strlen(e.what())));
            return false;
        }
    }

    void ProcessPotentialPPTFile(const std::wstring& filePath) {
        if (IsValidPPTFile(filePath) && IsRemovableDrive(filePath) && HasFileChanged(filePath)) {
            logger.Info(L"检测到移动设备上的PPT文件: " + filePath);
            CopyPPTFile(filePath);
        }
    }

    void MonitorViaProcess() {
        DWORD processes[1024];
        DWORD bytesReturned;
        
        if (!EnumProcesses(processes, sizeof(processes), &bytesReturned)) {
            return;
        }
        
        DWORD numProcesses = bytesReturned / sizeof(DWORD);
        
        for (DWORD i = 0; i < numProcesses; i++) {
            if (processes[i] == 0) continue;
            
            HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, processes[i]);
            if (hProcess) {
                wchar_t processName[MAX_PATH];
                if (GetModuleBaseNameW(hProcess, NULL, processName, MAX_PATH)) {
                    std::wstring name(processName);
                    std::transform(name.begin(), name.end(), name.begin(), ::towlower);
                    
                    if (name.find(L"powerpnt") == 0) {
                        // 获取进程打开的文件
                        // 这里简化处理，实际需要通过其他方式获取
                    }
                }
                CloseHandle(hProcess);
            }
        }
    }

    void MonitorViaCOM() {
        // COM监控需要更复杂的实现，这里简化处理
        // 实际可以通过GetActiveObject等API获取PowerPoint实例
    }

    void CleanupOldBackups() {
        try {
            auto now = std::chrono::system_clock::now();
            auto retentionTime = now - std::chrono::hours(24 * maxRetentionDays);
            
            for (const auto& entry : fs::directory_iterator(baseBackupDir)) {
                if (entry.is_directory()) {
                    std::wstring dirName = entry.path().filename().wstring();
                    
                    struct tm tm_dir;
                    if (wcsftime(const_cast<wchar_t*>(dirName.c_str()), dirName.length() + 1, L"%Y-%m-%d", &tm_dir)) {
                        std::tm tm = {};
                        std::wistringstream iss(dirName);
                        int year, month, day;
                        wchar_t dash;
                        iss >> year >> dash >> month >> dash >> day;
                        
                        tm.tm_year = year - 1900;
                        tm.tm_mon = month - 1;
                        tm.tm_mday = day;
                        
                        time_t dirTime = mktime(&tm);
                        auto dirTimePoint = std::chrono::system_clock::from_time_t(dirTime);
                        
                        if (dirTimePoint < retentionTime) {
                            fs::remove_all(entry.path());
                            logger.Info(L"已删除过期备份文件夹: " + entry.path().wstring());
                        }
                    }
                }
            }
        }
        catch (const std::exception& e) {
            logger.Error(L"清理旧备份时出错");
        }
    }

    void ScanDevicesPeriodically() {
        logger.Info(L"设备定期扫描线程已启动");
        
        while (scanThreadRunning) {
            std::this_thread::sleep_for(std::chrono::seconds(scanInterval));
            
            if (!scanThreadRunning) break;
            
            std::vector<std::wstring> currentDrives = deviceMonitor.GetRemovableDrives();
            std::set<std::wstring> currentDrivesSet(currentDrives.begin(), currentDrives.end());
            
            std::set<std::wstring> cachedDrivesSet;
            {
                std::lock_guard<std::mutex> lock(drivesMutex);
                cachedDrivesSet = connectedDrives;
            }
            
            std::vector<std::wstring> newDrives;
            for (const auto& drive : currentDrivesSet) {
                if (cachedDrivesSet.find(drive) == cachedDrivesSet.end()) {
                    newDrives.push_back(drive);
                }
            }
            
            if (!newDrives.empty()) {
                logger.Warning(L"定期扫描发现遗漏的设备插入事件");
                OnDeviceEvent(newDrives, true);
            }
            
            std::vector<std::wstring> removedDrives;
            for (const auto& drive : cachedDrivesSet) {
                if (currentDrivesSet.find(drive) == currentDrivesSet.end()) {
                    removedDrives.push_back(drive);
                }
            }
            
            if (!removedDrives.empty()) {
                logger.Warning(L"定期扫描发现遗漏的设备移除事件");
                OnDeviceEvent(removedDrives, false);
            }
        }
        
        logger.Info(L"设备定期扫描线程已停止");
    }

    void MonitoringLoop() {
        logger.Info(L"监控主循环已启动");
        
        while (running) {
            try {
                auto now = std::chrono::system_clock::now();
                auto time_t_now = std::chrono::system_clock::to_time_t(now);
                struct tm tm_now;
                localtime_s(&tm_now, &time_t_now);
                
                // 每天0点清理过期备份
                if (tm_now.tm_hour == 0 && tm_now.tm_min == 0 && tm_now.tm_sec < 10) {
                    CleanupOldBackups();
                }
                
                fileManager.CleanupOldState();
                MonitorViaCOM();
                MonitorViaProcess();
                
                std::this_thread::sleep_for(std::chrono::seconds(2));
            }
            catch (...) {
                logger.Error(L"监控过程中出现错误");
                std::this_thread::sleep_for(std::chrono::seconds(2));
            }
        }
        
        logger.Info(L"监控主循环已停止");
    }

public:
    PPTMonitor(Logger& log) 
        : logger(log), configManager(log), fileManager(configManager.GetBackupDir(), log), 
          deviceMonitor(log), running(false), scanThreadRunning(false) {
        
        baseBackupDir = configManager.GetBackupDir();
        maxRetentionDays = configManager.GetMaxRetentionDays();
        scanInterval = configManager.GetScanInterval();
        
        if (!fs::exists(baseBackupDir)) {
            fs::create_directories(baseBackupDir);
        }
        
        logger.Info(L"初始化PPT监控器");
        logger.Info(L"备份目录: " + baseBackupDir);
        logger.Info(L"最大保留天数: " + std::to_wstring(maxRetentionDays));
        logger.Info(L"设备扫描间隔: " + std::to_wstring(scanInterval) + L"秒");
    }

    void OnDeviceEvent(const std::vector<std::wstring>& drives, bool isInsert) {
        std::lock_guard<std::mutex> lock(drivesMutex);
        
        if (isInsert) {
            for (const auto& drive : drives) {
                connectedDrives.insert(drive);
                removableDrivesCache.insert(drive);
            }
            logger.Info(L"设备插入事件");
        }
        else {
            for (const auto& drive : drives) {
                connectedDrives.erase(drive);
                removableDrivesCache.erase(drive);
            }
            logger.Info(L"设备移除事件");
        }
        
        std::wstring drivesStr;
        for (const auto& drive : connectedDrives) {
            drivesStr += drive + L" ";
        }
        logger.Info(L"当前连接设备: " + drivesStr);
    }

    std::vector<std::wstring> GetConnectedDrives() {
        std::lock_guard<std::mutex> lock(drivesMutex);
        return std::vector<std::wstring>(connectedDrives.begin(), connectedDrives.end());
    }

    void StartMonitoring() {
        if (running) return;
        
        running = true;
        scanThreadRunning = true;
        
        deviceMonitor.StartListening([this](const std::vector<std::wstring>& drives, bool isInsert) {
            OnDeviceEvent(drives, isInsert);
        });
        
        scanThread = std::make_unique<std::thread>(&PPTMonitor::ScanDevicesPeriodically, this);
        monitorThread = std::make_unique<std::thread>(&PPTMonitor::MonitoringLoop, this);
        
        logger.Info(L"PPT监控程序已启动");
    }

    void StopMonitoring() {
        running = false;
        scanThreadRunning = false;
        
        deviceMonitor.StopListening();
        
        if (scanThread && scanThread->joinable()) {
            scanThread->join();
        }
        
        if (monitorThread && monitorThread->joinable()) {
            monitorThread->join();
        }
        
        fileManager.SaveStateImmediately();
        logger.Info(L"PPT监控程序已停止");
    }

    size_t GetProcessedCount() const {
        return fileManager.GetProcessedCount();
    }
};

// 系统托盘应用
class SystemTrayApp {
private:
    HWND hwnd;
    NOTIFYICONDATAW nid;
    std::unique_ptr<PPTMonitor> monitor;
    std::unique_ptr<std::thread> monitorThread;
    ConfigManager configManager;
    Logger& logger;
    bool running;

    static LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
        SystemTrayApp* pThis = nullptr;
        
        if (msg == WM_NCCREATE) {
            CREATESTRUCT* pCreate = reinterpret_cast<CREATESTRUCT*>(lParam);
            pThis = reinterpret_cast<SystemTrayApp*>(pCreate->lpCreateParams);
            SetWindowLongPtr(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(pThis));
        }
        else {
            pThis = reinterpret_cast<SystemTrayApp*>(GetWindowLongPtr(hWnd, GWLP_USERDATA));
        }
        
        if (pThis) {
            return pThis->HandleMessage(hWnd, msg, wParam, lParam);
        }
        
        return DefWindowProc(hWnd, msg, wParam, lParam);
    }

    LRESULT HandleMessage(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
        switch (msg) {
        case WM_TRAYICON:
            if (lParam == WM_RBUTTONUP) {
                POINT pt;
                GetCursorPos(&pt);
                ShowContextMenu(pt.x, pt.y);
            }
            else if (lParam == WM_LBUTTONDBLCLK) {
                // 双击托盘图标
            }
            break;
            
        case WM_COMMAND:
            switch (LOWORD(wParam)) {
            case 1000: // 打开配置文件
                OpenConfigFile();
                break;
            case 1001: // 退出
                ExitApp();
                break;
            }
            break;
            
        case WM_DESTROY:
            PostQuitMessage(0);
            return 0;
        }
        
        return DefWindowProc(hWnd, msg, wParam, lParam);
    }

    void ShowContextMenu(int x, int y) {
        HMENU hMenu = CreatePopupMenu();
        
        std::wstring statusText;
        std::wstring retentionText;
        std::wstring scanIntervalText;
        std::wstring processedText;
        
        if (monitor) {
            auto drives = monitor->GetConnectedDrives();
            if (!drives.empty()) {
                statusText = L"连接设备: ";
                for (const auto& drive : drives) {
                    statusText += drive + L" ";
                }
            }
            else {
                statusText = L"无连接设备";
            }
            
            retentionText = L"保留天数: " + std::to_wstring(configManager.GetMaxRetentionDays()) + L"天";
            scanIntervalText = L"扫描间隔: " + std::to_wstring(configManager.GetScanInterval()) + L"秒";
            processedText = L"今日处理: " + std::to_wstring(monitor->GetProcessedCount()) + L"个文件";
        }
        else {
            statusText = L"无连接设备";
            retentionText = L"保留天数: " + std::to_wstring(configManager.GetMaxRetentionDays()) + L"天";
            scanIntervalText = L"扫描间隔: " + std::to_wstring(configManager.GetScanInterval()) + L"秒";
            processedText = L"今日处理: 0个文件";
        }
        
        AppendMenuW(hMenu, MF_GRAYED, 0, statusText.c_str());
        AppendMenuW(hMenu, MF_GRAYED, 0, retentionText.c_str());
        AppendMenuW(hMenu, MF_GRAYED, 0, scanIntervalText.c_str());
        AppendMenuW(hMenu, MF_GRAYED, 0, processedText.c_str());
        AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
        AppendMenuW(hMenu, MF_STRING, 1000, L"打开配置文件");
        AppendMenuW(hMenu, MF_STRING, 1001, L"退出");
        
        SetForegroundWindow(hwnd);
        TrackPopupMenu(hMenu, TPM_LEFTALIGN, x, y, 0, hwnd, NULL);
        PostMessage(hwnd, WM_NULL, 0, 0);
        
        DestroyMenu(hMenu);
    }

    void OpenConfigFile() {
        std::wstring configFile = fs::path(configManager.GetBackupDir()).parent_path().wstring() + L"\\ppt_monitor.ini";
        ShellExecuteW(NULL, L"open", configFile.c_str(), NULL, NULL, SW_SHOW);
    }

    void ExitApp() {
        running = false;
        
        if (monitor) {
            monitor->StopMonitoring();
        }
        
        Shell_NotifyIconW(NIM_DELETE, &nid);
        PostQuitMessage(0);
    }

    void CreateTrayIcon() {
        nid.cbSize = sizeof(NOTIFYICONDATAW);
        nid.hWnd = hwnd;
        nid.uID = 1;
        nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
        nid.uCallbackMessage = WM_TRAYICON;
        nid.hIcon = LoadIcon(NULL, IDI_APPLICATION);
        wcscpy_s(nid.szTip, L"PPT文件备份监控");
        
        Shell_NotifyIconW(NIM_ADD, &nid);
    }

public:
    SystemTrayApp(Logger& log) : logger(log), configManager(log), running(true) {
        logger.Info(L"初始化系统托盘应用");
        
        WNDCLASSEXW wc = {};
        wc.cbSize = sizeof(WNDCLASSEXW);
        wc.lpfnWndProc = WndProc;
        wc.hInstance = GetModuleHandle(NULL);
        wc.lpszClassName = L"PPTMonitorTrayClass";
        
        RegisterClassExW(&wc);
        
        hwnd = CreateWindowExW(
            0,
            L"PPTMonitorTrayClass",
            L"PPT Monitor",
            WS_OVERLAPPEDWINDOW,
            CW_USEDEFAULT, CW_USEDEFAULT, 200, 200,
            NULL, NULL,
            GetModuleHandle(NULL),
            this
        );
        
        CreateTrayIcon();
    }

    ~SystemTrayApp() {
        Shell_NotifyIconW(NIM_DELETE, &nid);
    }

    void StartMonitoring() {
        monitor = std::make_unique<PPTMonitor>(logger);
        monitor->StartMonitoring();
        logger.Info(L"监控线程已启动");
    }

    void Run() {
        StartMonitoring();
        logger.Info(L"PPT监控程序已启动，托盘图标已显示");
        
        MSG msg;
        while (running && GetMessage(&msg, NULL, 0, 0)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
        
        logger.Info(L"应用程序主循环已结束");
    }
};

// 主函数
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    // 设置控制台编码为UTF-8
    SetConsoleOutputCP(CP_UTF8);
    
    Logger logger;
    logger.Info(L"==================================================");
    logger.Info(L"PowerPoint PPT文件备份监控程序启动...");
    logger.Info(L"仅监控PowerPoint打开的移动存储设备上的PPT文件");
    
    try {
        SystemTrayApp app(logger);
        app.Run();
    }
    catch (const std::exception& e) {
        logger.Error(L"程序运行出错");
    }
    
    logger.Info(L"程序已完全退出");
    return 0;
}
