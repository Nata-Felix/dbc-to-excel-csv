#include <windows.h>
#include <shlobj.h>

#include <cstdint>
#include <string>

namespace {

constexpr int kCoreResource = 301;
constexpr int kDuckDbResource = 302;
constexpr wchar_t kVersion[] = L"1.1.0";

std::wstring windows_error(DWORD code) {
    wchar_t* buffer = nullptr;
    FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM |
                       FORMAT_MESSAGE_IGNORE_INSERTS,
                   nullptr, code, 0, reinterpret_cast<wchar_t*>(&buffer), 0, nullptr);
    std::wstring message = buffer ? buffer : L"erro desconhecido";
    if (buffer) LocalFree(buffer);
    return message;
}

bool same_size(const std::wstring& path, DWORD size) {
    WIN32_FILE_ATTRIBUTE_DATA data{};
    if (!GetFileAttributesExW(path.c_str(), GetFileExInfoStandard, &data)) return false;
    return data.nFileSizeHigh == 0 && data.nFileSizeLow == size;
}

bool extract_resource(HINSTANCE instance, int id, const std::wstring& destination,
                      std::wstring& error) {
    HRSRC resource = FindResourceW(instance, MAKEINTRESOURCEW(id), RT_RCDATA);
    if (!resource) {
        error = L"Recurso interno não encontrado: " + std::to_wstring(id);
        return false;
    }
    const DWORD size = SizeofResource(instance, resource);
    if (same_size(destination, size)) return true;
    HGLOBAL loaded = LoadResource(instance, resource);
    const void* bytes = loaded ? LockResource(loaded) : nullptr;
    if (!bytes || size == 0) {
        error = L"Não foi possível ler um recurso interno.";
        return false;
    }

    const std::wstring temporary = destination + L".tmp";
    HANDLE file = CreateFileW(temporary.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS,
                              FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        error = L"Não foi possível preparar os arquivos locais. " + windows_error(GetLastError());
        return false;
    }
    DWORD written = 0;
    const BOOL ok = WriteFile(file, bytes, size, &written, nullptr);
    const DWORD write_error = ok ? ERROR_SUCCESS : GetLastError();
    CloseHandle(file);
    if (!ok || written != size) {
        DeleteFileW(temporary.c_str());
        error = L"Falha ao gravar os arquivos internos. " + windows_error(write_error);
        return false;
    }
    if (!MoveFileExW(temporary.c_str(), destination.c_str(),
                     MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
        const DWORD move_error = GetLastError();
        DeleteFileW(temporary.c_str());
        error = L"Falha ao atualizar os arquivos locais. " + windows_error(move_error);
        return false;
    }
    return true;
}

void show_error(const std::wstring& message) {
    MessageBoxW(nullptr, message.c_str(), L"DBC para Excel, CSV e Parquet - SOLPPE",
                MB_OK | MB_ICONERROR);
}

}  // namespace

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int) {
    wchar_t local_app_data[MAX_PATH]{};
    if (FAILED(SHGetFolderPathW(nullptr, CSIDL_LOCAL_APPDATA | CSIDL_FLAG_CREATE, nullptr,
                                SHGFP_TYPE_CURRENT, local_app_data))) {
        show_error(L"Não foi possível localizar a pasta de dados do Windows.");
        return 2;
    }

    const std::wstring directory = std::wstring(local_app_data) +
        L"\\SOLPPE\\DBC-to-Excel-CSV\\" + kVersion;
    if (SHCreateDirectoryExW(nullptr, directory.c_str(), nullptr) != ERROR_SUCCESS &&
        GetFileAttributesW(directory.c_str()) == INVALID_FILE_ATTRIBUTES) {
        show_error(L"Não foi possível criar a pasta local do aplicativo.");
        return 2;
    }

    const std::wstring core = directory + L"\\DBC_Converter_Core.exe";
    const std::wstring duckdb = directory + L"\\duckdb.dll";
    std::wstring error;
    if (!extract_resource(instance, kCoreResource, core, error) ||
        !extract_resource(instance, kDuckDbResource, duckdb, error)) {
        show_error(error);
        return 2;
    }

    std::wstring command_line = L"\"" + core + L"\"";
    STARTUPINFOW startup{};
    startup.cb = sizeof(startup);
    PROCESS_INFORMATION process{};
    if (!CreateProcessW(core.c_str(), command_line.data(), nullptr, nullptr, FALSE,
                        CREATE_UNICODE_ENVIRONMENT, nullptr, directory.c_str(), &startup, &process)) {
        show_error(L"Não foi possível abrir o conversor. " + windows_error(GetLastError()));
        return 2;
    }
    CloseHandle(process.hThread);
    CloseHandle(process.hProcess);
    return 0;
}
