// Copyright (c) 2016 dacci.org

#include <zip.h>

#pragma warning(push, 3)
#pragma warning(disable : 4244)
#include <base/files/file_util.h>
#include <base/strings/sys_string_conversions.h>
#include <base/win/scoped_comptr.h>
#pragma warning(pop)

#include "app/application.h"

HRESULT UnlockProject(IStorage* storage);

void UnlockOpenXML(const wchar_t* path_string) {
  base::FilePath path(path_string), temp_path;
  if (!base::CreateTemporaryFileInDir(path.DirName(), &temp_path))
    return;

  zip_t* zip = nullptr;

  do {
    base::win::ScopedHandle temp_file(
        CreateFile(temp_path.value().c_str(), GENERIC_WRITE, 0, nullptr,
                   OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL));
    if (!temp_file.IsValid())
      break;

    auto zip_path = base::SysWideToUTF8(path_string);
    zip = zip_open(zip_path.c_str(), 0, nullptr);
    if (zip == nullptr)
      break;

    auto index = zip_name_locate(zip, "xl/vbaProject.bin", 0);
    if (index == -1)
      break;

    auto zip_file = zip_fopen_index(zip, index, ZIP_FL_UNCHANGED);
    if (zip_file == nullptr)
      break;

    char buffer[256];
    auto error = false;
    while (true) {
      auto read = zip_fread(zip_file, buffer, sizeof(buffer));
      if (read < 0) {
        error = true;
        break;
      }

      DWORD written;
      if (!WriteFile(temp_file.Get(), buffer, static_cast<DWORD>(read),
                     &written, nullptr) ||
          written != read) {
        error = true;
        break;
      }

      if (read < sizeof(buffer))
        break;
    }

    zip_fclose(zip_file);

    if (error)
      break;

    temp_file.Close();

    base::win::ScopedComPtr<IStorage> project_storage;
    auto result = StgOpenStorageEx(
        temp_path.value().c_str(),
        STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_TRANSACTED, STGFMT_STORAGE,
        0, nullptr, nullptr, project_storage.iid(),
        project_storage.ReceiveVoid());
    if (FAILED(result))
      break;

    result = UnlockProject(project_storage.get());
    if (FAILED(result))
      break;

    project_storage.Release();

    auto source_path = base::SysWideToUTF8(temp_path.value());
    auto source = zip_source_file_create(source_path.c_str(), 0, -1, nullptr);
    if (source == nullptr)
      break;

    index = zip_file_replace(zip, index, source, 0);
    if (index == -1)
      break;

    zip_close(zip);
    zip = nullptr;
  } while (false);

  if (zip != nullptr) {
    zip_discard(zip);
    zip = nullptr;
  }

  base::DeleteFile(temp_path, false);
}
