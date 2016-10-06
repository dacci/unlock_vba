// Copyright (c) 2016 dacci.org

#include "app/application.h"

#include <stdio.h>

#pragma warning(push, 3)
#pragma warning(disable : 4244)
#include <base/win/scoped_com_initializer.h>
#include <base/win/scoped_handle.h>
#pragma warning(pop)

enum class FileFormats {
  kUnknown,
  kCompoundBinary,
  kZip,
};

static FileFormats GetFileFormat(const wchar_t* path) {
  auto format = FileFormats::kUnknown;

  do {
    base::win::ScopedHandle file(CreateFile(path, GENERIC_READ, 0, nullptr,
                                            OPEN_EXISTING,
                                            FILE_ATTRIBUTE_NORMAL, NULL));
    if (!file.IsValid())
      break;

    DWORD signature, length;
    if (!ReadFile(file.Get(), &signature, sizeof(signature), &length,
                  nullptr) ||
        length != sizeof(signature))
      break;

    if (signature == 0x04034B50) {
      format = FileFormats::kZip;
      break;
    }

    if (signature != 0xE011CFD0)
      break;

    if (!ReadFile(file.Get(), &signature, sizeof(signature), &length,
                  nullptr) ||
        length != sizeof(signature))
      break;

    if (signature != 0xE11AB1A1)
      break;

    format = FileFormats::kCompoundBinary;
  } while (false);

  return format;
}

static void UnlockVBA(const wchar_t* path) {
  switch (GetFileFormat(path)) {
    case FileFormats::kCompoundBinary:
      UnlockCompoundBinary(path);
      break;

    case FileFormats::kZip:
      UnlockOpenXML(path);
      break;

    default:
      printf("unsupported\n");
  }
}

int wmain(int argc, wchar_t** argv) {
  base::win::ScopedCOMInitializer com_initializer;

  for (auto i = 1; i < argc; ++i)
    UnlockVBA(argv[i]);

  return 0;
}
