// Copyright (c) 2016 dacci.org

#include <base/win/scoped_comptr.h>

#include "app/application.h"

HRESULT UnlockProject(IStorage* storage);

HRESULT UnlockCompoundBinary(const wchar_t* path) {
  base::win::ScopedComPtr<IStorage> root_storage;
  auto result = StgOpenStorageEx(
      path, STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_TRANSACTED,
      STGFMT_STORAGE, 0, nullptr, nullptr, root_storage.iid(),
      root_storage.ReceiveVoid());
  if (FAILED(result))
    return result;

  base::win::ScopedComPtr<IStorage> project_storage;
  result = root_storage->OpenStorage(
      L"_VBA_PROJECT_CUR", nullptr,
      STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_TRANSACTED, nullptr, 0,
      project_storage.Receive());
  if (FAILED(result))
    return result;

  result = UnlockProject(project_storage.get());
  project_storage.Release();

  if (FAILED(result))
    return root_storage->Revert();

  return root_storage->Commit(STGC_DEFAULT);
}
