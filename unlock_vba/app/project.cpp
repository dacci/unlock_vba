// Copyright (c) 2016 dacci.org

#include <base/win/scoped_comptr.h>

#include <string>
#include <strstream>

#include "app/application.h"

static HRESULT UnlockStream(IStream* stream) {
  HRESULT result;

  do {
    std::stringstream buffer;
    while (true) {
      char chunk[256];
      ULONG read;
      result = stream->Read(chunk, sizeof(chunk), &read);
      if (FAILED(result))
        break;

      buffer.write(chunk, read);

      if (result == S_FALSE || read < sizeof(chunk))
        break;
    }
    if (FAILED(result))
      break;

    result = stream->Seek(LARGE_INTEGER{0}, STREAM_SEEK_SET, nullptr);
    if (FAILED(result))
      break;

    result = stream->SetSize(ULARGE_INTEGER{0});
    if (FAILED(result))
      break;

    std::string line;
    while (std::getline(buffer, line)) {
      if (strncmp(line.c_str(), "CMG=", 4) == 0 ||
          strncmp(line.c_str(), "DPB=", 4) == 0 ||
          strncmp(line.c_str(), "GC=", 3) == 0)
        continue;

      result =
          stream->Write(line.data(), static_cast<ULONG>(line.size()), nullptr);
      if (FAILED(result))
        break;

      result = stream->Write("\x0A", 1, nullptr);
      if (FAILED(result))
        break;
    }
  } while (false);

  if (FAILED(result))
    return stream->Revert();

  return stream->Commit(STGC_DEFAULT);
}

HRESULT UnlockProject(IStorage* storage) {
  base::win::ScopedComPtr<IStream> stream;
  auto result = storage->OpenStream(L"PROJECT", nullptr,
                                    STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0,
                                    stream.Receive());
  if (FAILED(result))
    return result;

  result = UnlockStream(stream.get());
  stream.Release();

  if (FAILED(result))
    return storage->Revert();

  return storage->Commit(STGC_DEFAULT);
}
