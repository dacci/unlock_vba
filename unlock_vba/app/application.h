// Copyright (c) 2016 dacci.org

#ifndef UNLOCK_VBA_APP_APPLICATION_H_
#define UNLOCK_VBA_APP_APPLICATION_H_

#include <winerror.h>

void UnlockOpenXML(const wchar_t* path);
HRESULT UnlockCompoundBinary(const wchar_t* path);

#endif  // UNLOCK_VBA_APP_APPLICATION_H_
