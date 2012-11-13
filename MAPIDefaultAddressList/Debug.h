#pragma once

#ifdef _DEBUG

#include <mapidefs.h>

// In debug mode, declare functions
void TraceDefaultDir(IAddrBook &AddrBook);

#else

// In release mode, remove functionality
#define TraceDefaultDir(x)

#endif