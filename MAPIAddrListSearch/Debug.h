#pragma once

#ifdef _DEBUG

#include <mapidefs.h>

// In debug mode, declare functions
void TraceSearchPath(IAddrBook &AddrBook);

#else

// In release mode, remove functionality
#define TraceSearchPath(x)

#endif
