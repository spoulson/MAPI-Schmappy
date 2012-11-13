#ifdef _DEBUG

#include <mapix.h>
#include <iostream>
#include <iomanip>
using namespace std;

void TraceDefaultDir(IAddrBook &AddrBook) {
   cout << "Default address list: ";
   ULONG cbEntryID;
   LPENTRYID lpEntryID;
   AddrBook.GetDefaultDir(&cbEntryID, &lpEntryID);
   for (ULONG i = 0; i < cbEntryID; i++) {
      cout << hex << setfill('0') << setw(2) << (unsigned int)((BYTE *)lpEntryID)[i];
   }
   cout << endl;
}

#endif