#ifdef _DEBUG

#include <mapix.h>
#include <mapidefs.h>
#include <iostream>
#include <sstream>
using namespace std;

// Get string representation of SPropValue
static void ToString(const SPropValue &pv, string &OutString) {
   stringstream ss;
   switch (PROP_TYPE(pv.ulPropTag)) {
   case PT_SHORT:
      ss << pv.Value.i;
      break;
   case PT_LONG:
      ss << pv.Value.l;
      break;
   case PT_FLOAT:
      ss << pv.Value.flt;
      break;
   case PT_DOUBLE:
      ss << pv.Value.dbl;
      break;
   case PT_BOOLEAN:
      ss << pv.Value.b;
      break;
   case PT_STRING8:
      ss << pv.Value.lpszA;
      break;
   case PT_UNICODE:
      ss << pv.Value.lpszW;
      break;
   case PT_BINARY:
      ss.setf(ios::hex, ios::basefield);
      ss.width(2);
      ss.fill('0');
      for (ULONG i = 0; i < pv.Value.bin.cb; i++) {
         ss << (int)pv.Value.bin.lpb[i];
      }
      break;
   default:
      break;
   }

   OutString = ss.str();
}

// Show dump of LPADRBOOK structure
void TraceSearchPath(IAddrBook &AddrBook) {
   LPSRowSet lpAddrSearchPath = NULL;
   HRESULT hr = AddrBook.GetSearchPath(0, &lpAddrSearchPath);

   cout << "Trace AddressBook list:" << endl;

   for (ULONG i = 0; i < lpAddrSearchPath->cRows; i++) {
      SRow &row = lpAddrSearchPath->aRow[i];
      cout << "SRow.cValues: " << row.cValues << endl;
      string Value;
      ToString(*row.lpProps, Value);
      cout << "SRow.SPropValue: " << Value << endl;
   }
}

#endif
