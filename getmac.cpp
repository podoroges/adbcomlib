#include <classes.hpp>
#include <iphlpapi.h>

#pragma comment (lib, "iphlpapi.lib")

AnsiString GetMAC(){
  AnsiString st;
  IP_ADAPTER_INFO AdapterInfo[16];
  DWORD dwBufLen = sizeof(AdapterInfo);
  DWORD dwStatus = GetAdaptersInfo(AdapterInfo,&dwBufLen);
  if(dwStatus != ERROR_SUCCESS)
    return "GetAdaptersInfo Error";
  TStringList * str = new TStringList;
  PIP_ADAPTER_INFO pAdapterInfo = AdapterInfo;
  do{
    for(unsigned int i=0;i<pAdapterInfo->AddressLength;i++){
      if(st.Length())
        st = (AnsiString)st+":";
      st = (AnsiString)st+AnsiString().sprintf("%.2X",(int) pAdapterInfo->Address[i]);
    }
    str->Add(st);
    pAdapterInfo = pAdapterInfo->Next;
  }
  while(pAdapterInfo);
  st = str->Text.Trim();
  delete str;
  return st;
}

