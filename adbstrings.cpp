#include "classes.hpp"
//---------------------------------------------------------------------------
//  ExtractString1("archive1.zip","arc",".zip") = "hive1"
AnsiString ExtractString1(AnsiString source,AnsiString start,AnsiString end){
  if(!source.Pos(start))return "";
  source = source.Delete(1,source.Pos(start)+start.Length()-1);
  if(!source.Pos(end))return source;
  source = source.Delete(source.Pos(end),source.Length());
  return source;
}
//---------------------------------------------------------------------------
// Replace in "st" all "from" to "to" (BCC version, copyright Never)
AnsiString StrRep(AnsiString st,AnsiString from,AnsiString to){
  while(st.Pos(from)){
    int a = st.Pos(from);
    st.Delete(a,from.Length());
    st.Insert(to,a);
  }
  return st;
}
//---------------------------------------------------------------------------
// Base64
// Source: http://fm4dd.com/programming/base64/base64_stringencode_c.htm
char b64[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
/* decodeblock - decode 4 '6-bit' characters into 3 8-bit binary bytes */
AnsiString decodeblock(unsigned char in[]) {
  unsigned char out[4];
  out[0] = in[0] << 2 | in[1] >> 4;
  out[1] = in[1] << 4 | in[2] >> 2;
  out[2] = in[2] << 6 | in[3] >> 0;
  out[3] = '\0';
  return AnsiString((char *)out);
}
AnsiString DecodeBase64(AnsiString src) {
  AnsiString st;
  char *b64src = src.c_str();
  int c, phase, i;
  unsigned char in[4];
  char *p;

  phase = 0; i=0;
  while(b64src[i]) {
    c = (int) b64src[i];
    if(c == '=') {
      st = (AnsiString)st+decodeblock(in);
      break;
    }
    p = strchr(b64, c);
    if(p) {
      in[phase] = p - b64;
      phase = (phase + 1) % 4;
      if(phase == 0) {
        st = (AnsiString)st+decodeblock(in);
        in[0]=in[1]=in[2]=in[3]=0;
      }
    }
    i++;
  }
  return st;
}
/* encodeblock - encode 3 8-bit binary bytes as 4 '6-bit' characters */
AnsiString encodeblock( unsigned char in[], int len ) {
    unsigned char out[5];
    out[0] = b64[ in[0] >> 2 ];
    out[1] = b64[ ((in[0] & 0x03) << 4) | ((in[1] & 0xf0) >> 4) ];
    out[2] = (unsigned char) (len > 1 ? b64[ ((in[1] & 0x0f) << 2) |
             ((in[2] & 0xc0) >> 6) ] : '=');
    out[3] = (unsigned char) (len > 2 ? b64[ in[2] & 0x3f ] : '=');
    out[4] = '\0';
    return AnsiString((char *)out);
}
/* encode - base64 encode a stream, adding padding if needed */
AnsiString b64_encode(char *clrstr, int cnt) {
  AnsiString st;
  unsigned char in[3];
  int i, len = 0;
  int j = 0;

  while(j<cnt) {
    len = 0;
    for(i=0; i<3; i++) {
     in[i] = (unsigned char) clrstr[j];
     if(j<cnt) {
        len++; j++;
      }
      else in[i] = 0;
    }
    if( len ) {
      st = (AnsiString) st+encodeblock( in, len );
    }
  }

  return st;
}

AnsiString EncodeBase64(AnsiString st){
  return b64_encode(st.c_str(), st.Length());
}

AnsiString EncodeBase64(TStream * s){
  int size = s->Size;
  char * buf = new char[size+1];
  s->ReadBuffer(buf,size);
  AnsiString st = b64_encode(buf, size);
  delete [] buf;
  return st;
}


// Make base64 strings url friendly
// http://stackoverflow.com/questions/1374753/passing-base64-encoded-strings-in-url
AnsiString UrlEncodeBase64(AnsiString st){
	st = EncodeBase64(st);
	st = StrRep(st,"+","-");
	st = StrRep(st,"/","_");
	return st;
}
AnsiString UrlDecodeBase64(AnsiString st){
	st = StrRep(st,"-","+");
	st = StrRep(st,"_","/");
	return DecodeBase64(st);
}
