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
// Конвертируем 4-символьные 6-битные в 3-символьные 8-битные группы
// Большая часть кода позаимствована отсюда:
// http://base64.sourceforge.net/b64.c
// DecodeBase64("RndkOiDQ0snXxdTJy8kpKQ==") = "Fwd: приветики))"
AnsiString DecodeBase64(AnsiString st)
{
  const char cd64[]="|$$$}rstuvwxyz{$$$$$$$>?@ABCDEFGHIJKLMNOPQRSTUVW$$$$$$XYZ[\\]^_`abcdefghijklmnopq";
	                  
  AnsiString st2;
  unsigned char in[4],out[3],v;
  int i,len,j = 0;
  while(j<st.Length()){
    for(len=0,i=0;i<4&&(j<st.Length());i++){
      v=0;
      while((j<st.Length())&&(v==0)){
        v=(unsigned char)st[++j];
        v=(unsigned char)((v<43||v>122)?0:cd64[v-43]);
        if(v)
          v=(unsigned char)((v=='$')?0:v-61);
      }
      if(j<st.Length()){
        len++;
        if(v)
          in[i]=(unsigned char)(v-1);
      }
      else{
        in[i]=0;
      }
    }
    if(len){
      out[0]=(unsigned char)(in[0]<<2|in[1]>>4);
      out[1]=(unsigned char)(in[1]<<4|in[2]>>2);
      out[2]=(unsigned char)(((in[2]<<6)&0xc0)|in[3]);
      for(i=0;i<len-1;i++)
        st2 = st2+(char)out[i];
    }
  }
  return st2;
}

AnsiString EncodeBase64(AnsiString st){
  static const char cb64[]="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  AnsiString st2;
  int i = 0;
  int j = 0;
  unsigned char char_array_3[3];
  unsigned char char_array_4[4];
  int k = 0;

  while (++k<=st.Length()) {
    char_array_3[i++] = st[k];
    if (i == 3) {
      char_array_4[0] = (char_array_3[0] & 0xfc) >> 2;
      char_array_4[1] = ((char_array_3[0] & 0x03) << 4) + ((char_array_3[1] & 0xf0) >> 4);
      char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) + ((char_array_3[2] & 0xc0) >> 6);
      char_array_4[3] = char_array_3[2] & 0x3f;

      for(i = 0; (i <4) ; i++)
        st2 = st2+cb64[char_array_4[i]];
      i = 0;
    }
  }

  if (i)
  {
    for(j = i; j < 3; j++)
      char_array_3[j] = '\0';

    char_array_4[0] = (char_array_3[0] & 0xfc) >> 2;
    char_array_4[1] = ((char_array_3[0] & 0x03) << 4) + ((char_array_3[1] & 0xf0) >> 4);
    char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) + ((char_array_3[2] & 0xc0) >> 6);
    char_array_4[3] = char_array_3[2] & 0x3f;

    for (j = 0; (j < i + 1); j++)
      st2 = st2+cb64[char_array_4[j]];

    while((i++ < 3))
     st2 = st2+'=';

  }

  return st2;
}


AnsiString EncodeBase64(TStream * s){
  static const char cb64[]="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  AnsiString st2;
  int i = 0;
  int j = 0;
  unsigned char char_array_3[3];
  unsigned char char_array_4[4];
  s->Position = 0;
  while (s->Position<s->Size) {
    s->ReadBuffer(&char_array_3[i++] ,1);
    if (i == 3) {
      char_array_4[0] = (char_array_3[0] & 0xfc) >> 2;
      char_array_4[1] = ((char_array_3[0] & 0x03) << 4) + ((char_array_3[1] & 0xf0) >> 4);
      char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) + ((char_array_3[2] & 0xc0) >> 6);
      char_array_4[3] = char_array_3[2] & 0x3f;
      for(i = 0; (i <4) ; i++)
        st2 = st2+cb64[char_array_4[i]];
      i = 0;
    }
  }
  if (i)
  {
    for(j = i; j < 3; j++)
      char_array_3[j] = '\0';
    char_array_4[0] = (char_array_3[0] & 0xfc) >> 2;
    char_array_4[1] = ((char_array_3[0] & 0x03) << 4) + ((char_array_3[1] & 0xf0) >> 4);
    char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) + ((char_array_3[2] & 0xc0) >> 6);
    char_array_4[3] = char_array_3[2] & 0x3f;
    for (j = 0; (j < i + 1); j++)
      st2 = st2+cb64[char_array_4[j]];
    while((i++ < 3))
     st2 = st2+'=';
  }
  return st2;
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
