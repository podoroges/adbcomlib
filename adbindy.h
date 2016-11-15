  //---------------------------------------------------------------------------
  // We are looking for pattern with length patlen in buf with length buflen
  // starting from start. Returning first occurence, or -1 if not found
  int memsearch(char * buf,int buflen,int start,char * pattern,int patlen)
  {
    int found;
    for(int a=start;a<buflen-patlen+1;a++)
    {
      found = 1;
      int b=0;
      while(b<patlen)
      {
        if(buf[a+b]!=pattern[b])
        {
          found = 0;
          b = patlen;
        }
        b++;
      }
      if(found)
        return a;
    }
    return -1;
  }


int CatchUpploadedFileEx(TIdHTTPRequestInfo *ARequestInfo,AnsiString fname){
      //Узнаем размер полного MIME-потока
      int size = ARequestInfo->PostStream->Size;
      //Создадим буфер, и скинем весь поток туда - так быстрее можно найти начало и конец
      char * buf = (char*)calloc(size+1,sizeof(char));
      ARequestInfo->PostStream->Position = 0;
      ARequestInfo->PostStream->ReadBuffer(buf,size);
      //Создадим буфер для MIMEboundary и прочитаем его
      int bound_length = memsearch(buf,size,0,"\r\n",2);
      if(bound_length<0)ErrorLog("CrewDocs Upload bound_length");
      char * bound = (char*)calloc(bound_length+1,sizeof(char));
      memcpy(bound,buf,bound_length);
      int p1 = memsearch(buf,size,0,"filename=",9); //Найдем "filename="
      p1 = memsearch(buf,size,p1,"\r\n\r\n",4); //Найдем начало файла: два CRLF подряд
      p1 = p1+4; // Пропустим их
      //Найдем конец файла по MIMEboundary
      int p2 = memsearch(buf,size,p1,bound,bound_length);
      if(p2<0)return 0;
      p2 = p2-2; // Вернемся назад на один CRLF
      //Откроем файл для записи
      FILE * f = fopen(fname.c_str(),"wb");
      if(!f)return 0;
      //Запишем файл на диск
      fwrite(buf+p1,sizeof(char),p2-p1,f);
      fclose(f);
      //Вернёмся в просмотр каталога
      free(buf);
      free(bound);
      return 1;  
  }
