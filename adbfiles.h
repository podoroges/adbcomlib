class CSearchFiles // Win version
{
  public:
  TStringList * List;
  CSearchFiles(AnsiString path);
  ~CSearchFiles();
};
CSearchFiles::CSearchFiles(AnsiString path){
  List = new TStringList();
  WIN32_FIND_DATA ffd;
  HANDLE hFind = FindFirstFile(path.c_str(), &ffd);
  do{
    AnsiString st = ffd.cFileName;
    if((st!=AnsiString(".."))&&(st!=AnsiString(".")))
      List->Add(st);
  }
  while (FindNextFile(hFind, &ffd) != 0);
  FindClose(hFind);
}
CSearchFiles::~CSearchFiles(){
  delete List;
}
