unit Commune_APIUtilities;

{ ==============================================================================================================================
  Centralize Utilities for accessing the Windows API

  MIT License

  Copyright (c) 2020 Keith Latham

  Permission is hereby granted, free of charge, to any person obtaining a copy
  of this software and associated documentation files (the "Software"), to deal
  in the Software without restriction, including without limitation the rights
  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
  copies of the Software, and to permit persons to whom the Software is
  furnished to do so, subject to the following conditions:

  The above copyright notice and this permission notice shall be included in all
  copies or substantial portions of the Software.

  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
  SOFTWARE.

  ============================================================================================================================== }
{$IF DECLARED(FireMonkeyVersion)}
{$DEFINE HAS_FMX}
{$ELSE}
{$DEFINE HAS_VCL}
{$ENDIF}

interface

uses
  ActiveX,
  ComObj,
  ShlObj,
  System.Types,
  Winapi.Windows,
  Winapi.KnownFolders,
  System.Classes
{$IFDEF HAS_FMX}
    ,
  FMX.Types
{$ENDIF}
    ;

type
  dwOperationFlags = DWORD;

type
  TFileSystemBindData = class(TInterfacedObject, IFileSystemBindData)
    fw32fd: TWin32FindData;
    function GetFindData(var w32fd: TWin32FindData): HRESULT; stdcall;
    function SetFindData(var w32fd: TWin32FindData): HRESULT; stdcall;
  end;

type
  fsScaleFactor = (fss_Byte, fss_KiloByte, fss_Megabyte, fss_Gigabyte, fss_Terabyte, fss_Petabyte);

type
  tFileOperationProgressSink = class(TInterfacedObject, IFileOperationProgressSink)
    function FinishOperations(hrResult: HRESULT): HRESULT; virtual; stdcall;
    function PauseTimer: HRESULT; virtual; stdcall;
    function PostCopyItem(dwFlags: DWORD; const psiItem: IShellItem; const psiDestinationFolder: IShellItem;
      pszNewName: LPCWSTR; hrCopy: HRESULT; const psiNewlyCreated: IShellItem): HRESULT; virtual; stdcall;
    function PostDeleteItem(dwFlags: DWORD; const psiItem: IShellItem; hrDelete: HRESULT;
      const psiNewlyCreated: IShellItem): HRESULT; virtual; stdcall;
    function PostMoveItem(dwFlags: DWORD; const psiItem: IShellItem; const psiDestinationFolder: IShellItem;
      pszNewName: LPCWSTR; hrMove: HRESULT; const psiNewlyCreated: IShellItem): HRESULT; virtual; stdcall;
    function PostNewItem(dwFlags: DWORD; const psiDestinationFolder: IShellItem; pszNewName: LPCWSTR;
      pszTemplateName: LPCWSTR; dwFileAttributes: DWORD; hrNew: HRESULT; const psiNewItem: IShellItem): HRESULT;
      virtual; stdcall;
    function PostRenameItem(dwFlags: DWORD; const psiItem: IShellItem; pszNewName: LPCWSTR; hrRename: HRESULT;
      const psiNewlyCreated: IShellItem): HRESULT; virtual; stdcall;
    function PreCopyItem(dwFlags: DWORD; const psiItem: IShellItem; const psiDestinationFolder: IShellItem;
      pszNewName: LPCWSTR): HRESULT; virtual; stdcall;
    function PreDeleteItem(dwFlags: DWORD; const psiItem: IShellItem): HRESULT; virtual; stdcall;
    function PreMoveItem(dwFlags: DWORD; const psiItem: IShellItem; const psiDestinationFolder: IShellItem;
      pszNewName: LPCWSTR): HRESULT; virtual; stdcall;
    function PreNewItem(dwFlags: DWORD; const psiDestinationFolder: IShellItem; pszNewName: LPCWSTR): HRESULT;
      virtual; stdcall;
    function PreRenameItem(dwFlags: DWORD; const psiItem: IShellItem; pszNewName: LPCWSTR): HRESULT; virtual; stdcall;
    function ResetTimer: HRESULT; virtual; stdcall;
    function ResumeTimer: HRESULT; virtual; stdcall;
    function StartOperations: HRESULT; virtual; stdcall;
    function UpdateProgress(iWorkTotal: UINT; iWorkSoFar: UINT): HRESULT; virtual; stdcall;
  end;

type
  tAPIUtility = class
    private
      class function API_GetDriveType(aDriveLetter: pChar): integer;
      class function API_GetDriveTypeString(aDriveLetter: pChar): string;
      class function API_GetVolumeLabel(DriveChar: pChar): string;
      class function API_IFileOperation(const aFunction: cardinal; const SourceFile, TargetFile: string)
        : boolean; static;
      class function API_SHFileOperation(const aFunction: cardinal; const SourceFile, TargetFile: string;
        const aOwnerHandle: HWnd = 0): boolean;
      class procedure COM_FileOperation_Close; static;
      class function COM_FileOperation_INITCopy(aFileOp: IFileOperation; const aSourceFile, aTargetFile: string)
        : HRESULT; static;
      class function COM_FileOperation_INITDelete(aFileOp: IFileOperation; const aSourceFile: string): HRESULT; static;
      class function COM_FileOperation_INITMove(aFileOp: IFileOperation; const aSourceFile, aTargetFile: string)
        : HRESULT; static;
      class function COM_FileOperation_INITRename(aFileOp: IFileOperation; const aSourceFile, aTargetFile: string)
        : HRESULT; static;
      class function COM_FileOperation_Open(const aCoInit: Longint; const aOPFlags: dwOperationFlags;
        out aHresult: HRESULT): IFileOperation; static;
      class function COM_FileOperation_SetSource(const aSourceFile, aTargetDir: string; out pbc: IBindCtx;
        out aHresult: HRESULT): IShellItem; static;
      class function COM_FileOperation_SetTarget(const aTargetDir: string; pbc: IBindCtx; out aHresult: HRESULT)
        : IShellItem; static;
    public
      class function EnumerateAllDrives: tStringList;
      class function EnumerateDirectories(const aRoot: string): tStringList;
      class function EnumerateLogicalDrives: tStringList;
      class function EnumeratePathFiles(aPath: string; aFilter: string = '*.*'): tStringList;
      class function FileCopy(const aFrom, aTo: string): boolean;
      class function FileCopyX(const aFrom, aTo: string): boolean; static;
      class function FileDelete(const aFrom: string): boolean;
      class function FileDeleteX(const aFrom: string): boolean;
      class function FileMove(const aFrom, aTo: string): boolean;
      class function FileRename(const aFrom, aTo: string): boolean;
      class function FileSizeOf(aFile: string; aScale: fsScaleFactor = fss_Byte): nativeint;
      class function KnownFolderPath(const aKnownFolderCSIDL: integer): string; overload;
      class function KnownFolderPath(const aKnownFolderGUID: tGUID): string; overload;
      class function SanitizeFilename(aInputFileName: string): string;
  end;

type
  tKnownFolder = class
    public
      class function ProgramDataPath: string;
  end;

implementation

uses
  // CodeSiteLogging,
  System.StrUtils,
  System.SysUtils,
  // JclAnsiStrings,
  Winapi.ShellAPI,
  FMX.Platform.Win,
  System.IOUtils,
  System.Math;

{ tAPIUtility }

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.API_GetDriveType(aDriveLetter: pChar): integer;
var
  TargetDrive: pChar;
begin
  TargetDrive := pChar(aDriveLetter + ':\');
  result      := GetDriveType(TargetDrive);
  // Codesite.Send('API_GetDriveType ' + aDriveLetter, Result);
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.API_GetDriveTypeString(aDriveLetter: pChar): string;
var
  myDriveType: integer;
begin
  myDriveType := API_GetDriveType(aDriveLetter);
  case myDriveType of
    DRIVE_UNKNOWN: result     := 'UNKNOWN';
    DRIVE_NO_ROOT_DIR: result := 'NO_ROOT_DIR';
    DRIVE_REMOVABLE: result   := 'REMOVABLE';
    DRIVE_FIXED: result       := 'FIXED';
    DRIVE_REMOTE: result      := 'REMOTE';
    DRIVE_CDROM: result       := 'CDROM';
    DRIVE_RAMDISK: result     := 'RAMDISK';
  end;
  // Codesite.Send('API_GetDriveTypeString ' + aDriveLetter, Result);
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.API_GetVolumeLabel(DriveChar: pChar): string;
var
  Buf: array [0 .. MAX_PATH] of Char;
  NotUsed: DWORD;
  VolumeFlags: DWORD;
  VolumeInfo: array [0 .. MAX_PATH] of Char;
  VolumeSerialNumber: DWORD;
begin
  VolumeInfo[0] := 'C'; // do something nonsensical to keep the compiler happy
  GetVolumeInformation(pChar(DriveChar + ':\'), Buf, SizeOf(VolumeInfo), @VolumeSerialNumber, NotUsed,
    VolumeFlags, nil, 0);
  SetString(result, Buf, StrLen(Buf)); { Set return result }
  result := AnsiUpperCase(result);
  // Codesite.Send('API_GetVolumeLabel ' + DriveChar, Result);
end;

class function tAPIUtility.API_IFileOperation(const aFunction: cardinal; const SourceFile, TargetFile: string): boolean;
{ ----------------------------------------------------------------------------------------------------------------------------- }

const
  myApartment: shortint = (COINIT_APARTMENTTHREADED or COINIT_DISABLE_OLE1DDE);
  // Yes to all for any dialog box  // Do NOT allow progress box to be minimized
  myOPFlags: dwOperationFlags = (FOF_NOCONFIRMATION or FOFX_NOMINIMIZEBOX or FOFX_DONTDISPLAYSOURCEPATH or
    FOF_NOCONFIRMMKDIR);
var
  fileOp: IFileOperation;
  r: HRESULT;
begin
  result := false;

  fileOp := COM_FileOperation_Open(myApartment, myOPFlags, r); // CoInitializeEx;

  if succeeded(r) then
    begin
      // set up the copy operation
      case aFunction of
        FO_COPY: r   := COM_FileOperation_INITCopy(fileOp, SourceFile, TargetFile);
        FO_MOVE: r   := COM_FileOperation_INITMove(fileOp, SourceFile, TargetFile);
        FO_DELETE: r := COM_FileOperation_INITDelete(fileOp, SourceFile);
        FO_RENAME: r := COM_FileOperation_INITRename(fileOp, SourceFile, TargetFile);
      else
        begin
        end;
      end;
      // execute
      if succeeded(r) then r := fileOp.PerformOperations;
      result                 := succeeded(r);
      OleCheck(r);
    end;

  COM_FileOperation_Close; // CoUninitialize;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.API_SHFileOperation(const aFunction: cardinal; const SourceFile, TargetFile: string;
  const aOwnerHandle: HWnd = 0): boolean;
{ ----------------------------------------------------------------------------------------------------------------------------- }

// FMX callers will need to use FormToHWND(aOwnerHandle) to get the HWnd;
var
  Aborted: boolean;
  SHFoSInfo: TSHFileOpStruct;
  Source: string;
  Target: string;
begin
  Aborted := false;
  Source  := SourceFile + #0;
  Target  := TargetFile + #0;
  with SHFoSInfo do
    begin
      Wnd := aOwnerHandle;
      // Wnd := hWndOwner;
      // Wnd := FmxHandleToHWND(aOwnerHandle);
      { From Microsoft's Help:
        wFunc = Operation to perform. This member can be one of the following values:
        FO_COPY Copies the files specified by pFrom to the location specified by pTo.
        FO_DELETE Deletes the files specified by pFrom (pTo is ignored).
        FO_MOVE Moves the files specified by pFrom to the location specified by pTo.
        FO_RENAME Renames the files specified by pFrom. }

      wFunc                       := aFunction;
      pFrom                       := pWideChar(Source);
      pTo                         := pWideChar(Target);
      fFlags                      := FOF_NOCONFIRMMKDIR; // AND FOF_NOCONFIRMATION;
      fAnyOperationsAborted       := Aborted;
      SHFoSInfo.lpszProgressTitle := pFrom;
    end;
  try
    SHFileOperation(SHFoSInfo);
  finally
    result := not Aborted;
  end;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class procedure tAPIUtility.COM_FileOperation_Close;
{ ======= }
begin
  CoUninitialize;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.COM_FileOperation_INITCopy(aFileOp: IFileOperation;
  const aSourceFile, aTargetFile: string): HRESULT;
{ ======= }
var
  destFileFolder, destFileName: string;
  siSrcFile: IShellItem;
  siDestFolder: IShellItem;
  pbc: IBindCtx;
begin
  destFileFolder := ExtractFileDir(aTargetFile);
  destFileName   := ExtractFileName(aTargetFile);
  // get source shell item
  siSrcFile := COM_FileOperation_SetSource(aSourceFile, destFileFolder, pbc, result);
  // get destination folder shell item
  siDestFolder := COM_FileOperation_SetTarget(destFileFolder, pbc, result);
  // add copy operation to file operation list
  if succeeded(result) then result := aFileOp.CopyItem(siSrcFile, siDestFolder, pChar(destFileName), nil);
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.COM_FileOperation_INITDelete(aFileOp: IFileOperation; const aSourceFile: string): HRESULT;
{ ======= }
var
  siSrcFile: IShellItem;
  pbc: IBindCtx;
begin
  // get source shell item
  siSrcFile := COM_FileOperation_SetSource(aSourceFile, '', pbc, result);
  // add copy operation to file operation list
  if succeeded(result) then result := aFileOp.DeleteItem(siSrcFile, nil);
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.COM_FileOperation_INITMove(aFileOp: IFileOperation;
  const aSourceFile, aTargetFile: string): HRESULT;
begin
  { TODO : Implement this routine }
  result := E_NOTIMPL;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.COM_FileOperation_INITRename(aFileOp: IFileOperation;
  const aSourceFile, aTargetFile: string): HRESULT;
begin
  { TODO : Implement this routine }
  result := E_NOTIMPL;
end;

// class function tAPIUtility.API_IFileOperation(const aFunction: cardinal; const SourceFile, TargetFile: string;
// const aOwnerHandle: HWnd = 0): boolean;

/// / begin
/// /
/// / function CopyFileIFileOperationForceDirectories(const srcFile, destFile : string) : boolean;
/// / works on Windows >= Vista and 2008 server
// var
// r                           : HRESULT;
// fileOp                      : IFileOperation;
// siSrcFile                   : IShellItem;
// siDestFolder                : IShellItem;
// destFileFolder, destFileName: string;
// pbc                         : IBindCtx;
// w32fd                       : TWin32FindData;
// ifs                         : TFileSystemBindData;
// begin
// result := false;
//
// destFileFolder := ExtractFileDir(TargetFile);
// destFileName   := ExtractFileName(TargetFile);
//
// // init com
// r := CoInitializeEx(nil, COINIT_APARTMENTTHREADED or COINIT_DISABLE_OLE1DDE);
// if Succeeded(r) then begin
// // create IFileOperation interface
// r := CoCreateInstance(CLSID_FileOperation, nil, CLSCTX_ALL, IFileOperation, fileOp);
// if Succeeded(r) then begin
// // set operations flags
// r := fileOp.SetOperationFlags(FOF_NOCONFIRMATION or FOFX_NOMINIMIZEBOX);
// if Succeeded(r) then begin
// // get source shell item
// r := SHCreateItemFromParsingName(pChar(SourceFile), nil, IShellItem, siSrcFile);
// if Succeeded(r) then begin
// // create binding context to pretend there is a folder there
// if not DirectoryExists(destFileFolder) then begin
// ZeroMemory(@w32fd, SizeOf(TWin32FindData));
// w32fd.dwFileAttributes := FILE_ATTRIBUTE_DIRECTORY;
// ifs                    := TFileSystemBindData.Create;
// ifs.SetFindData(w32fd);
// r := CreateBindCtx(0, pbc);
// r := pbc.RegisterObjectParam(STR_FILE_SYS_BIND_DATA, ifs);
// end
// else pbc := nil;
//
// // get destination folder shell item
// r := SHCreateItemFromParsingName(pChar(destFileFolder), pbc, IShellItem, siDestFolder);
//
// // add copy operation
// if Succeeded(r) then r := fileOp.CopyItem(siSrcFile, siDestFolder, pChar(destFileName), nil);
// end;
//
// // execute
// if Succeeded(r) then r := fileOp.PerformOperations;
//
// result := Succeeded(r);
//
// OleCheck(r);
// end;
// end;
//
// CoUninitialize;
// end;
// end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.COM_FileOperation_Open(const aCoInit: Longint; const aOPFlags: dwOperationFlags;
  out aHresult: HRESULT): IFileOperation;
{ ======= }
begin
  result := nil;
  // initialize COM
  aHresult := CoInitializeEx(nil, aCoInit);
  if (aHresult <> RPC_E_CHANGED_MODE) then
    begin
      // create IFileOperation interface
      aHresult := CoCreateInstance(CLSID_FileOperation, // we are doing a file operation
        nil,                                            // not an aggregate
        CLSCTX_ALL,                                     // any execution context
        IFileOperation,                                 // interface class
        result);                                        // interface object result
    end;
  // set operations flags
  if succeeded(aHresult) then aHresult := result.SetOperationFlags(aOPFlags);
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.COM_FileOperation_SetSource(const aSourceFile, aTargetDir: string; out pbc: IBindCtx;
  out aHresult: HRESULT): IShellItem;
{ ======= }
var
  w32fd: TWin32FindData;
  ifs: TFileSystemBindData;
begin
  aHresult := SHCreateItemFromParsingName(pChar(aSourceFile), // source file (directory?) name
    nil,                                                      // no bind context (yet)
    IShellItem,                                               // we want a shell item interface
    result);                                                  // resulting source file shell item interface
  if succeeded(aHresult) then
    begin
      // create binding context to pretend there is a folder there
      if ((length(aTargetDir) > 0) and (not DirectoryExists(aTargetDir))) then
        begin
          ZeroMemory(@w32fd, SizeOf(TWin32FindData));
          w32fd.dwFileAttributes := FILE_ATTRIBUTE_DIRECTORY;
          ifs                    := TFileSystemBindData.Create;
          ifs.SetFindData(w32fd);
          aHresult := CreateBindCtx(0, pbc);
          aHresult := pbc.RegisterObjectParam(STR_FILE_SYS_BIND_DATA, ifs);
        end
      else pbc := nil;
    end;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.COM_FileOperation_SetTarget(const aTargetDir: string; pbc: IBindCtx; out aHresult: HRESULT)
  : IShellItem;
{ ======= }
begin
  aHresult := SHCreateItemFromParsingName(pChar(aTargetDir), pbc, IShellItem, result);
end;

class function tAPIUtility.EnumerateAllDrives: tStringList;
{ ----------------------------------------------------------------------------------------------------------------------------- }
var
  DriveList: tStringList;
  DriveString: string;
  myDriveChar: pChar;
  myDriveString: string;
  myDriveTarget: string;
  myDriveType: integer;
  myVolumeLabel: string;
begin
  result    := tStringList.Create;
  DriveList := EnumerateLogicalDrives;
  try
    for DriveString in DriveList do
      begin
        myDriveChar := pChar(DriveString);
        myDriveType := API_GetDriveType(myDriveChar);
        if (myDriveType > 0) then
          begin
            myDriveTarget := DriveString + ':';
            myDriveString := API_GetDriveTypeString(myDriveChar);
            myVolumeLabel := API_GetVolumeLabel(myDriveChar);
            result.Add(myDriveTarget + '=<' + myVolumeLabel + '> (' + myDriveString + ')');
          end;
      end;
  finally
    DriveList.Free;
  end;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.EnumerateDirectories(const aRoot: string): tStringList;

  procedure HandleFind(asr: tSearchRec; aresult: tStringList);
  begin
    if ((asr.Attr and faDirectory) = faDirectory) and (asr.Name <> '.') and (asr.Name <> '..') then
        aresult.Add(asr.Name);
  end;

var
  sr: tSearchRec;
begin
  result := tStringList.Create;
  if FindFirst(aRoot + '*', faDirectory, sr) = 0 then
    repeat HandleFind(sr, result);
    until FindNext(sr) <> 0;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.EnumerateLogicalDrives: tStringList;
var
  i: integer;
  length: integer;
  MyStr: pChar;
const
  Size: integer = 200;
begin
  result := tStringList.Create;
  GetMem(MyStr, Size);
  length := GetLogicalDriveStrings(Size, MyStr);
  for i  := 0 to length - 1 do
    if (MyStr[i] >= 'A') and (MyStr[i] <= 'Z') then result.Add(MyStr[i]);
  FreeMem(MyStr);
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.EnumeratePathFiles(aPath: string; aFilter: string = '*.*'): tStringList;
var
  thisFile: string;
begin
  result := tStringList.Create;
  for thisFile in TDirectory.GetFiles(aPath, aFilter, tSearchOption.soTopDirectoryOnly) do
    begin
      result.Add(thisFile);
    end;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.FileCopy(const aFrom, aTo: string): boolean;
begin
  // codesite.Send('(Redirected to FileCopyX) FileCopy SOURCE: ' + aFrom + ' ==> TARGET: ' + aTo);
  // result := tAPIUtility.API_SHFileOperation(FO_Copy, aFrom, aTo, 0);
  result := tAPIUtility.FileCopyX(aFrom, aTo);
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.FileCopyX(const aFrom, aTo: string): boolean;
{ class function API_IFileOperation(const aFunction: cardinal; const SourceFile, TargetFile: string): boolean; static;
}
begin
  // codesite.Send('FileCopyX SOURCE: ' + aFrom + ' ==> TARGET: ' + aTo);
  result := tAPIUtility.API_IFileOperation(FO_COPY, aFrom, aTo);
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.FileDelete(const aFrom: string): boolean;
begin
  // codesite.Send('(Redirected to FileDeleteX) FileDelete SOURCE: ' + aFrom);
  // result := tAPIUtility.API_SHFileOperation(FO_DELETE, aFrom, aTo, 0);
  result := tAPIUtility.FileDeleteX(aFrom);
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.FileDeleteX(const aFrom: string): boolean;
begin
  result := tAPIUtility.API_IFileOperation(FO_DELETE, aFrom, '');
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.FileMove(const aFrom, aTo: string): boolean;
begin
  result := tAPIUtility.API_SHFileOperation(FO_MOVE, aFrom, aTo, 0);
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.FileRename(const aFrom, aTo: string): boolean;
begin
  result := tAPIUtility.API_SHFileOperation(FO_RENAME, aFrom, aTo, 0);
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.FileSizeOf(aFile: string; aScale: fsScaleFactor = fss_Byte): nativeint;
var
  bytecount: nativeint;
  F: file of byte;
  scaled: extended;
  scaler: extended;
begin
  AssignFile(F, aFile);
  reset(F);
  try
    bytecount := FileSize(F);
  finally
    closefile(F);
  end;
  scaler := IntPower(1024, ord(aScale));
  scaled := bytecount / scaler;
  result := ceil(scaled);
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.KnownFolderPath(const aKnownFolderCSIDL: integer): string;
var
  myHandle: tHandle;
  myToken: cardinal;
  myFlags: cardinal;
  mypszPath: array [0 .. MAX_PATH] of Char;
  APIResult: integer;
begin
  myHandle  := 0;
  myToken   := 0;
  myFlags   := 0;
  APIResult := SHGetFolderPath(myHandle, aKnownFolderCSIDL, myToken, myFlags, mypszPath);
  result    := mypszPath;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.KnownFolderPath(const aKnownFolderGUID: tGUID): string;
// REMEMBER: add Winapi.ShlObj to USES
var
  path: LPWSTR;
begin
  if succeeded(SHGetKnownFolderPath(aKnownFolderGUID, 0, 0, path)) then
    begin
      try
        result := path;
      finally
        CoTaskMemFree(path);
      end;
    end
  else result := '';
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tAPIUtility.SanitizeFilename(aInputFileName: string): string;
var
  ANSI: AnsiString;
begin
  result := aInputFileName;
  result := AnsiReplaceStr(result, '/', '_');
  result := AnsiReplaceStr(result, '\', '_');
  result := AnsiReplaceStr(result, '"', '''');
  result := AnsiReplaceStr(result, '#', 'No.');
  result := AnsiReplaceStr(result, '%23', 'No.');
  ANSI   := AnsiString(result);
  result := ANSI;
  // If you have access to the JEDI JCL library, then StrSmartCase makes the filename look pretty
  // you will have to uncomment the "JCLAnsiStrings" line in the implementation uses clause too
  // result := string(StrSmartCase(ANSI, [' ', #9, '\', #13, #10]));
end;

{ TFileSystemBindData }

{ ----------------------------------------------------------------------------------------------------------------------------- }
function TFileSystemBindData.GetFindData(var w32fd: TWin32FindData): HRESULT;
begin
  w32fd  := fw32fd;
  result := S_OK;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
function TFileSystemBindData.SetFindData(var w32fd: TWin32FindData): HRESULT;
begin
  fw32fd := w32fd;
  result := S_OK;
end;

{ tFileOperationProgressSink }

const
  FileOperationSuccessfull: HRESULT = 0;

  { ----------------------------------------------------------------------------------------------------------------------------- }
function tFileOperationProgressSink.FinishOperations(hrResult: HRESULT): HRESULT;
begin
  result := FileOperationSuccessfull;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
function tFileOperationProgressSink.PauseTimer: HRESULT;
begin
  result := FileOperationSuccessfull;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
function tFileOperationProgressSink.PostCopyItem(dwFlags: DWORD; const psiItem, psiDestinationFolder: IShellItem;
  pszNewName: LPCWSTR; hrCopy: HRESULT; const psiNewlyCreated: IShellItem): HRESULT;
begin
  result := FileOperationSuccessfull;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
function tFileOperationProgressSink.PostDeleteItem(dwFlags: DWORD; const psiItem: IShellItem; hrDelete: HRESULT;
  const psiNewlyCreated: IShellItem): HRESULT;
begin
  result := FileOperationSuccessfull;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
function tFileOperationProgressSink.PostMoveItem(dwFlags: DWORD; const psiItem, psiDestinationFolder: IShellItem;
  pszNewName: LPCWSTR; hrMove: HRESULT; const psiNewlyCreated: IShellItem): HRESULT;
begin
  result := FileOperationSuccessfull;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
function tFileOperationProgressSink.PostNewItem(dwFlags: DWORD; const psiDestinationFolder: IShellItem;
  pszNewName, pszTemplateName: LPCWSTR; dwFileAttributes: DWORD; hrNew: HRESULT; const psiNewItem: IShellItem): HRESULT;
begin
  result := FileOperationSuccessfull;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
function tFileOperationProgressSink.PostRenameItem(dwFlags: DWORD; const psiItem: IShellItem; pszNewName: LPCWSTR;
  hrRename: HRESULT; const psiNewlyCreated: IShellItem): HRESULT;
begin
  result := FileOperationSuccessfull;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
function tFileOperationProgressSink.PreCopyItem(dwFlags: DWORD; const psiItem, psiDestinationFolder: IShellItem;
  pszNewName: LPCWSTR): HRESULT;
begin
  result := FileOperationSuccessfull;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
function tFileOperationProgressSink.PreDeleteItem(dwFlags: DWORD; const psiItem: IShellItem): HRESULT;
begin
  result := FileOperationSuccessfull;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
function tFileOperationProgressSink.PreMoveItem(dwFlags: DWORD; const psiItem, psiDestinationFolder: IShellItem;
  pszNewName: LPCWSTR): HRESULT;
begin
  result := FileOperationSuccessfull;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
function tFileOperationProgressSink.PreNewItem(dwFlags: DWORD; const psiDestinationFolder: IShellItem;
  pszNewName: LPCWSTR): HRESULT;
begin
  result := FileOperationSuccessfull;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
function tFileOperationProgressSink.PreRenameItem(dwFlags: DWORD; const psiItem: IShellItem;
  pszNewName: LPCWSTR): HRESULT;
begin
  result := FileOperationSuccessfull;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
function tFileOperationProgressSink.ResetTimer: HRESULT;
begin
  result := FileOperationSuccessfull;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
function tFileOperationProgressSink.ResumeTimer: HRESULT;
begin
  result := FileOperationSuccessfull;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
function tFileOperationProgressSink.StartOperations: HRESULT;
begin
  result := FileOperationSuccessfull;
end;

{ ----------------------------------------------------------------------------------------------------------------------------- }
function tFileOperationProgressSink.UpdateProgress(iWorkTotal, iWorkSoFar: UINT): HRESULT;
begin
  { TODO : }
  result := E_NOTIMPL;
end;

{ tKnownFolder }

{ ----------------------------------------------------------------------------------------------------------------------------- }
class function tKnownFolder.ProgramDataPath: string;
begin
  { DONE : move this to commune utilities }
  result := tAPIUtility.KnownFolderPath(FOLDERID_ProgramData);
end;

end.
