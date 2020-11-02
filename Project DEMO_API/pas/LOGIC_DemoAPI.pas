unit LOGIC_DemoAPI;

interface

uses
  System.Classes,
  Vcl.StdCtrls,
  Commune_APIUtilities,
  Winapi.Windows;

type
  tDEMO = class
    private
      class function UTIL_ByteCountScaled(aByteCount: single; aScale: fsScaleFactor = fss_Byte): single; static;
      class function UTIL_DiskFree(aDriveLetter: char): Int64; static;
      class function UTIL_DiskSize(aDriveLetter: char): Int64; static;
      class function UTIL_DriveLetter(aLetter: string): char; static;
      class function UTIL_ExtractDriveLetter(aDriveCombo: tComboBox; aIndex: integer): char; static;
      class function UTIL_SYSUTIL_InternalGetDiskSpace(Drive: char; var TotalSpace, FreeSpaceAvailable: Int64)
        : Bool; static;
    public
      class procedure INFO_AllDrives(aDriveCombo: tComboBox; aMemo: tMemo); static;
      class function INFO_CapacityOf(aDriveLetter: char): string; static;
      class function INFO_OneDrive(aDriveLetter: char): string; static;
      class function INFO_UnusedSpaceOn(aDriveLetter: char): string; static;
      class procedure LOAD_Combo(aCombo: tComboBox; aStringlist: tStringList); static;
      class procedure LOAD_Memo(aMemo: tMemo; aStringlist: tStringList); static;
  end;

implementation

uses
  System.SysUtils,
  math;

const
  GLOBAL_SCALAR     = fss_Gigabyte;
  GLOBAL_SCALARTEXT = 'GB';

  { tDEMO }

class procedure tDEMO.INFO_AllDrives(aDriveCombo: tComboBox; aMemo: tMemo);
{ --------------------------------------------------------------------------------------------------------------------------- }
var
  thisDriveString: string;
  myDrive: char;
  myDetails: string;
begin
  aMemo.Clear;
  for thisDriveString in aDriveCombo.Items do
    begin
      myDrive   := UTIL_DriveLetter(thisDriveString);
      myDetails := INFO_OneDrive(myDrive);
      aMemo.Lines.Add(thisDriveString + ' ' + myDetails);
    end;
end;

class function tDEMO.INFO_CapacityOf(aDriveLetter: char): string;
{ --------------------------------------------------------------------------------------------------------------------------- }
var
  myCapacity: Int64;
  myscaled: single;
  mystring: string;
begin
  myCapacity := UTIL_DiskSize(aDriveLetter);
  myscaled   := UTIL_ByteCountScaled(myCapacity, GLOBAL_SCALAR);
  mystring   := floattostr(myscaled);
  result     := 'Capacity=' + mystring + GLOBAL_SCALARTEXT;
end;

class function tDEMO.INFO_OneDrive(aDriveLetter: char): string;
{ --------------------------------------------------------------------------------------------------------------------------- }
var
  theCapacity: string;
  theUnused: string;
begin
  theCapacity := INFO_CapacityOf(aDriveLetter);
  theUnused   := INFO_UnusedSpaceOn(aDriveLetter);
  result      := theCapacity + ' ' + theUnused;
end;

class function tDEMO.INFO_UnusedSpaceOn(aDriveLetter: char): string;
{ --------------------------------------------------------------------------------------------------------------------------- }
var
  myDASDIndex: integer;
  myFreeSpace: Int64;
  myscaled: single;
  mystring: string;

begin
  myFreeSpace := UTIL_DiskFree(aDriveLetter);
  myscaled    := UTIL_ByteCountScaled(myFreeSpace, GLOBAL_SCALAR);
  mystring    := floattostr(myscaled);
  result      := 'Available= ' + mystring + GLOBAL_SCALARTEXT;
end;

class procedure tDEMO.LOAD_Combo(aCombo: tComboBox; aStringlist: tStringList);
{ --------------------------------------------------------------------------------------------------------------------------- }
var
  myDrive: string;
begin
  aCombo.Clear;
  for myDrive in aStringlist do
    begin
      aCombo.Items.Add(myDrive);
    end;
  aCombo.ItemIndex                                := -1;
  if aCombo.Items.Count > 0 then aCombo.ItemIndex := 0;
end;

class procedure tDEMO.LOAD_Memo(aMemo: tMemo; aStringlist: tStringList);
{ --------------------------------------------------------------------------------------------------------------------------- }
var
  myDrive: string;
begin
  aMemo.Clear;
  for myDrive in aStringlist do
    begin
      aMemo.Lines.Add(myDrive);
    end;
end;

class function tDEMO.UTIL_ByteCountScaled(aByteCount: single; aScale: fsScaleFactor = fss_Byte): single;
{ ---------------------------------------------------------------------------------------------------------------------------- }
var
  scaler: extended;
  scaled: extended;
begin
  scaler := IntPower(1024, ord(aScale));
  scaled := aByteCount / scaler;
  result := ceil(scaled);
end;

class function tDEMO.UTIL_DiskFree(aDriveLetter: char): Int64;
{ --------------------------------------------------------------------------------------------------------------------------- }
var
  mySize, myFree: Int64;
  theDrive: char;
begin
  if not UTIL_SYSUTIL_InternalGetDiskSpace(aDriveLetter, mySize, myFree) then result := -1
  else result := myFree;
end;

class function tDEMO.UTIL_DiskSize(aDriveLetter: char): Int64;
{ --------------------------------------------------------------------------------------------------------------------------- }
var
  mySize, myFree: Int64;
  theDrive: char;
begin;
  if not UTIL_SYSUTIL_InternalGetDiskSpace(aDriveLetter, mySize, myFree) then result := -1
  else result := mySize;
end;

class function tDEMO.UTIL_DriveLetter(aLetter: string): char;
{ --------------------------------------------------------------------------------------------------------------------------- }
begin
  result := aLetter[low(aLetter)];
end;

class function tDEMO.UTIL_ExtractDriveLetter(aDriveCombo: tComboBox; aIndex: integer): char;
{ --------------------------------------------------------------------------------------------------------------------------- }
var
  theLetter: string;
begin
  theLetter := aDriveCombo.Items[aIndex];
  result    := UTIL_DriveLetter(theLetter);
end;

class function tDEMO.UTIL_SYSUTIL_InternalGetDiskSpace(Drive: char; var TotalSpace, FreeSpaceAvailable: Int64): Bool;
{ --------------------------------------------------------------------------------------------------------------------------- }
var
  RootPath: array [0 .. 4] of char;
  RootPtr: PChar;
begin
  RootPtr := nil;
  if Drive <> '' then
    begin
      RootPath[0] := Drive;
      RootPath[1] := ':';
      RootPath[2] := '\';
      RootPath[3] := #0;
      RootPtr     := RootPath;
    end;
  result := GetDiskFreeSpaceEx(RootPtr, FreeSpaceAvailable, TotalSpace, nil);
end;

end.
