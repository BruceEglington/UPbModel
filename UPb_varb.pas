unit UPb_varb;

{$J+}

interface

uses
  SysUtils, AllSorts;
const
  iModel = 0;
  iStart = 1;
  iStage1 = 2;
  iPrimaryIsochron = 3;
  iStage1Results = 4;
  iStage2 = 5;
  iSecondaryIsochron = 6;
  iStage2Results = 7;
  iSamples = 8;
  iMu = 0;
  iKappa = 0;
  iAge = 0;
  ShowI = 0;

var
  GlobalChosenStyle : string;
  SampleNoStr, ShowDataStr,
  s238204Str, s232238Str,
  s206204Str, s207204Str,
  s208204Str, s207206Str,
  s208206Str                   : string;
  SampleNoCol, ShowDataCol,
  n238204Col, n232238Col,
  n206204Col, n207204Col,
  n208204Col, n207206Col,
  n208206Col                   : integer;
  DataPath                     : string;
  FullFileName                 : string;
  RowCount                     : array[1..10] of integer;
  ImportSheetNumber            : integer;
  FromRowValueString,
  ToRowValueString             : string;
  FromRow, ToRow               : integer;
  PbModelFileName              : string;

  function ConvertCol2Int(AnyString : string) : integer;

implementation


function ConvertCol2Int(AnyString : string) : integer;
var
  itmp    : integer;
  tmpStr  : string;
  tmpChar : char;
begin
    AnyString := UpperCase(AnyString);
    tmpStr := AnyString;
    ClearNull(tmpStr);
    Result := 0;
    if (length(tmpStr) = 2) then
    begin
      tmpChar := tmpStr[1];
      itmp := (ord(tmpChar)-64)*26;
      tmpChar := tmpStr[2];
      Result := itmp+(ord(tmpChar)-64);
    end else
    begin
      tmpChar := tmpStr[1];
      Result := (ord(tmpChar)-64);
    end;
end;


end.
