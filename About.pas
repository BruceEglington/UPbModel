unit About;

interface

uses Windows, Classes, Graphics, Forms, Controls, StdCtrls,
SysUtils,
  Buttons, ExtCtrls, System.ImageList, Vcl.ImgList, Vcl.VirtualImageList;

type
  TAboutBox = class(TForm)
    Panel1: TPanel;
    OKButton: TButton;
    ProgramIcon: TImage;
    ProductName: TLabel;
    Version: TLabel;
    Copyright: TLabel;
    Comments: TLabel;
    VirtualImageList1: TVirtualImageList;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  AboutBox: TAboutBox;

implementation

{$R *.DFM}

uses UPb_dm;

procedure TAboutBox.FormShow(Sender: TObject);

function  GetAppVersion:string;
var
  Size, Size2: DWord;
  Pt, Pt2: Pointer;
begin
 Size := GetFileVersionInfoSize(PChar (ParamStr (0)), Size2);
 if Size > 0 then
 begin
   GetMem (Pt, Size);
   try
      GetFileVersionInfo (PChar (ParamStr (0)), 0, Size, Pt);
      VerQueryValue (Pt, '\', Pt2, Size2);
      with TVSFixedFileInfo (Pt2^) do
      begin
        Result:= ' '+
                 IntToStr (HiWord (dwFileVersionMS)) + '.' +
                 IntToStr (LoWord (dwFileVersionMS)) + '.' +
                 IntToStr (HiWord (dwFileVersionLS)) + '.' +
                 IntToStr (LoWord (dwFileVersionLS));
     end;
   finally
     FreeMem (Pt);
   end;
 end;
end;

begin
  Version.Caption := 'Version '+ GetAppVersion;
end;

end.
 
