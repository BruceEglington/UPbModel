program Project1;

uses
  Vcl.Forms,
  PbLoss3 in 'PbLoss3.pas' {fmPbLoss3},
  UPb_dm in 'UPb_dm.pas' {dmUPb: TDataModule},
  UPb_varb in 'UPb_varb.pas',
  UPb_ShtIm in 'UPb_ShtIm.pas' {fmSheetImport},
  About in 'About.pas' {AboutBox},
  Allsorts in '..\Eglington Delphi common code items\Allsorts.pas',
  GDW_varb in '..\GDW3x\GDW_varb.pas',
  Gd_file in '..\GDW3x\Gd_file.pas',
  gd_drv in '..\GDW3x\gd_drv.pas' {fmOptDir},
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  TStyleManager.TrySetStyle('Aqua Light Slate');
  Application.CreateForm(TdmUPb, dmUPb);
  Application.CreateForm(TfmPbLoss3, fmPbLoss3);
  Application.CreateForm(TfmSheetImport, fmSheetImport);
  Application.CreateForm(TAboutBox, AboutBox);
  Application.CreateForm(TfmOptDir, fmOptDir);
  Application.Run;
end.
