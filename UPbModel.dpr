program UPbModel;

uses
  madExcept,
  madLinkDisAsm,
  madListHardware,
  madListProcesses,
  madListModules,
  Vcl.Forms,
  PbLoss3 in 'PbLoss3.pas' {fmPbLoss3},
  UPb_dm in 'UPb_dm.pas' {dmUPb: TDataModule},
  UPb_varb in 'UPb_varb.pas',
  UPb_ShtIm in 'UPb_ShtIm.pas' {fmSheetImport},
  About in 'About.pas' {AboutBox},
  Allsorts in '..\Eglington Delphi common code items\Allsorts.pas',
  Vcl.Themes,
  Vcl.Styles,
  gd_drv in '..\GeoDate-VCL\gd_drv.pas',
  GDW_varb in '..\GeoDate-VCL\GDW_varb.pas',
  dmGdtmpDB in '..\GeoDate-VCL\dmGdtmpDB.pas' {dmGdwtmp: TDataModule},
  Gd_file in '..\GeoDate-VCL\Gd_file.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TdmUPb, dmUPb);
  Application.CreateForm(TfmPbLoss3, fmPbLoss3);
  Application.CreateForm(TfmSheetImport, fmSheetImport);
  Application.CreateForm(TAboutBox, AboutBox);
  Application.CreateForm(TdmGdwtmp, dmGdwtmp);
  Application.Run;
end.
