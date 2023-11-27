unit PbLoss3;

{need to correct code to plot samples in 232Th/238U vs 208Pb/206Pb space}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, Buttons, ToolWin, ComCtrls, OleCtrls,
  DBCtrls, Grids, DBGrids, DB, System.UITypes,
  System.IOUtils, VCL.Themes,
  Provider, DBClient, IniFiles, Menus,
  midaslib, VclTee.TeeGDIPlus, VCLTee.Series, VCLTee.TeEngine,
  VCLTee.TeeProcs, VCLTee.Chart, Vcl.ExtDlgs, VCLTee.TeeComma, System.ImageList,
  Vcl.ImgList, Vcl.VirtualImageList, SVGIconVirtualImageList;

type
  TfmPbLoss3
   = class(TForm)
    sbPbLoss: TStatusBar;
    ToolBar1: TToolBar;
    ExitBtn: TSpeedButton;
    sbCalculate: TSpeedButton;
    OpenDialog1: TOpenDialog;
    SaveDialog1: TSaveDialog;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    sbFileOpen: TSpeedButton;
    SpeedButton1: TSpeedButton;
    sbFileSave: TSpeedButton;
    SpeedButton4: TSpeedButton;
    sbFileSaveGEODATE: TSpeedButton;
    SaveDialogGEODATE: TSaveDialog;
    pc1: TPageControl;
    tsDefinition: TTabSheet;
    gbStartUniform: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    eT1: TEdit;
    eR64_0: TEdit;
    eR74_0: TEdit;
    Panel1: TPanel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    eFromMu: TEdit;
    eToMu: TEdit;
    Panel2: TPanel;
    rbRandomMu: TRadioButton;
    rbSteppedMu: TRadioButton;
    gbChange: TGroupBox;
    Label7: TLabel;
    Label8: TLabel;
    eT2: TEdit;
    eTSigma: TEdit;
    Panel3: TPanel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    eDelMuLo: TEdit;
    eDelMuHi: TEdit;
    Panel4: TPanel;
    rbRandomChange: TRadioButton;
    rbSteppedChange: TRadioButton;
    pMuModel: TPanel;
    Label9: TLabel;
    rbSingleStage: TRadioButton;
    rbTwoStage: TRadioButton;
    gbRegression: TGroupBox;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    eWeight6: TEdit;
    eWeight7: TEdit;
    eR_Coeff: TEdit;
    TabUranogenic: TTabSheet;
    TabThorogenic: TTabSheet;
    TabPercentChange: TTabSheet;
    Label14: TLabel;
    eR84_0: TEdit;
    Panel5: TPanel;
    Label15: TLabel;
    Label16: TLabel;
    Label20: TLabel;
    eFromK: TEdit;
    eToK: TEdit;
    Panel6: TPanel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    eDelKLo: TEdit;
    eDelKHi: TEdit;
    Label24: TLabel;
    eWeight8: TEdit;
    Label25: TLabel;
    eR_Coeff8: TEdit;
    Bevel1: TBevel;
    TabMu1MuSource: TTabSheet;
    TabMu1Mu2: TTabSheet;
    TabK1K2: TTabSheet;
    rbInverseSteppedChange: TRadioButton;
    Label13: TLabel;
    eNPoints: TEdit;
    tsSamples: TTabSheet;
    DBGrid1: TDBGrid;
    tsPb46Pb76: TTabSheet;
    tsPb76Pb86: TTabSheet;
    gbStartMix: TGroupBox;
    Label26: TLabel;
    Label27: TLabel;
    Label28: TLabel;
    Label29: TLabel;
    Label30: TLabel;
    eR64_0s1: TEdit;
    eR74_0s1: TEdit;
    Panel7: TPanel;
    Label31: TLabel;
    Label32: TLabel;
    Label33: TLabel;
    eFromMuMix: TEdit;
    eToMuMix: TEdit;
    Panel8: TPanel;
    rbRandomMuMix: TRadioButton;
    rbSteppedMuMix: TRadioButton;
    eR84_0s1: TEdit;
    Panel9: TPanel;
    Label34: TLabel;
    Label35: TLabel;
    Label36: TLabel;
    eFromKMix: TEdit;
    eToKMix: TEdit;
    eNPointsMix: TEdit;
    rbUniformStart: TRadioButton;
    rbMixedStart: TRadioButton;
    Label37: TLabel;
    eR64_0s2: TEdit;
    eR74_0s2: TEdit;
    eR84_0s2: TEdit;
    ProgressBar1: TProgressBar;
    tsThUPbrad: TTabSheet;
    Label38: TLabel;
    eT1Mix: TEdit;
    gbScale: TGroupBox;
    Label39: TLabel;
    Label40: TLabel;
    Label41: TLabel;
    e64Min: TEdit;
    e64Max: TEdit;
    Label42: TLabel;
    e74Min: TEdit;
    e74Max: TEdit;
    Label43: TLabel;
    e84Min: TEdit;
    e84Max: TEdit;
    cbUseMinMax: TCheckBox;
    eTitle: TEdit;
    Label44: TLabel;
    sbCalculateModelValues: TSpeedButton;
    Panel10: TPanel;
    DBNavigator1: TDBNavigator;
    cbShowSamples: TCheckBox;
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    E1: TMenuItem;
    N1: TMenuItem;
    Open1: TMenuItem;
    Save1: TMenuItem;
    SaveasGeodatefile1: TMenuItem;
    Calculate1: TMenuItem;
    Help1: TMenuItem;
    About1: TMenuItem;
    N2: TMenuItem;
    Importsampledata1: TMenuItem;
    TabUPb: TTabSheet;
    ts6486: TTabSheet;
    ChUPb: TChart;
    Series1: TLineSeries;
    Series2: TPointSeries;
    Series3: TLineSeries;
    Series4: TLineSeries;
    Series5: TPointSeries;
    Series6: TLineSeries;
    Series7: TLineSeries;
    Series8: TPointSeries;
    Series9: TPointSeries;
    ChPbLossU: TChart;
    LineSeries1: TLineSeries;
    PointSeries1: TPointSeries;
    LineSeries2: TLineSeries;
    LineSeries3: TLineSeries;
    PointSeries2: TPointSeries;
    LineSeries4: TLineSeries;
    LineSeries5: TLineSeries;
    PointSeries3: TPointSeries;
    PointSeries4: TPointSeries;
    ChPbLossTh: TChart;
    LineSeries6: TLineSeries;
    PointSeries5: TPointSeries;
    LineSeries7: TLineSeries;
    LineSeries8: TLineSeries;
    PointSeries6: TPointSeries;
    LineSeries9: TLineSeries;
    LineSeries10: TLineSeries;
    PointSeries7: TPointSeries;
    PointSeries8: TPointSeries;
    ChMu1Mu2: TChart;
    PointSeries9: TPointSeries;
    ChK1K2: TChart;
    PointSeries13: TPointSeries;
    ChMu1MuSource: TChart;
    PointSeries17: TPointSeries;
    ChPb46Pb76: TChart;
    LineSeries26: TLineSeries;
    PointSeries21: TPointSeries;
    LineSeries27: TLineSeries;
    LineSeries28: TLineSeries;
    PointSeries22: TPointSeries;
    LineSeries29: TLineSeries;
    LineSeries30: TLineSeries;
    PointSeries23: TPointSeries;
    PointSeries24: TPointSeries;
    ChPb76Pb86: TChart;
    LineSeries31: TLineSeries;
    PointSeries25: TPointSeries;
    LineSeries32: TLineSeries;
    LineSeries33: TLineSeries;
    PointSeries26: TPointSeries;
    LineSeries34: TLineSeries;
    LineSeries35: TLineSeries;
    PointSeries27: TPointSeries;
    PointSeries28: TPointSeries;
    ChThUPb86: TChart;
    LineSeries36: TLineSeries;
    PointSeries29: TPointSeries;
    LineSeries37: TLineSeries;
    LineSeries38: TLineSeries;
    PointSeries30: TPointSeries;
    LineSeries39: TLineSeries;
    LineSeries40: TLineSeries;
    PointSeries31: TPointSeries;
    PointSeries32: TPointSeries;
    ChPb64Pb86: TChart;
    LineSeries41: TLineSeries;
    PointSeries33: TPointSeries;
    LineSeries42: TLineSeries;
    LineSeries43: TLineSeries;
    PointSeries34: TPointSeries;
    LineSeries44: TLineSeries;
    LineSeries45: TLineSeries;
    PointSeries35: TPointSeries;
    PointSeries36: TPointSeries;
    Series10: TLineSeries;
    Series11: TLineSeries;
    Series12: TLineSeries;
    Printgraph1: TMenuItem;
    Exportgraph1: TMenuItem;
    N3: TMenuItem;
    SavePictureDialog1: TSavePictureDialog;
    Series13: TPointSeries;
    tcChUPb: TTeeCommander;
    tcChPbLossU: TTeeCommander;
    tcPbLossTh: TTeeCommander;
    tcChPb46Pb76: TTeeCommander;
    tcChPb76Pb86: TTeeCommander;
    tcChThUPb86: TTeeCommander;
    tcChPb64Pb86: TTeeCommander;
    tcChMu1Mu2: TTeeCommander;
    tcChK1K2: TTeeCommander;
    tcChMu1MuSource: TTeeCommander;
    tsAge64: TTabSheet;
    tsAgeThU: TTabSheet;
    ChAgePb64: TChart;
    LineSeries11: TLineSeries;
    PointSeries10: TPointSeries;
    LineSeries12: TLineSeries;
    LineSeries13: TLineSeries;
    PointSeries11: TPointSeries;
    LineSeries14: TLineSeries;
    LineSeries15: TLineSeries;
    PointSeries12: TPointSeries;
    PointSeries14: TPointSeries;
    ChAgeThU: TChart;
    LineSeries16: TLineSeries;
    PointSeries15: TPointSeries;
    LineSeries17: TLineSeries;
    LineSeries18: TLineSeries;
    PointSeries16: TPointSeries;
    LineSeries19: TLineSeries;
    LineSeries20: TLineSeries;
    PointSeries18: TPointSeries;
    PointSeries19: TPointSeries;
    tcChAgePb64: TTeeCommander;
    tcChAgeThU: TTeeCommander;
    tsAge84: TTabSheet;
    tcChAgePb84: TTeeCommander;
    ChAgePb84: TChart;
    LineSeries21: TLineSeries;
    PointSeries20: TPointSeries;
    LineSeries22: TLineSeries;
    LineSeries23: TLineSeries;
    PointSeries37: TPointSeries;
    LineSeries24: TLineSeries;
    LineSeries25: TLineSeries;
    PointSeries38: TPointSeries;
    PointSeries39: TPointSeries;
    Options1: TMenuItem;
    VirtualImageList1: TVirtualImageList;
    Styles1: TMenuItem;
    SVGIconVirtualImageList1: TSVGIconVirtualImageList;
    procedure ExitBtnClick(Sender: TObject);
    procedure sbCalculateClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure sbFileOpenClick(Sender: TObject);
    procedure sbFileSaveClick(Sender: TObject);
    procedure sbFileSaveGEODATEClick(Sender: TObject);
    procedure eT1Change(Sender: TObject);
    procedure rbUniformStartClick(Sender: TObject);
    procedure rbMixedStartClick(Sender: TObject);
    procedure rbTwoStageClick(Sender: TObject);
    procedure rbSingleStageClick(Sender: TObject);
    procedure sbCalculateModelValuesClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var MyAction: TCloseAction);
    procedure About1Click(Sender: TObject);
    procedure Importsampledata1Click(Sender: TObject);
    procedure ChUPbDblClick(Sender: TObject);
    procedure Printgraph1Click(Sender: TObject);
    procedure Exportgraph1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure StyleClick(Sender: TObject);
  private
    { Private declarations }
    procedure GetModelParameters;
    procedure UpdateModelParameters;
    procedure GetIniFile;
    procedure SetIniFile;
  public
    { Public declarations }
    function CalcMu (Age, StartAge,
                     Model64, Model74,
                     R64, R74 : double) : double;
  end;

var
  fmPbLoss3: TfmPbLoss3;

implementation

{$R *.DFM}
{$J+}

uses
  gdw_varb, gd_file, About, UPb_varb, UPb_dm, UPb_ShtIm;

const
  MaxPoints = 250; //should be 250
  NPoints : integer = 20;
  Precision : double = 0.01;
  Weight6 : double = 0.1;
  Weight7 : double = 0.1;
  Weight8 : double = 0.15;
  Zero : double = 0.0;
  R_Coeff : double = 0.9;
  R_Coeff8 : double = 0.8;
  Decay238 : double = 1.55125e-10;
  Decay235 : double = 9.8485e-10;
  Decay232 : double = 4.9745e-11;
  Factor : double = 137.88;
  FromMu : double = 9.5;
  ToMu : double = 9.8;
  DelMuLo : double = -10.0;
  DelMuHi : double = 10.0;
  FromK : double = 3.6;
  ToK : double = 4.5;
  DelKLo : double = -10.0;
  DelKHi : double = 10.0;
  TCD : double = 4568.0;
  T0 : double = 3700.0;
  T1 : double = 2000.0;
  TSigma : double = 0.0;
  R64_CD : double = 9.307;
  R74_CD : double = 10.294;
  R84_CD : double = 29.487;
  R64_0 : double = 11.152;
  R74_0 : double = 12.998;
  R84_0 : double = 31.230;
  MuRand : char = 'R';
  ChangeRand : char = 'R';
  R64_0s1 : double = 11.152;
  R74_0s1 : double = 12.998;
  R84_0s1 : double = 31.230;
  R64_0s2 : double = 13.152;
  R74_0s2 : double = 14.998;
  R84_0s2 : double = 33.230;

var
  OutFile : text;
  Tx,MuRange,DelRange,Del,RangeK,DelRangeK,DelK : double;
  RMu, RMuMix,
  R64,R74,R84,
  SmpMu,NewMu,SmpK,NewK,MuSource,
  R64Mix,R74Mix,R84Mix : array[0..MaxPoints] of double;
  ParamFileName : string;
  ParamFilePath : string;
  ModelStart,
  Model64_CD, Model74_CD, Model84_CD,
  Model64_0, Model74_0, Model84_0,
  ModelMu, ModelK : double;
  StepSize1, StepSize2 : double;

  ImportForm : TfmSheetImport;

procedure TfmPbLoss3.StyleClick(Sender: TObject);
var
  StyleName : String;
  i : integer;
begin
  //get style name
  StyleName := TMenuItem(Sender).Caption;
  StyleName := StringReplace(StyleName, '&', '',
    [rfReplaceAll,rfIgnoreCase]);
  GlobalChosenStyle := StyleName;
  dmUPb.ChosenStyle := GlobalChosenStyle;
  //set active style
  Application.ProcessMessages;
  TStyleManager.SetStyle(GlobalChosenStyle);
  dmUPb.ChosenStyle := GlobalChosenStyle;
  Application.ProcessMessages;
  //check the currently selected menu item
  (Sender as TMenuItem).Checked := true;
  //uncheck all other style menu items
  for i := 0 to Styles1.Count-1 do
  begin
    if not Styles1.Items[i].Equals(Sender) then
      Styles1.Items[i].Checked := false;
  end;
  for i := 0 to Styles1.Count-1 do
  begin
    if Styles1.Items[i].Checked then GlobalChosenStyle := StringReplace(Styles1.Items[i].Caption, '&', '',
    [rfReplaceAll,rfIgnoreCase]);
  end;
  TStyleManager.SetStyle(GlobalChosenStyle);
  try
    dmUPb.ChosenStyle := GlobalChosenStyle;
  finally
    dmUPb.ChosenStyle := GlobalChosenStyle;
  end;
end;

procedure TfmPbLoss3.GetModelParameters;
var
  iCode : integer;
  //tmpStr : string;
begin
  if (rbUniformStart.Checked) then
  begin
    //tmpStr := eT1.Text;
    Val(eT1.Text,T0,iCode);
    Val(eR64_0.Text,R64_0,iCode);
    Val(eR74_0.Text,R74_0,iCode);
    Val(eR84_0.Text,R84_0,iCode);
    Val(eFromMu.Text,FromMu,iCode);
    Val(eToMu.Text,ToMu,iCode);
    Val(eNPoints.Text,NPoints,iCode);
    Val(eFromK.Text,FromK,iCode);
    Val(eToK.Text,ToK,iCode);
    if rbRandomMu.Checked then MuRand := 'R'
                          else MuRand := 'S';
  end;
  if (rbMixedStart.Checked) then
  begin
    Val(eT1Mix.Text,T0,iCode);
    Val(eR64_0s1.Text,R64_0s1,iCode);
    Val(eR74_0s1.Text,R74_0s1,iCode);
    Val(eR84_0s1.Text,R84_0s1,iCode);
    Val(eR64_0s2.Text,R64_0s2,iCode);
    Val(eR74_0s2.Text,R74_0s2,iCode);
    Val(eR84_0s2.Text,R84_0s2,iCode);
    Val(eFromMuMix.Text,FromMu,iCode);
    Val(eToMuMix.Text,ToMu,iCode);
    Val(eNPointsMix.Text,NPoints,iCode);
    Val(eFromKMix.Text,FromK,iCode);
    Val(eToKMix.Text,ToK,iCode);
    if rbRandomMuMix.Checked then MuRand := 'R'
                             else MuRand := 'S';
  end;
  Val(eWeight6.Text,Weight6,iCode);
  Val(eWeight7.Text,Weight7,iCode);
  Val(eWeight8.Text,Weight8,iCode);
  Val(eR_Coeff.Text,R_Coeff,iCode);
  Val(eR_Coeff8.Text,R_Coeff8,iCode);
  Val(eT2.Text,T1,iCode);
  Val(eTSigma.Text,TSigma,iCode);
  Val(eDelMuLo.Text,DelMuLo,iCode);
  if (DelMuLo < -100.0) then
  begin
    DelMuLo := -100.0;
    eDelMuLo.Text := '-100.00';
    Messagedlg('Minimum change may not be less than -100.00 %',mtWarning,[mbOK],0);
  end;
  Val(eDelMuHi.Text,DelMuHi,iCode);
  Val(eDelKLo.Text,DelKLo,iCode);
  if (DelKLo < -100.0) then
  begin
    DelKLo := -100.0;
    eDelKLo.Text := '-100.00';
    Messagedlg('Minimum change may not be less than -100.00 %',mtWarning,[mbOK],0);
  end;
  Val(eDelKHi.Text,DelKHi,iCode);
  if rbRandomChange.Checked then ChangeRand := 'R';
  if rbSteppedChange.Checked then ChangeRand := 'S';
  if rbInverseSteppedChange.Checked then ChangeRand := 'I';
end;

procedure TfmPbLoss3.UpdateModelParameters;
begin
  if (rbUniformStart.Checked) then
  begin
    eT1.Text :=      FormatFloat(' ###0.0',T0);
    eR64_0.Text :=   FormatFloat('  ##0.000',R64_0);
    eR74_0.Text :=   FormatFloat('  ##0.000',R74_0);
    eR84_0.Text :=   FormatFloat('  ##0.000',R84_0);
    eFromMu.Text :=  FormatFloat('  ##0.00',FromMu);
    eToMu.Text :=    FormatFloat('  ##0.00',ToMu);
    eNPoints.Text := FormatFloat('   #0',NPoints);
    eFromK.Text :=  FormatFloat('  ##0.00',FromK);
    eToK.Text :=    FormatFloat('  ##0.00',ToK);
    if (MuRand = 'R') then rbRandomMu.Checked := true
                      else rbSteppedMu.Checked := true;
  end;
  if (rbMixedStart.Checked) then
  begin
    eT1Mix.Text :=      FormatFloat(' ###0.0',T0);
    eR64_0s1.Text :=   FormatFloat('  ##0.000',R64_0s1);
    eR74_0s1.Text :=   FormatFloat('  ##0.000',R74_0s1);
    eR84_0s1.Text :=   FormatFloat('  ##0.000',R84_0s1);
    eR64_0s2.Text :=   FormatFloat('  ##0.000',R64_0s2);
    eR74_0s2.Text :=   FormatFloat('  ##0.000',R74_0s2);
    eR84_0s2.Text :=   FormatFloat('  ##0.000',R84_0s2);
    eFromMuMix.Text :=  FormatFloat('  ##0.00',FromMu);
    eToMuMix.Text :=    FormatFloat('  ##0.00',ToMu);
    eNPointsMix.Text := FormatFloat('   #0',NPoints);
    eFromKMix.Text :=  FormatFloat('  ##0.00',FromK);
    eToKMix.Text :=    FormatFloat('  ##0.00',ToK);
    if (MuRand = 'R') then rbRandomMuMix.Checked := true
                      else rbSteppedMuMix.Checked := true;
  end;
  eWeight6.Text := FormatFloat('  ##0.000',Weight6);
  eWeight7.Text := FormatFloat('  ##0.000',Weight7);
  eWeight8.Text := FormatFloat('  ##0.000',Weight8);
  eR_Coeff.Text := FormatFloat('    0.000',R_Coeff);
  eR_Coeff8.Text := FormatFloat('    0.000',R_Coeff8);
  eT2.Text :=      FormatFloat(' ###0.0',T1);
  eTSigma.Text :=  FormatFloat(' ###0.0',TSigma);
  eDelMuLo.Text := FormatFloat('####0.00',DelMuLo);
  eDelMuHi.Text := FormatFloat('####0.00',DelMuHi);
  eDelKLo.Text := FormatFloat('####0.00',DelKLo);
  eDelKHi.Text := FormatFloat('####0.00',DelKHi);
  if (ChangeRand = 'R') then rbRandomChange.Checked := true;
  if (ChangeRand = 'S') then rbSteppedChange.Checked := true;
  if (ChangeRand = 'I') then rbInverseSteppedChange.Checked := true;
  Model64_CD := R64_CD;
  Model74_CD := R74_CD;
  Model84_CD := R84_CD;
  if rbSingleStage.Checked then
  begin
    ModelStart := 4570.0;
    ModelMu := 7.19;
    Model64_0 := 9.307;
    Model74_0 := 10.294;
    ModelK := 4.619;
    Model84_0 := 29.487;
  end
  else begin
    ModelStart := 3700.0;
    ModelMu := 9.74;
    Model64_0 := 11.152;
    Model74_0 := 12.998;
    ModelK := 3.782;
    Model84_0 := 31.230;
  end;
end;

procedure TfmPbLoss3.ExitBtnClick(Sender: TObject);
begin
  SetIniFile;
  Close;
end;

procedure TfmPbLoss3.Exportgraph1Click(Sender: TObject);
begin
  SavePictureDialog1.FileName := 'UPbModel.jpg';
  if SavePictureDialog1.Execute then
  begin
    if (pc1.ActivePage = TabUPb) then ChUPb.SaveToBitmapFile(SavePictureDialog1.FileName);
    if (pc1.ActivePage = TabUPb) then ChUPb.SaveToBitmapFile(SavePictureDialog1.FileName);
    if (pc1.ActivePage = TabUranogenic) then ChPbLossU.SaveToBitmapFile(SavePictureDialog1.FileName);
    if (pc1.ActivePage = TabThorogenic) then ChPbLossTh.SaveToBitmapFile(SavePictureDialog1.FileName);
    if (pc1.ActivePage = TabMu1Mu2) then ChMu1Mu2.SaveToBitmapFile(SavePictureDialog1.FileName);
    if (pc1.ActivePage = TabK1K2) then ChK1K2.SaveToBitmapFile(SavePictureDialog1.FileName);
    if (pc1.ActivePage = TabMu1MuSource) then ChMu1MuSource.SaveToBitmapFile(SavePictureDialog1.FileName);
    if (pc1.ActivePage = tsPb46Pb76) then ChPb46Pb76.SaveToBitmapFile(SavePictureDialog1.FileName);
    if (pc1.ActivePage = tsPb76Pb86) then ChPb76Pb86.SaveToBitmapFile(SavePictureDialog1.FileName);
    if (pc1.ActivePage = tsThUPbrad) then ChThUPb86.SaveToBitmapFile(SavePictureDialog1.FileName);
    if (pc1.ActivePage = ts6486) then ChPb64Pb86.SaveToBitmapFile(SavePictureDialog1.FileName);
    if (pc1.ActivePage = tsAgeThU) then ChThUPb86.SaveToBitmapFile(SavePictureDialog1.FileName);
    if (pc1.ActivePage = tsAge64) then ChAgePb64.SaveToBitmapFile(SavePictureDialog1.FileName);
    if (pc1.ActivePage = tsAge84) then ChAgePb84.SaveToBitmapFile(SavePictureDialog1.FileName);
    if (pc1.ActivePage = tsThUPbrad) then ChThUPb86.SaveToBitmapFile(SavePictureDialog1.FileName);
  end;
end;

procedure TfmPbLoss3.sbCalculateClick(Sender: TObject);
const
  MaxCol = 8;
  {
  StepSize = 20.0;
  }
var
  StepSize : double;
  i, j : integer;
  iCode : integer;
  MaxX, MaxY, MaxZ, MaxMu, MaxK, MaxMuSource : single;
  R64Now, R74Now, R84Now, R46Now, R76Now, Slope, Intercept : double;
  R64Start2, R74Start2, R84Start2, R46Start2, R76Start2 : double;
  Tmu : double;
  Min64, Max64, Min74, Max74, Min84, Max84 : double;
  tmpSlope76, tmpIntercept76, tmpSlope86, tmpIntercept86 : double;
  KappaCalc : double;
begin
  ProgressBar1.Position := 0;
  ProgressBar1.Visible := true;
  pc1.ActivePage := tsDefinition;
  AnalType := '3';
  IAnalTyp := 3;
  MaxX := 0.0;
  MaxY := 0.0;
  MaxZ := 0.0;
  MaxMu := 0.0;
  MaxK := 0.0;
  MaxMuSource := 0.0;
  if cbUseMinMax.Checked then
  begin
    Val(e64Min.Text,Min64,iCode);
    Val(e74Min.Text,Min74,iCode);
    Val(e84Min.Text,Min84,iCode);
    Val(e64Max.Text,Max64,iCode);
    Val(e74Max.Text,Max74,iCode);
    Val(e84Max.Text,Max84,iCode);
  end;
  sbPbLoss.Panels[0].Text := ' ';
  sbPbLoss.Panels[1].Text := 'Initialising data arrays';
  sbPbLoss.Refresh;
  //U-Pb
  try
  for j := 0 to MaxCol do
  begin
    ChUPb.Series[j].Clear;
    ChPbLossU.Series[j].Clear;
    ChPbLossTh.Series[j].Clear;
    ChPb46Pb76.Series[j].Clear;
    ChPb76Pb86.Series[j].Clear;
    ChThUPb86.Series[j].Clear;
    ChPb64Pb86.Series[j].Clear;
    ChAgePb64.Series[j].Clear;
    ChAgePb84.Series[j].Clear;
    ChAgeThU.Series[j].Clear;
    ChUPb.Series[j].XValues.Order := loNone;
    ChUPb.Series[j].YValues.Order := loNone;
    ChPbLossU.Series[j].XValues.Order := loNone;
    ChPbLossU.Series[j].YValues.Order := loNone;
    ChPbLossTh.Series[j].XValues.Order := loNone;
    ChPbLossTh.Series[j].YValues.Order := loNone;
    ChPb46Pb76.Series[j].XValues.Order := loNone;
    ChPb46Pb76.Series[j].YValues.Order := loNone;
    ChPb76Pb86.Series[j].XValues.Order := loNone;
    ChPb76Pb86.Series[j].YValues.Order := loNone;
    ChThUPb86.Series[j].XValues.Order := loNone;
    ChThUPb86.Series[j].YValues.Order := loNone;
    ChPb64Pb86.Series[j].XValues.Order := loNone;
    ChPb64Pb86.Series[j].YValues.Order := loNone;
    ChAgePb64.Series[j].XValues.Order := loNone;
    ChAgePb64.Series[j].YValues.Order := loNone;
    ChAgePb84.Series[j].XValues.Order := loNone;
    ChAgePb84.Series[j].YValues.Order := loNone;
    ChAgeThU.Series[j].XValues.Order := loNone;
    ChAgeThU.Series[j].YValues.Order := loNone;
  end;
  ChMu1Mu2.Series[iMu].Clear;
  ChMu1Mu2.Series[iMu+1].Clear;
  ChK1K2.Series[iKappa].Clear;
  ChK1K2.Series[iKappa+1].Clear;
  ChMu1MuSource.Series[iMu].Clear;
  ChMu1MuSource.Series[iMu+1].Clear;
  ChMu1Mu2.Series[iMu].XValues.Order := loNone;
  ChMu1Mu2.Series[iMu].YValues.Order := loNone;
  ChK1K2.Series[iKappa].XValues.Order := loNone;
  ChK1K2.Series[iKappa].YValues.Order := loNone;
  ChMu1MuSource.Series[iMu].XValues.Order := loNone;
  ChMu1MuSource.Series[iMu].YValues.Order := loNone;
  except
     MessageDlg('Error initialising data arrays',mtWarning,[mbOK],0);
  end;
  ProgressBar1.Position := ProgressBar1.Position + 1;
  GetModelParameters;
  UpDateModelParameters;
  FillChar(RMu,SizeOf(RMu),0);
  FillChar(R64,SizeOf(R64),0);
  FillChar(R74,SizeOf(R74),0);
  FillChar(R84,SizeOf(R84),0);
  if rbSingleStage.Checked then
  begin
    ModelStart := 4570.0;
    ModelMu := 7.19;
    Model64_0 := 9.307;
    Model74_0 := 10.294;
    ModelK := 4.619;
    Model84_0 := 29.487;
  end
  else begin
    ModelStart := 3700.0;
    ModelMu := 9.74;
    Model64_0 := 11.152;
    Model74_0 := 12.998;
    ModelK := 3.782;
    Model84_0 := 31.230;
  end;
  MuRange:=ToMu-FromMu;
  DelRange:=DelMuHi-DelMuLo;
  RangeK:=ToK-FromK;
  DelRangeK:=DelKHi-DelKLo;

  ProgressBar1.Position := ProgressBar1.Position + 1;

  if (rbMixedStart.Checked) then
  begin
    tmpSlope76 := (R74_0s2-R74_0s1)/(R64_0s2-R64_0s1);
    tmpIntercept76 := R74_0s1-tmpSlope76*R64_0s1;
    tmpSlope86 := (R84_0s2-R84_0s1)/(R64_0s2-R64_0s1);
    tmpIntercept86 := R84_0s1-tmpSlope86*R64_0s1;
    for i:=1 to NPoints do
    begin
      //Calculate starting Pb isotope ratios for mixing line
      if rbRandomMuMix.Checked then
      begin
        RMuMix[i] := ModelMu + Random*(ToMu-FromMu);
        R64Mix[i] := R64_0s1 + Random*(R64_0s2-R64_0s1);
        R74Mix[i] := tmpIntercept76 + R64Mix[i]*tmpSlope76;
        R84Mix[i] := tmpIntercept86 + R64Mix[i]*tmpSlope86;
      end else
      begin
        RMuMix[i] := FromMu + ((1.0*i-1.0)*(ToMu-FromMu)/(1.0*NPoints));
        R64Mix[i] := R64_0s1 + ((1.0*i-1.0)*(R64_0s2-R64_0s1)/(1.0*NPoints));
        R74Mix[i] := R74_0s1 + ((1.0*i-1.0)*(R74_0s2-R74_0s1)/(1.0*NPoints));
        R84Mix[i] := R84_0s1 + ((1.0*i-1.0)*(R84_0s2-R84_0s1)/(1.0*NPoints));
      end;
    end;
  end;
  if (rbUniformStart.Checked) then
  begin
    for i:=1 to NPoints do
    begin
      RMuMix[i] := ModelMu;
      R64Mix[i] := R64_0;
      R74Mix[i] := R74_0;
      R84Mix[i] := R84_0;
    end;
  end;
  sbPbLoss.Panels[1].Text := 'Starting compositions';
  sbPbLoss.Refresh;
  //starting compositions
  for i := 1 to NPoints do
  begin
    //Starting sample mu's and k's
    if MuRand='R' then
    begin
      SmpMu[i]:=FromMu + Random*MuRange;
      SmpK[i]:=FromK + Random*RangeK;
    end else
    begin
      SmpMu[i]:= FromMu+((1.0*i-1.0)*MuRange/(1.0*NPoints));
      SmpK[i]:= FromK+((1.0*i-1.0)*RangeK/(1.0*NPoints));
    end;
    if (SmpMu[i] < 0.0) then SmpMu[i] := 0.0;
    if (SmpK[i] < 0.0) then SmpK[i] := 0.0;
    ChUPb.Series[iStart].AddXY(SmpMu[i],R64Mix[i]);
    ChPbLossU.Series[iStart].AddXY(R64Mix[i],R74Mix[i]);
    ChPbLossTh.Series[iStart].AddXY(R64Mix[i],R84Mix[i]);
    ChAgeThU.Series[iStart].AddXY(T0,SmpK[i]);
    if (R64Mix[i] > 0.0) then ChPb46Pb76.Series[iStart].AddXY(1.0/R64Mix[i],R74Mix[i]/R64Mix[i]);
    if (R64Mix[i] > 0.0) then ChPb64Pb86.Series[iStart].AddXY(R64Mix[i],R84Mix[i]/R64Mix[i]);
    if (R64Mix[i] > 0.0) then ChPb76Pb86.Series[iStart].AddXY(R74Mix[i]/R64Mix[i],R84Mix[i]/R64Mix[i]);
    if (R64Mix[i] > 0.0) then ChAgePb64.Series[iStart].AddXY(T0,R64Mix[i]);
    if (R84Mix[i] > 0.0) then ChAgePb84.Series[iStart].AddXY(T0,R84Mix[i]);
    //if ((R64Mix[i] > 0.0) and (R64Mix[i] <> Model64_0)) then ChThUPb86.Series[iStart].AddXY(SmpK[i],(R84Mix[i]-Model84_0)/(R64Mix[i]-Model64_0));
    if ((R64Mix[i] <> Model64_0) and (R84Mix[i] <> Model84_0)) then
    begin
      KappaCalc := (R84Mix[i]-Model84_CD)/(R64Mix[i]-Model64_CD)*(exp(Decay238*TCD*1.0e6)-1.0)/(exp(Decay232*TCD*1.0e6)-1.0);
      ChAgeThU.Series[iStart].AddXY(T0,KappaCalc);
      ChThUPb86.Series[iStart].AddXY(KappaCalc,(R84Mix[i]-Model84_CD)/(R64Mix[i]-Model64_CD));
    end;
  end;
  sbPbLoss.Panels[1].Text := 'Second stage average development';
  sbPbLoss.Refresh;
  //second stage average development
  Tx := 0.0;
  StepSize := (ModelStart)/200.0;
  i := 1;
  NewMu[1] := SmpMu[1] + SmpMu[1]*(DelMuHi + DelMuLo)/200.0;
  SmpK[1] := (ToK + FromK)/2.0;
  NewK[1] := SmpK[1] + SmpK[1]*(DelKHi + DelKLo)/200.0;
  if (NewMu[1] < 0.0) then NewMu[1] := 0.0;
  if (NewK[1] < 0.0) then NewK[1] := 0.0;
  SmpMu[1] := (ToMu + FromMu)/2.0;
  //ShowMessage(IntToStr(i)+' '+FormatFloat('###.###',SmpMu[1])+' '+FormatFloat('###.###',NewMu[1]));
  repeat
       RMu[1]:=NewMu[1]
                -1.0*(exp(Decay238*T1*1.0e+6)-exp(Decay238*Tx*1.0e+6));
       R64[1]:=R64Mix[1]
                +SmpMu[1]*(exp(Decay238*T0*1.0e+6)-exp(Decay238*T1*1.0e+6))
                +NewMu[1]*(exp(Decay238*T1*1.0e+6)-exp(Decay238*Tx*1.0e+6));
       R74[1]:=R74Mix[1]
           +SmpMu[1]/Factor*(exp(Decay235*T0*1.0e+6)-exp(Decay235*T1*1.0e+6))
           +NewMu[1]/Factor*(exp(Decay235*T1*1.0e+6)-exp(Decay235*Tx*1.0e+6));
       R84[1]:=R84Mix[1]
           +SmpK[1]*SmpMu[1]*(exp(Decay232*T0*1.0e+6)-exp(Decay232*T1*1.0e+6))
           +NewK[1]*NewMu[1]*(exp(Decay232*T1*1.0e+6)-exp(Decay232*Tx*1.0e+6));
    ChUPb.Series[iStage2].AddXY(RMu[1],R64[1]);
    ChPbLossU.Series[iStage2].AddXY(R64[1],R74[1]);
    ChPbLossTh.Series[iStage2].AddXY(R64[1],R84[1]);
    //ChAgeThU.Series[iStage2].AddXY(T1,SmpK[1]);
    if (R64[1] > 0.0) then ChPb46Pb76.Series[iStage2].AddXY(1.0/R64[1],R74[1]/R64[1]);
    //if (R64[1] > 0.0) then ChAgePb64.Series[iStage2].AddXY(T1,R64[1]);
    //ShowMessage(IntToStr(i)+' '+FormatFloat('###.###',RMu[1])+' '+FormatFloat('###.###',R64[1]));
    Tx := Tx + StepSize;
    i := i + 1;
  until (Tx >= T1);
  RMu[1]:=SmpMu[1]
                -1.0*(exp(Decay238*T0*1.0e+6)-exp(Decay238*T1*1.0e+6));
  R64[1]:=R64Mix[1]
                +SmpMu[1]*(exp(Decay238*T0*1.0e+6)-exp(Decay238*T1*1.0e+6));
  ChUPb.Series[iStage2].AddXY(RMu[1],R64[1]);
  //ShowMessage(IntToStr(i)+' '+FormatFloat('###.###',RMu[1])+' '+FormatFloat('###.###',R64[1]));
  R64Start2 := R64[1];
  R74Start2 := R74[1];
  R84Start2 := R84[1];
  R46Start2 := 1.0/R64[1];
  R76Start2 := R74[1]/R64[1];
  ProgressBar1.Position := ProgressBar1.Position + 1;
  Tx := 0.0;
  i := 1;
  repeat
       R64[1]:=R64Mix[1]
                +SmpMu[1]*(exp(Decay238*T0*1.0e+6)-exp(Decay238*T1*1.0e+6))
                +NewMu[1]*(exp(Decay238*T1*1.0e+6)-exp(Decay238*Tx*1.0e+6));
       R74[1]:=R74Mix[1]
           +SmpMu[1]/Factor*(exp(Decay235*T0*1.0e+6)-exp(Decay235*T1*1.0e+6))
           +NewMu[1]/Factor*(exp(Decay235*T1*1.0e+6)-exp(Decay235*Tx*1.0e+6));
       R84[1]:=R84Mix[1]
           +SmpK[1]*SmpMu[1]*(exp(Decay232*T0*1.0e+6)-exp(Decay232*T1*1.0e+6))
           +NewK[1]*NewMu[1]*(exp(Decay232*T1*1.0e+6)-exp(Decay232*Tx*1.0e+6));
    if (R64[1] <> R64Mix[1]) then ChPb76Pb86.Series[iStage2].AddXY((R74[1])/(R64[1]),(R84[1])/(R64[1]));
    //if ((R64[1] <> R64Mix[1]) and (R64[1] <> Model64_0)) then ChThUPb86.Series[iStage2].AddXY(NewK[1],(R84[1]-Model84_0)/(R64[1]-Model64_0));
    if (R64[1] <> R64Mix[1]) then ChPb64Pb86.Series[iStage2].AddXY(R64[1],(R84[1])/(R64[1]));
    //if (R64[1] <> R64Mix[1]) then ChAgePb64.Series[iStage2].AddXY(T1,R64[1]);
    //if (R64[1] <> R64Mix[1]) then ChAgeThU.Series[iStage2].AddXY(T1,NewK[1]);
    //if (R64[1] <> R64Mix[1]) then ChPb76Pb86.Series[iStage2].AddXY((R74[1]-R74Mix[1])/(R64[1]-R64Mix[1]),(R84[1]-R84Mix[1])/(R64[1]-R64Mix[1]));
    //if (R64[1] <> R64Mix[1]) then ChThUPb86.Series[iStage2].AddXY(NewK[1],(R84[1]-R84Mix[1])/(R64[1]-R64Mix[1]));
    //if (R64[1] <> R64Mix[1]) then ChPb64Pb86.Series[iStage2].AddXY(R64[1],(R84[1]-R84Mix[1])/(R64[1]-R64Mix[1]));
    Tx := Tx + StepSize;
    i := i + 1;
  until (Tx >= T1);
  ProgressBar1.Position := ProgressBar1.Position + 1;
  sbPbLoss.Panels[1].Text := 'First stage average development';
  sbPbLoss.Refresh;
  //first stage average development
  Tx := T1;
  i := 1;
  repeat
       RMu[1]:=SmpMu[1]
                -1.0*(exp(Decay238*T0*1.0e+6)-exp(Decay238*Tx*1.0e+6));
       R64[1]:=R64Mix[1]
                +SmpMu[1]*(exp(Decay238*T0*1.0e+6)-exp(Decay238*Tx*1.0e+6));
       R74[1]:=R74Mix[1]
           +SmpMu[1]/Factor*(exp(Decay235*T0*1.0e+6)-exp(Decay235*Tx*1.0e+6));
       R84[1]:=R84Mix[1]
                +SmpK[1]*SmpMu[1]*(exp(Decay232*T0*1.0e+6)-exp(Decay232*Tx*1.0e+6));
    ChUPb.Series[iStage1].AddXY(RMu[1],R64[1]);
    ChPbLossU.Series[iStage1].AddXY(R64[1],R74[1]);
    ChPbLossTh.Series[iStage1].AddXY(R64[1],R84[1]);
    if (R64[1] > 0.0) then ChPb46Pb76.Series[iStage1].AddXY(1.0/R64[1],R74[1]/R64[1]);
    if (R64[1] <> R64Mix[1]) then ChPb76Pb86.Series[iStage1].AddXY((R74[1])/(R64[1]),(R84[1])/(R64[1]));
    //if ((R64[1] <> R64Mix[1]) and (R64[1] <> Model64_0)) then ChThUPb86.Series[iStage1].AddXY(SmpK[1],(R84[1]-Model84_0)/(R64[1]-Model64_0));
    if (R64[1] <> R64Mix[1]) then ChPb64Pb86.Series[iStage1].AddXY(R64[1],(R84[1])/(R64[1]));
    //if (R64[1] <> R64Mix[1]) then ChPb76Pb86.Series[iStage1].AddXY((R74[1]-R74Mix[1])/(R64[1]-R64Mix[1]),(R84[1]-R84Mix[1])/(R64[1]-R64Mix[1]));
    //if (R64[1] <> R64Mix[1]) then ChThUPb86.Series[iStage1].AddXY(SmpK[1],(R84[1]-R84Mix[1])/(R64[1]-R64Mix[1]));
    //if (R64[1] <> R64Mix[1]) then ChPb64Pb86.Series[iStage1].AddXY(R64[1],(R84[1]-R84Mix[1])/(R64[1]-R64Mix[1]));
    Tx := Tx + StepSize;
    i := i + 1;
  until (Tx >= T0);
  ProgressBar1.Position := ProgressBar1.Position + 1;

  FillChar(RMu,SizeOf(RMu),0);
  FillChar(R64,SizeOf(R64),0);
  FillChar(R74,SizeOf(R74),0);
  FillChar(R84,SizeOf(R84),0);
  Title:='T1 '+FormatFloat('####.#',T0)+' T2 '+FormatFloat('####.#',T1)
          +' sigma '+FormatFloat('###.#',TSigma)+' mu '+FormatFloat('####.#',FromMu)
          +' to '+FormatFloat('####.#',ToMu)+MuRand
          +' %delta mu '+FormatFloat('####',DelMuLo)+' to '+FormatFloat('####',DelMuHi)
          +ChangeRand;
  {
  sbPbLoss.Panels[1].Text := Title;
  }
  RMu[0]:=ModelMu;
  R64[0]:=R64_0;
  R74[0]:=R74_0;
  R84[0]:=R84_0;
  AnalTyp[1]:='3';
  ErrTyp[1]:='1';
  RecordNo[1]:=1;
  N_Rep := 999;
  SmpNo[1]:='Zero';
  RFlg[1]:='N';
  PFlg[1]:='Y';
  Conc[1,1]:=0.0;
  Conc[1,2]:=0.0;
  Conc[1,3]:=0.0;
  Ratio[1,1]:=R64[0];
  Ratio[1,2]:=R74[0];
  Ratio[1,3]:=0.0;
  Ratio[1,4]:=0.0;
  Ratio[1,5]:=0.0;
  XPrec[1]:=Precision;
  YPrec[1]:=Precision;
  ZPrec[1]:=Precision;
  ErrorWt[1,1]:=Weight6;
  ErrorWt[1,2]:=Weight7;
  ErrorWt[1,3]:=0.0;
  R[1]:=R_Coeff;
  Randomize;
  sbPbLoss.Panels[1].Text := 'Stage 1 results';
  sbPbLoss.Refresh;

  //domain development - Stage 1 results
  //ShowMessage(IntToStr(i)+' '+FormatFloat('###.###',T1)+' '+FormatFloat('###.###',ModelStart));
  //ShowMessage(IntToStr(i)+' '+FormatFloat('###.###',Model64_0)+' '+FormatFloat('###.###',Model74_0));
  for i:=1 to NPoints do
  begin
    //Give Starting sample mu's and k's
    {
    if MuRand='R' then
    begin
      SmpMu[i]:=FromMu + Random*MuRange;
      SmpK[i]:=FromK + Random*RangeK;
    end else
    begin
      SmpMu[i]:= FromMu+((1.0*i-1.0)*MuRange/(1.0*NPoints));
      SmpK[i]:= FromK+((1.0*i-1.0)*RangeK/(1.0*NPoints));
    end;
    }
    Tx:=T1-(TSigma*0.5);
    Tx:=Tx+(Random*TSigma);
    //first stage results
    RMu[i]:=SmpMu[i]
             -1.0*(exp(Decay238*T0*1.0e+6)-exp(Decay238*Tx*1.0e+6));
    R64[i]:=R64Mix[i]
             +SmpMu[i]*(exp(Decay238*T0*1.0e+6)-exp(Decay238*Tx*1.0e+6));
    R74[i]:=R74Mix[i]
             +SmpMu[i]/Factor*(exp(Decay235*T0*1.0e+6)-exp(Decay235*Tx*1.0e+6));
    R84[i]:=R84Mix[i]
             +SmpK[i]*SmpMu[i]*(exp(Decay232*T0*1.0e+6)-exp(Decay232*Tx*1.0e+6));
    ChUPb.Series[iStage1Results].AddXY(RMu[i],R64[i]);
    ChPbLossU.Series[iStage1Results].AddXY(R64[i],R74[i]);
    ChPbLossTh.Series[iStage1Results].AddXY(R64[i],R84[i]);
    if (R64[i] > 0.0) then ChPb46Pb76.Series[iStage1Results].AddXY(1.0/R64[i],R74[i]/R64[i]);
    if (R64[i] <> R64Mix[i]) then ChPb76Pb86.Series[iStage1Results].AddXY((R74[i])/(R64[i]),(R84[i])/(R64[i]));
    //if ((R64[i] <> R64Mix[i]) and (R64[i] <> Model64_0)) then ChThUPb86.Series[iStage1Results].AddXY(SmpK[i],(R84[i]-Model84_0)/(R64[i]-Model64_0));
    if (R64[i] <> R64Mix[i]) then ChPb64Pb86.Series[iStage1Results].AddXY(R64[i],(R84[i])/(R64[i]));
    if (R64[i] <> R64Mix[i]) then ChAgePb64.Series[iStage1Results].AddXY(T1,R64[i]);
    if (R84[i] <> R84Mix[i]) then ChAgePb84.Series[iStage1Results].AddXY(T1,R84[i]);
    if ((R64[i] <> R64Mix[i]) and (R84[i] <> R84Mix[i]) and (R64[i] <> Model64_0)) then
    begin
      KappaCalc := (R84[i]-Model84_CD)/(R64[i]-Model64_CD)*(exp(Decay238*TCD*1.0e6)-1.0)/(exp(Decay232*TCD*1.0e6)-1.0);
      ChAgeThU.Series[iStage1Results].AddXY(T1,KappaCalc);
      ChThUPb86.Series[iStage1Results].AddXY(KappaCalc,(R84[i]-Model84_CD)/(R64[i]-Model64_CD));
    end;
    //if (R64[i] <> R64Mix[i]) then ChPb76Pb86.Series[iStage1Results].AddXY((R74[i]-R74Mix[i])/(R64[i]-R64Mix[i]),(R84[i]-R84Mix[i])/(R64[i]-R64Mix[i]));
    //if (R64[i] <> R64Mix[i]) then ChThUPb86.Series[iStage1Results].AddXY(SmpK[i],(R84[i]-R84Mix[i])/(R64[i]-R64Mix[i]));
    //if (R64[i] <> R64Mix[i]) then ChPb64Pb86.Series[iStage1Results].AddXY(R64[i],(R84[i]-R84Mix[i])/(R64[i]-R64Mix[i]));
   ProgressBar1.Position := ProgressBar1.Position + 1;
    case ChangeRand of
      'I' : begin
        NewMu[i]:=SmpMu[i]*(100.0+DelMuHi-(i-1.0)*DelRange/NPoints)/100.0;
        NewK[i]:=SmpK[i]*(100.0+DelKLo+(i-1.0)*DelRangeK/NPoints)/100.0;
      end;
      'R' :begin
        NewMu[i]:=SmpMu[i]*(100.0+(DelMuLo+Random*DelRange))/100.0;
        NewK[i]:=SmpK[i]*(100.0+(DelKLo+Random*DelRangeK))/100.0;
      end;
      'S' : begin
        NewMu[i]:=SmpMu[i]*(100.0+DelMuLo+(i-1.0)*DelRange/NPoints)/100.0;
        NewK[i]:=SmpK[i]*(100.0+DelKLo+(i-1.0)*DelRangeK/NPoints)/100.0;
      end;
    end;
    if (NewMu[i] < 0.0) then NewMu[i] := 0.0;
    if (NewK[i] < 0.0) then NewK[i] := 0.0;
    Del:=100.0*(NewMu[i]/SmpMu[i] - 1.0);
    DelK:=100.0*(NewK[i]/SmpK[i] - 1.0);
    sbPbLoss.Panels[1].Text := 'Stage 2 resuls';
    sbPbLoss.Refresh;
    //second stage results
       RMu[i]:=NewMu[i]
                -1.0*(exp(Decay238*Tx*1.0e+6)-1.0);
       if (RMu[i] < 0.0) then RMu[i] := 0.0;
       R64[i]:=R64Mix[i]
                +SmpMu[i]*(exp(Decay238*T0*1.0e+6)-exp(Decay238*Tx*1.0e+6))
                +NewMu[i]*(exp(Decay238*Tx*1.0e+6)-1.0);
       R74[i]:=R74Mix[i]
           +SmpMu[i]/Factor*(exp(Decay235*T0*1.0e+6)-exp(Decay235*Tx*1.0e+6))
           +NewMu[i]/Factor*(exp(Decay235*Tx*1.0e+6)-1.0);
       R84[i]:=R84Mix[i]
                +SmpK[i]*SmpMu[i]*(exp(Decay232*T0*1.0e+6)-exp(Decay232*Tx*1.0e+6))
                +NewK[i]*NewMu[i]*(exp(Decay232*Tx*1.0e+6)-1.0);
    MuSource[i] := CalcMu(T1,ModelStart,Model64_0,Model74_0,R64[i],R74[i]);
  //ShowMessage(IntToStr(i)+' '+FormatFloat('###.###',R74[i])+' '+FormatFloat('###.###',MuSource[i]));
  //ShowMessage(IntToStr(i)+' '+FormatFloat('###.###',R64Mix[1])+' '+FormatFloat('###.###',R74Mix[1]));
  //ShowMessage(IntToStr(i)+' '+FormatFloat('###.#####',Slope)+' '+FormatFloat('###.#####',Intercept));
  //ShowMessage(IntToStr(i)+' '+FormatFloat('###.###',MaxX)+' '+FormatFloat('###.#####',StepSize1));
    AnalTyp[i+1]:='3';
    ErrTyp[i+1]:='1';
    RecordNo[i+1]:=i+1;
    SmpNo[i+1]:=FormatFloat('########.#',Del);
    RFlg[i+1]:='Y';
    PFlg[i+1]:='Y';
    Conc[i+1,1]:=SmpMu[i];
    Conc[i+1,2]:=NewMu[i];
    Conc[i+1,3]:=0.0;
    Ratio[i+1,1]:=R64[i];
    Ratio[i+1,2]:=R74[i];
    Ratio[i+1,3]:=0.0;
    Ratio[i+1,4]:=0.0;
    Ratio[i+1,5]:=0.0;
    XPrec[i+1]:=Precision;
    YPrec[i+1]:=Precision;
    ZPrec[i+1]:=Precision;
    ErrorWt[i+1,1]:=Weight6;
    ErrorWt[i+1,2]:=Weight7;
    ErrorWt[i+1,3]:=0.0;
    R[i+1]:=R_Coeff;
    if (MaxX < R64[i]) then MaxX := R64[i];
    if (MaxY < R74[i]) then MaxY := R74[i];
    if (MaxZ < R84[i]) then MaxZ := R84[i];
    if (MaxMu < SmpMu[i]) then MaxMu := SmpMu[i];
    if (MaxMu < NewMu[i]) then MaxMu := NewMu[i];
    if (MaxK < SmpK[i]) then MaxK := SmpK[i];
    if (MaxK < NewK[i]) then MaxK := NewK[i];
    if (MaxMuSource < SmpMu[i]) then MaxMuSource := SmpMu[i];
    if (MaxMuSource < MuSource[i]) then MaxMuSource := MuSource[i];
    ChUPb.Series[iStage2Results].AddXY(RMu[i],R64[i]);
    ChPbLossU.Series[iStage2Results].AddXY(R64[i],R74[i]);
    ChPbLossTh.Series[iStage2Results].AddXY(R64[i],R84[i]);
    if (R64[i] > 0.0) then ChPb46Pb76.Series[iStage2Results].AddXY(1.0/R64[i],R74[i]/R64[i]);
    if (R64[i] <> R64Mix[i]) then ChPb76Pb86.Series[iStage2Results].AddXY((R74[i])/(R64[i]),(R84[i])/(R64[i]));
    //if (R64[i] <> R64Mix[i]) then ChThUPb86.Series[iStage2Results].AddXY(SmpK[i],(R84[i]-Model84_0)/(R64[i]-Model64_0));
    if (R64[i] <> R64Mix[i]) then ChPb64Pb86.Series[iStage2Results].AddXY(R64[i],(R84[i])/(R64[i]));
    if (R64[i] <> R64Mix[i]) then ChAgePb64.Series[iStage2Results].AddXY(0.0,R64[i]);
    if (R84[i] <> R84Mix[i]) then ChAgePb84.Series[iStage2Results].AddXY(0.0,R84[i]);
    if ((R64[i] <> R64Mix[i]) and (R84[i] <> R84Mix[i]) and (R64[i] <> Model64_0)) then
    begin
      KappaCalc := (R84[i]-Model84_CD)/(R64[i]-Model64_CD)*(exp(Decay238*TCD*1.0e6)-1.0)/(exp(Decay232*TCD*1.0e6)-1.0);
      ChAgeThU.Series[iStage2Results].AddXY(0.0,KappaCalc);
      ChThUPb86.Series[iStage2Results].AddXY(KappaCalc,(R84[i]-Model84_CD)/(R64[i]-Model64_CD));
    end;
    //if (R64[i] <> R64Mix[i]) then ChPb76Pb86.Series[iStage2Results].AddXY((R74[i]-R74Mix[i])/(R64[i]-R64Mix[i]),(R84[i]-R84Mix[i])/(R64[i]-R64Mix[i]));
    //if (R64[i] <> R64Mix[i]) then ChThUPb86.Series[iStage2Results].AddXY(SmpK[i],(R84[i]-R84Mix[i])/(R64[i]-R64Mix[i]));
    //if (R64[i] <> R64Mix[i]) then ChPb64Pb86.Series[iStage2Results].AddXY(R64[i],(R84[i]-R84Mix[i])/(R64[i]-R64Mix[i]));
    ChMu1Mu2.Series[iMu].AddXY(SmpMu[i],NewMu[i]);
    ChMu1Mu2.Series[iMu+1].AddXY(SmpMu[i],SmpMu[i]);
    ChMu1Mu2.Series[iMu+1].AddXY(NewMu[i],NewMu[i]);
    ChK1K2.Series[iKappa].AddXY(SmpK[i],NewK[i]);
    ChK1K2.Series[iKappa+1].AddXY(SmpK[i],SmpK[i]);
    ChK1K2.Series[iKappa+1].AddXY(NewK[i],NewK[i]);
    ChMu1MuSource.Series[iMu].AddXY(SmpMu[i],MuSource[i]);
    ChMu1MuSource.Series[iMu+1].AddXY(SmpMu[i],SmpMu[i]);
    ChMu1MuSource.Series[iMu+1].AddXY(MuSource[i],MuSource[i]);
  end;
  R64Now := R64[1];
  R74Now := R74[1];
  R84Now := R84[1];
  R46Now := 1.0/R64[1];
  R76Now := R74[1]/R64[1];
  NumberOfPoints:=NPoints+1;
  ProgressBar1.Position := ProgressBar1.Position + 1;

  sbPbLoss.Panels[1].Text := 'Model curve';
  sbPbLoss.Panels[0].Text := ' ';
  sbPbLoss.Refresh;
  //StepSize := (ModelStart)/100.0;
  //ShowMessage('StepSize is '+FormatFloat('###0.000',StepSize));

  //model curve
  Tx := 0.0;
  i := 1;
  repeat
    if (ShowI > 0) then
    begin
      sbPbLoss.Panels[0].Text := IntToStr(i);
      sbPbLoss.Refresh;
    end;
    RMu[i]:=ModelMu
             -1.0*(exp(Decay238*ModelStart*1.0e+6)-exp(Decay238*Tx*1.0e+6));
    R64[i]:=Model64_0
             +ModelMu*(exp(Decay238*ModelStart*1.0e+6)-exp(Decay238*Tx*1.0e+6));
    R74[i]:=Model74_0
             +ModelMu/Factor*(exp(Decay235*ModelStart*1.0e+6)-exp(Decay235*Tx*1.0e+6));
    R84[i]:=Model84_0
             +ModelK*ModelMu*(exp(Decay232*ModelStart*1.0e+6)-exp(Decay232*Tx*1.0e+6));
    ChUPb.Series[iModel].AddXY(RMu[i],R64[i]);
    ChPbLossU.Series[iModel].AddXY(R64[i],R74[i]);
    ChPbLossTh.Series[iModel].AddXY(R64[i],R84[i]);
    if (R64[i] > 0.0) then ChPb46Pb76.Series[iModel].AddXY(1.0/R64[i],R74[i]/R64[i]);
    if (R64[i] <> Model64_0) then ChPb76Pb86.Series[iModel].AddXY((R74[i])/(R64[i]),(R84[i])/(R64[i]));
    //if (R64[i] <> Model64_0) then ChPb76Pb86.Series[iModel].AddXY((R74[i]-Model74_0)/(R64[i]-Model64_0),(R84[i]-Model84_0)/(R64[i]-Model64_0));
    //if (R64[i] <> Model64_0) then ChThUPb86.Series[iModel].AddXY(SmpK[i],(R84[i]-Model84_0)/(R64[i]-Model64_0));
    if (R64[i] <> Model64_0) then ChPb64Pb86.Series[iModel].AddXY(R64[i],(R84[i])/(R64[i]));
    //if (R64[i] <> Model64_0) then ChPb64Pb86.Series[iModel].AddXY(R64[i],(R84[i]-Model84_0)/(R64[i]-Model64_0));
    //ChAgeThU.Series[iModel].AddXY(Tx,2.62+0.000636*Tx-0.0000000645*Tx*Tx);
    if ((R64[i] <> Model64_CD) and (R84[i] <> Model84_CD)) then
    begin
      KappaCalc := (R84[i]-Model84_CD)/(R64[i]-Model64_CD)*(exp(Decay238*TCD*1.0e6)-1.0)/(exp(Decay232*TCD*1.0e6)-1.0);
      ChAgeThU.Series[iModel].AddXY(Tx,KappaCalc);
      ChThUPb86.Series[iModel].AddXY(KappaCalc,(R84[i]-Model84_CD)/(R64[i]-Model64_CD));
    end;
    if (MaxX < R64[i]) then MaxX := R64[i];
    if (MaxY < R74[i]) then MaxY := R74[i];
    if (MaxZ < R84[i]) then MaxZ := R84[i];
    Tx := Tx + StepSize;
    i := i + 1;
  until (Tx >= T0);
  ProgressBar1.Position := ProgressBar1.Position + 1;

  sbPbLoss.Panels[1].Text := 'Primary isochron uranogenic';
  sbPbLoss.Refresh;
  //Primary isochron uranogenic
  StepSize1 := (MaxX-R64Mix[1])/100.0;
  Slope := (R74Now-R74Mix[1])/(R64Now-R64Mix[1]);
  Intercept := R74Now - Slope*R64Now;
  i := 1;
  Tx := MaxX;
  repeat
    R64[i] := Tx;
    R74[i] := R64[i]*Slope + Intercept;
    ChPbLossU.Series[iPrimaryIsochron].AddXY(R64[i],R74[i]);
    //ShowMessage(IntToStr(i)+' '+FormatFloat('###.###',R64[i])+' '+FormatFloat('###.###',R74[i]));
    i := i + 1;
    Tx := Tx - StepSize1;
  until (Tx <= R64Mix[1]);
  ProgressBar1.Position := ProgressBar1.Position + 1;

  //Primary isochron thorogenic
  {
  sbPbLoss.Panels[1].Text := 'Primary isochron thorogenic';
  sbPbLoss.Refresh;
  Slope := (R84Now-R84Mix[1])/(R64Now-R64Mix[1]);
  Intercept := R84Now - Slope*R64Now;
  i := 1;
  Tx := MaxX;
  repeat
    R64[i] := Tx;
    R84[i] := R64[i]*Slope + Intercept;
    //ChPbLossTh.Series[iPrimaryIsochron].AddXY(R64[i],R84[i]);
    i := i + 1;
    Tx := Tx - StepSize1;
  until (Tx <= R64Mix[1]);
  ProgressBar1.Position := ProgressBar1.Position + 1;
  }

  //Primary isochron 204Pb206Pb - 207Pb/206Pb
  sbPbLoss.Panels[1].Text := 'Primary isochron 4/6 - 7/6';
  sbPbLoss.Refresh;
  Slope := (R74Now-R74Mix[1])/(R64Now-R64Mix[1]);
  Intercept := R74Now - Slope*R64Now;
  i := 1;
  Tx := MaxX;
  //ShowMessage(IntToStr(i)+' '+FormatFloat('###.###',R76Now)+' '+FormatFloat('###.###',R46Now));
  //ShowMessage(IntToStr(i)+' '+FormatFloat('###.###',R64Mix[1])+' '+FormatFloat('###.###',R74Mix[1]));
  //ShowMessage(IntToStr(i)+' '+FormatFloat('###.#####',Slope)+' '+FormatFloat('###.#####',Intercept));
  //ShowMessage(IntToStr(i)+' '+FormatFloat('###.###',MaxX)+' '+FormatFloat('###.#####',StepSize1));
  repeat
    R64[i] := Tx;
    R74[i] := R64[i]*Slope + Intercept;
    if (R64[i] > 0.0) then ChPb46Pb76.Series[iPrimaryIsochron].AddXY(1.0/R64[i],R74[i]/R64[i]);
    i := i + 1;
    Tx := Tx - StepSize1;
  until (Tx <= R64Mix[1]);
  ProgressBar1.Position := ProgressBar1.Position + 1;

  //Primary isochron 232Th/238U - 208Pb*/206Pb*
  {
  i := 1;
  Tx := 2.0;
  StepSize2 := (ToK - Tx)/100.0;
  repeat
    R64[i]:=SmpMu[1]*(exp(Decay238*T0*1.0e+6)-exp(Decay238*T1*1.0e+6));
    R84[i]:=Tx*SmpMu[1]*(exp(Decay232*T0*1.0e+6)-exp(Decay232*T1*1.0e+6));
    //if (R64[i] > 0.0) then ChThUPb86.Series[iPrimaryIsochron].AddXY(Tx,R84[i]/R64[i]);
    i := i + 1;
    Tx := Tx + StepSize2;
  until (Tx >= ToK);
  ProgressBar1.Position := ProgressBar1.Position + 1;
  }

  sbPbLoss.Panels[1].Text := 'Secondary isochron uranogenic';
  sbPbLoss.Refresh;
  //secondary isochron uranogenic
  Tx := 0.0;
  i := 1;
  SmpMu[1] := (ToMu + FromMu)/2.0;
  NewMu[1] := SmpMu[1] + SmpMu[1]*(DelMuHi + DelMuLo)/200.0;
  if (NewMu[1] < 0.0) then NewMu[1] := 0.0;
  R64[1]:=R64Mix[1]
          +SmpMu[1]*(exp(Decay238*T0*1.0e+6)-exp(Decay238*T1*1.0e+6))
          +NewMu[1]*(exp(Decay238*T1*1.0e+6)-exp(Decay238*Tx*1.0e+6));
  R74[1]:=R74Mix[1]
          +SmpMu[1]/Factor*(exp(Decay235*T0*1.0e+6)-exp(Decay235*T1*1.0e+6))
          +NewMu[1]/Factor*(exp(Decay235*T1*1.0e+6)-exp(Decay235*Tx*1.0e+6));
  R64Now := R64[1];
  R74Now := R74[1];
  Tx := MaxX;
  if ((R64Now - R64Start2) > 1.0e-9) then
    Slope := (R74Now-R74Start2)/(R64Now-R64Start2)
  else
    Slope := 0.001;
  Intercept := R74Now - Slope*R64Now;
  //ShowMessage(IntToStr(i)+' '+FormatFloat('###.###',R64Now)+' '+FormatFloat('###.###',R64Mix[1]));
  //ShowMessage(IntToStr(i)+' '+FormatFloat('###.###',R74Now)+' '+FormatFloat('###.###',R74Mix[1]));
  //ShowMessage(IntToStr(i)+' '+FormatFloat('###.#####',Slope)+' '+FormatFloat('###.#####',Intercept));
  //ShowMessage(IntToStr(i)+' '+FormatFloat('###.###',MaxX)+' '+FormatFloat('###.#####',StepSize1));
  repeat
    R64[i] := Tx;
    R74[i] := R64[i]*Slope + Intercept;
    ChPbLossU.Series[iSecondaryIsochron].AddXY(R64[i],R74[i]);
    i := i + 1;
    Tx := Tx - StepSize1;
  until (Tx <= R64Start2);
  ProgressBar1.Position := ProgressBar1.Position + 1;

  //secondary isochron thorogenic
  {
  Tx := 0.0;
  i := 1;
  SmpMu[1] := (ToMu + FromMu)/2.0;
  NewMu[1] := SmpMu[1] + SmpMu[1]*(DelMuHi + DelMuLo)/200.0;
  SmpK[1] := (ToK + FromK)/2.0;
  NewK[1] := SmpK[1] + SmpK[1]*(DelKHi + DelKLo)/200.0;
  R64[1]:=R64Mix[1]
                +SmpMu[1]*(exp(Decay238*T0*1.0e+6)-exp(Decay238*T1*1.0e+6))
                +NewMu[1]*(exp(Decay238*T1*1.0e+6)-exp(Decay238*Tx*1.0e+6));
  R84[1]:=R84Mix[1]
           +SmpK[1]*SmpMu[1]*(exp(Decay232*T0*1.0e+6)-exp(Decay232*T1*1.0e+6))
           +NewK[1]*NewMu[1]*(exp(Decay232*T1*1.0e+6)-exp(Decay232*Tx*1.0e+6));
  Tx := MaxX;
  if ((R64Now - R64Start2) > 1.0e-9) then
    Slope := (R84Now-R84Start2)/(R64Now-R64Start2)
  else
    Slope := 0.001;
  Intercept := R84Now - Slope*R64Now;
  repeat
    R64[i] := Tx;
    R84[i] := R64[i]*Slope + Intercept;
    //ChPbLossTh.Series[iSecondaryIsochron].AddXY(R64[i],R84[i]);
    i := i + 1;
    Tx := Tx - StepSize1;
  until (Tx <= R64Start2);
  ProgressBar1.Position := ProgressBar1.Position + 1;
  }

  //Secondary isochron 204Pb206Pb - 207Pb/206Pb
  Tx := 0.0;
  i := 1;
  SmpMu[1] := (ToMu + FromMu)/2.0;
  NewMu[1] := SmpMu[1] + SmpMu[1]*(DelMuHi + DelMuLo)/200.0;
  R64[1]:=R64Mix[1]
                +SmpMu[1]*(exp(Decay238*T0*1.0e+6)-exp(Decay238*T1*1.0e+6))
                +NewMu[1]*(exp(Decay238*T1*1.0e+6)-exp(Decay238*Tx*1.0e+6));
  R74[1]:=R74Mix[1]
           +SmpMu[1]/Factor*(exp(Decay235*T0*1.0e+6)-exp(Decay235*T1*1.0e+6))
           +NewMu[1]/Factor*(exp(Decay235*T1*1.0e+6)-exp(Decay235*Tx*1.0e+6));
  Tx := MaxX;
  if ((R64Now - R64Start2) > 1.0e-9) then
    Slope := (R74Now-R74Start2)/(R64Now-R64Start2)
  else
    Slope := 0.001;
  Intercept := R74Now - Slope*R64Now;
  repeat
    R64[i] := Tx;
    R74[i] := R64[i]*Slope + Intercept;
    if (R64[i] > 0.0) then ChPb46Pb76.Series[iSecondaryIsochron].AddXY(1.0/R64[i],R74[i]/R64[i]);
    i := i + 1;
    Tx := Tx - StepSize1;
  until (Tx <= R64Start2);
  ProgressBar1.Position := ProgressBar1.Position + 1;

  //Secondary isochron 232Th/238U - 208Pb*/206Pb*
  {
  i := 1;
  Tx := 2.0;
  repeat
       R64[i]:=SmpMu[1]*(exp(Decay238*T0*1.0e+6)-exp(Decay238*T1*1.0e+6))
                +NewMu[1]*(exp(Decay238*T1*1.0e+6)-1.0);
       R84[i]:=SmpK[1]*SmpMu[1]*(exp(Decay232*T0*1.0e+6)-exp(Decay232*T1*1.0e+6))
                +Tx*NewMu[1]*(exp(Decay232*T1*1.0e+6)-1.0);
    //if (R64[i] > 0.0) then ChThUPb86.Series[iSecondaryIsochron].AddXY(Tx,R84[i]/R64[i]);
    i := i + 1;
    Tx := Tx + StepSize2;
  until (Tx >= ToK);
  ProgressBar1.Position := ProgressBar1.Position + 1;
  }

  //Secondary isochron 206Pb204Pb - 208Pb*/206Pb*
  {
  sbPbLoss.Panels[1].Text := 'Secondary isochron 206Pb/204Pb vs 208Pb*/206Pb*';
  sbPbLoss.Refresh;
  Tx := 0.0;
  i := 1;
  SmpMu[1] := (ToMu + FromMu)/2.0;
  NewMu[1] := SmpMu[1] + SmpMu[1]*(DelMuHi + DelMuLo)/200.0;
  R64[1]:=R64Mix[1]
                +SmpMu[1]*(exp(Decay238*T0*1.0e+6)-exp(Decay238*T1*1.0e+6))
                +NewMu[1]*(exp(Decay238*T1*1.0e+6)-exp(Decay238*Tx*1.0e+6));
  R84[1]:=SmpK[1]*SmpMu[1]*(exp(Decay232*T0*1.0e+6)-exp(Decay232*T1*1.0e+6))
                +Tx*NewMu[1]*(exp(Decay232*T1*1.0e+6)-1.0);
  Tx := MaxX;
  repeat
    R64[i] := Tx;
    R84[1]:=SmpK[1]*SmpMu[1]*(exp(Decay232*T0*1.0e+6)-exp(Decay232*T1*1.0e+6))
                +Tx*NewMu[1]*(exp(Decay232*T1*1.0e+6)-1.0);
    //if (R64[i] > 0.0) then ChPb64Pb86.Series[iSecondaryIsochron].AddXY(R64[i],R84[i]/R64[i]);
    i := i + 1;
    Tx := Tx - R64Start2/StepSize1;
  until (Tx <= R64Start2);
  ProgressBar1.Position := ProgressBar1.Position + 1;
  }

  //Samples
  sbPbLoss.Panels[1].Text := 'Samples';
  sbPbLoss.Refresh;
  if ((cbShowSamples.Checked) and (dmUPb.cdsPbModel.RecordCount > 0)) then
  begin
    dmUPb.cdsPbModel.DisableControls;
    dmUPb.cdsPbModel.First;
      for i := 1 to dmUPb.cdsPbModel.RecordCount do
      begin
        if (UpperCase(dmUPb.cdsPbModelShow.AsString) = 'Y') then
        begin
          ChUPb.Series[iSamples].AddXY(dmUPb.cdsPbModelU238Pb204.AsFloat,dmUPb.cdsPbModelPb206Pb204.AsFloat);
          ChPbLossU.Series[iSamples].AddXY(dmUPb.cdsPbModelPb206Pb204.AsFloat,dmUPb.cdsPbModelPb207Pb204.AsFloat);
          ChPbLossTh.Series[iSamples].AddXY(dmUPb.cdsPbModelPb206Pb204.AsFloat,dmUPb.cdsPbModelPb208Pb204.AsFloat);
          if (dmUPb.cdsPbModelPb206Pb204.AsFloat > 0.0) then ChPb46Pb76.Series[iSamples].AddXY(1.0/dmUPb.cdsPbModelPb206Pb204.AsFloat,dmUPb.cdsPbModelPb207Pb206.AsFloat);
          //if (dmUPb.cdsPbModelPb206Pb204.AsFloat <> R64Mix[1]) then
            //ChPb46Pb76.Series[iSamples].AddXY(1.0/(dmUPb.cdsPbModelPb206Pb204.AsFloat-R64Mix[1]),(dmUPb.cdsPbModelPb207Pb204.AsFloat-R74Mix[1])/(dmUPb.cdsPbModelPb206Pb204.AsFloat-R64Mix[1]));
          MuSource[i] := CalcMu(T1,ModelStart,Model64_0,Model74_0,dmUPb.cdsPbModelPb206Pb204.AsFloat,dmUPb.cdsPbModelPb207Pb204.AsFloat);
          ChMu1MuSource.Series[iMu+2].AddXY(SmpMu[1],MuSource[i]);
          if (dmUPb.cdsPbModelPb206Pb204.AsFloat > 0.0) then
            ChPb76Pb86.Series[iSamples].AddXY(dmUPb.cdsPbModelPb207Pb206.AsFloat,dmUPb.cdsPbModelPb208Pb206.AsFloat);
          if ((dmUPb.cdsPbModelPb206Pb204.AsFloat > 0.0) and (dmUPb.cdsPbModelPb208Pb204.AsFloat > 0.0)) then
          begin
            KappaCalc := (dmUPb.cdsPbModelPb208Pb204.AsFloat-Model84_CD)/(dmUPb.cdsPbModelPb206Pb204.AsFloat-Model64_CD)*(exp(Decay238*TCD*1.0e6)-1.0)/(exp(Decay232*TCD*1.0e6)-1.0);
            ChAgeThU.Series[iSamples].AddXY(0.0,KappaCalc);
            ChThUPb86.Series[iSamples].AddXY(KappaCalc,(dmUPb.cdsPbModelPb208Pb204.AsFloat-Model84_CD)/(dmUPb.cdsPbModelPb206Pb204.AsFloat-Model64_CD));
          end;
          {need to correct code in the next few lines}
          {
          ChPb76Pb86.Column := 18;
          if (dmUPb.cdsPbModelPb206Pb204.AsFloat <> R64Mix[1]) then
            ChPb76Pb86.Data := FormatFloat('#.00000',(dmUPb.cdsPbModelPb208Pb204.AsFloat-R84Mix[1])/(dmUPb.cdsPbModelPb206Pb204.AsFloat-R64Mix[1]))
            //ChPb76Pb86.Data := FormatFloat('#.00000',(dmUPb.cdsPbModelPb208Pb204.AsFloat-R84Mix[1])/(dmUPb.cdsPbModelPb206Pb204.AsFloat-R64Mix[1]))
          else
            ChPb76Pb86.Data := '';
          ChThUPb86.Row := i;
          ChThUPb86.Column := 17;
          if (dmUPb.cdsPbModelU238Th232.AsFloat > 0.0) then
            ChThUPb86.Data := FormatFloat('#.00000',dmUPb.cdsPbModelU238Th232.AsFloat)
          else
            ChThUPb86.Data := '';
          }
          {need to correct code in the next few lines}
          {
          ChThUPb86.Column := 18;
          if (dmUPb.cdsPbModelPb206Pb204.AsFloat <> R64Mix[1]) then
            ChThUPb86.Data := FormatFloat('#.00000',(dmUPb.cdsPbModelPb208Pb204.AsFloat-Model84_0)/(dmUPb.cdsPbModelPb206Pb204.AsFloat-Model64_0))
            //ChThUPb86.Data := FormatFloat('#.00000',(dmUPb.cdsPbModelPb208Pb204.AsFloat-R84Mix[1])/(dmUPb.cdsPbModelPb206Pb204.AsFloat-R64Mix[1]))
          else
            ChThUPb86.Data := '';
          ChPb64Pb86.Row := i;
          ChPb64Pb86.Column := 17;
          ChPb64Pb86.Data := FormatFloat('#.00000',dmUPb.cdsPbModelPb206Pb204.AsFloat);
          }
          {need to correct code in the next few lines}
          {
          ChPb64Pb86.Column := 18;
          if (dmUPb.cdsPbModelPb206Pb204.AsFloat <> R64Mix[1]) then
            ChPb64Pb86.Data := FormatFloat('#.00000',(dmUPb.cdsPbModelPb208Pb204.AsFloat-R84Mix[1])/(dmUPb.cdsPbModelPb206Pb204.AsFloat-R64Mix[1]))
          else
            ChPb64Pb86.Data := '';
          }
        end;
       dmUPb.cdsPbModel.Next;
      end;
    dmUPb.cdsPbModel.First;
    dmUPb.cdsPbModel.EnableControls;
  end else
  begin
    //ShowMessage('Insufficient samples');
  end;
  ProgressBar1.Position := ProgressBar1.Position + 1;
  {
  ChUPb.Axes.Bottom.Minimum := SmpMu[1] - 0.02*SmpMu[1];
  ChUPb.Axes.Bottom.Maximum := MaxX + 0.02*MaxX;
  ChUPb.Axes.Left.Minimum := R64Mix[1] - 0.02*R64Mix[1];
  ChUPb.Axes.Left.Maximum := MaxY + 0.02*MaxY;
  ChPbLossU.Axes.Bottom.Minimum := R64Mix[1] - 0.02*R64Mix[1];
  ChPbLossU.Axes.Left.Minimum := R74Mix[1] - 0.02*R74Mix[1];
  ChPbLossU.Axes.Bottom.Maximum := MaxX + 0.02*MaxX;
  ChPbLossU.Axes.Left.Maximum := MaxY + 0.02*MaxY;
  ChPbLossU.Axes.Bottom.Minimum := R64Mix[1] - 0.02*R64Mix[1];
  ChPbLossU.Axes.Left.Minimum := R84Mix[1] - 0.02*R84Mix[1];
  ChPbLossU.Axes.Bottom.Maximum := MaxX + 0.02*MaxX;
  ChPbLossU.Axes.Left.Maximum := MaxY + 0.02*MaxY;
  }
  {
  ChPb64Pb86.Plot.Axis[ChAxisIdX,1].ValueScale.Minimum := R64Mix[1] - 0.02*R64Mix[1];
  ChPb64Pb86.Plot.Axis[ChAxisIdY,1].ValueScale.Minimum := 0.0;
  ChPb64Pb86.Plot.Axis[ChAxisIdX,1].ValueScale.Maximum := MaxX + 0.02*MaxX;
  ChPb64Pb86.Plot.Axis[ChAxisIdY,1].ValueScale.Maximum := 2.2;
  }
  ChUPb.Visible := true;
  ChPbLossU.Visible := true;
  ChPbLossTh.Visible := true;
  ChPb64Pb86.Visible := true;
  if cbUseMinMax.Checked then
  begin
    {
    ChUPb.Plot.Axis[ChAxisIdY,1].ValueScale.Minimum := Min64;
    ChUPb.Plot.Axis[ChAxisIdY,1].ValueScale.Maximum := Max64;
    }
    ChUPb.LeftAxis.SetMinMax(Min64,Max64);
    ChUPb.Visible := true;
    {
    ChPbLossU.Plot.Axis[ChAxisIdX,1].ValueScale.Minimum := Min64;
    ChPbLossU.Plot.Axis[ChAxisIdY,1].ValueScale.Minimum := Min74;
    ChPbLossU.Plot.Axis[ChAxisIdX,1].ValueScale.Maximum := Max64;
    ChPbLossU.Plot.Axis[ChAxisIdY,1].ValueScale.Maximum := Max74;
    }
    ChPbLossU.BottomAxis.SetMinMax(Min64,Max64);
    ChPbLossU.LeftAxis.SetMinMax(Min74,Max74);
    ChPbLossU.Visible := true;
    {
    ChPbLossTh.Plot.Axis[ChAxisIdX,1].ValueScale.Minimum := Min64;
    ChPbLossTh.Plot.Axis[ChAxisIdY,1].ValueScale.Minimum := Min84;
    ChPbLossTh.Plot.Axis[ChAxisIdX,1].ValueScale.Maximum := Max64;
    ChPbLossTh.Plot.Axis[ChAxisIdY,1].ValueScale.Maximum := Max84;
    }
    ChPbLossTh.BottomAxis.SetMinMax(Min64,Max64);
    ChPbLossTh.LeftAxis.SetMinMax(Min84,Max84);
    ChPbLossTh.Visible := true;
    {
    ChPb64Pb86.Plot.Axis[ChAxisIdX,1].ValueScale.Minimum := Min64;
    ChPb64Pb86.Plot.Axis[ChAxisIdY,1].ValueScale.Minimum := 0.0;
    ChPb64Pb86.Plot.Axis[ChAxisIdX,1].ValueScale.Maximum := Max64;
    ChPb64Pb86.Plot.Axis[ChAxisIdY,1].ValueScale.Maximum := 2.2;
    }
    ChPb64Pb86.BottomAxis.SetMinMax(Min64,Max64);
    ChPb64Pb86.LeftAxis.SetMinMax(0.0,2.2);
    ChPb64Pb86.Visible := true;
  end;
  if not cbUseMinMax.Checked then
  begin
    ChUPb.LeftAxis.Automatic := true;
    ChUPb.Visible := true;
    ChPbLossU.BottomAxis.Automatic := true;
    ChPbLossU.LeftAxis.Automatic := true;
    ChPbLossU.Visible := true;
    ChPbLossTh.BottomAxis.Automatic := true;
    ChPbLossTh.LeftAxis.Automatic := true;
    ChPbLossTh.Visible := true;
    ChPb64Pb86.BottomAxis.Automatic := true;
    ChPb64Pb86.LeftAxis.Automatic := true;
    ChPb64Pb86.Visible := true;
  end;
  {
  ChPb46Pb76.Plot.Axis[ChAxisIdX,1].ValueScale.Minimum := 0.00;
  ChPb46Pb76.Plot.Axis[ChAxisIdY,1].ValueScale.Minimum := 0.0;
  ChPb46Pb76.Plot.Axis[ChAxisIdX,1].ValueScale.Maximum := 0.1;
  ChPb46Pb76.Plot.Axis[ChAxisIdY,1].ValueScale.Maximum := 1.2;
  }
  ChPb46Pb76.BottomAxis.SetMinMax(0.0,0.1);
  ChPb46Pb76.LeftAxis.SetMinMax(0.0,1.2);
  ChPb46Pb76.Visible := true;
  {
  ChPb76Pb86.Plot.Axis[ChAxisIdX,1].ValueScale.Minimum := 0.00;
  ChPb76Pb86.Plot.Axis[ChAxisIdY,1].ValueScale.Minimum := 0.0;
  ChPb76Pb86.Plot.Axis[ChAxisIdX,1].ValueScale.Maximum := 1.1;
  ChPb76Pb86.Plot.Axis[ChAxisIdY,1].ValueScale.Maximum := 2.2;
  }
  ChPb76Pb86.BottomAxis.SetMinMax(0.0,1.3);
  ChPb76Pb86.LeftAxis.SetMinMax(0.0,2.8);
  ChPb76Pb86.Visible := true;
  {
  ChThUPb86.Plot.Axis[ChAxisIdX,1].ValueScale.Minimum := 0.00;
  ChThUPb86.Plot.Axis[ChAxisIdY,1].ValueScale.Minimum := 0.0;
  ChThUPb86.Plot.Axis[ChAxisIdX,1].ValueScale.Maximum := 5.0;
  ChThUPb86.Plot.Axis[ChAxisIdY,1].ValueScale.Maximum := 2.2;
  }
  ChThUPb86.BottomAxis.SetMinMax(0.0,5.0);
  ChThUPb86.LeftAxis.SetMinMax(0.0,2.2);
  ChThUPb86.Visible := true;
  {
  ChMu1Mu2.Plot.Axis[ChAxisIdX,1].ValueScale.Minimum := 0.0;
  ChMu1Mu2.Plot.Axis[ChAxisIdY,1].ValueScale.Minimum := 0.0;
  ChMu1Mu2.Plot.Axis[ChAxisIdX,1].ValueScale.Maximum := MaxMu + 0.02*MaxMu;
  ChMu1Mu2.Plot.Axis[ChAxisIdY,1].ValueScale.Maximum := MaxMu + 0.02*MaxMu;
  }
  ChMu1Mu2.BottomAxis.SetMinMax(0.0,MaxMu+0.02*MaxMu);
  ChMu1Mu2.LeftAxis.SetMinMax(0.0,MaxMu+0.02*MaxMu);
  ChMu1Mu2.Visible := true;
  {
  ChK1K2.Plot.Axis[ChAxisIdX,1].ValueScale.Minimum := 0.0;
  ChK1K2.Plot.Axis[ChAxisIdY,1].ValueScale.Minimum := 0.0;
  ChK1K2.Plot.Axis[ChAxisIdX,1].ValueScale.Maximum := MaxK + 0.02*MaxK;
  ChK1K2.Plot.Axis[ChAxisIdY,1].ValueScale.Maximum := MaxK + 0.02*MaxK;
  }
  ChK1K2.BottomAxis.SetMinMax(0.0,MaxK+0.02*MaxK);
  ChK1K2.LeftAxis.SetMinMax(0.0,MaxK+0.02*MaxK);
  ChK1K2.Visible := true;
  {
  ChMu1MuSource.Plot.Axis[ChAxisIdX,1].ValueScale.Minimum := 0.0;
  ChMu1MuSource.Plot.Axis[ChAxisIdY,1].ValueScale.Minimum := 0.0;
  ChMu1MuSource.Plot.Axis[ChAxisIdX,1].ValueScale.Maximum := MaxMuSource + 0.02*MaxMuSource;
  ChMu1MuSource.Plot.Axis[ChAxisIdY,1].ValueScale.Maximum := MaxMuSource + 0.02*MaxMuSource;
  }
  ChMu1MuSource.BottomAxis.SetMinMax(0.0,MaxMuSource+0.02*MaxMuSource);
  ChMu1MuSource.LeftAxis.SetMinMax(0.0,MaxMuSource+0.02*MaxMuSource);
  ChMu1MuSource.Visible := true;
  ProgressBar1.Position := ProgressBar1.Position + 1;
  ProgressBar1.Visible := false;
  sbPbLoss.Panels[1].Text := Title;
  sbPbLoss.Refresh;
  pc1.ActivePage := TabUranogenic;
end;

procedure TfmPbLoss3.FormShow(Sender: TObject);
begin
  GetIniFile;
  PbModelFileName := cdsPath+'\PBMODEL.XML';
  UpdateModelParameters;
  ChUPb.Visible := false;
  ChPbLossU.Visible := false;
  ChPbLossTh.Visible := false;
  ChMu1Mu2.Visible := false;
  ChK1K2.Visible := false;
  ChMu1MuSource.Visible := false;
  ChPb46Pb76.Visible := false;
  ChPb76Pb86.Visible := false;
  ChThUPb86.Visible := false;
  ChPb64Pb86.Visible := false;
  pc1.ActivePage := tsDefinition;
  Title := '';
  //dmUPb.cdsPbModel.Open;
  dmUPb.cdsPbModel.LoadFromFile(PbModelFileName);
  FromRowValueString := '2';
  ToRowValueString := '2';
end;

procedure TfmPbLoss3.sbFileOpenClick(Sender: TObject);
begin
  if (OpenDialog1.Execute) then
  begin
    ParamFileName := OpenDialog1.FileName;
    AssignFile(OutFile,ParamFileName);
    Reset(OutFile);
    Readln(OutFile,T0,R64_0,R74_0,R84_0);
    Readln(OutFile,T1,TSigma);
    Readln(OutFile,DelMuLo,DelMuHi,DelKLo,DelKHi);
    Readln(OutFile,ChangeRand);
    Readln(OutFile,NPoints,FromMu,ToMu,FromK,ToK);
    Readln(OutFile,MuRand);
    Readln(OutFile,Weight6,Weight7,R_Coeff,Weight8,R_Coeff8);
    CloseFile(OutFile);
    ParamFileName := ExtractFileName(OpenDialog1.FileName);
    ParamFilePath := ExtractFilePath(OpenDialog1.FileName);
    OpenDialog1.InitialDir := ParamFilePath;
    SetLength(ParamFileName,Length(ParamFileName) - 4);
    UpdateModelParameters;
  end else
  begin
    ParamFileName := '';
  end;
end;

procedure TfmPbLoss3.sbFileSaveClick(Sender: TObject);
begin
  SaveDialog1.FileName := Title+SaveDialog1.DefaultExt;
  if SaveDialog1.Execute then
  begin
    ParamFileName := SaveDialog1.FileName;
    AssignFile(OutFile,ParamFileName);
    Rewrite(OutFile);
    Writeln(OutFile,T0,' ',R64_0,' ',R74_0,' ',R84_0);
    Writeln(OutFile,T1,' ',TSigma);
    Writeln(OutFile,DelMuLo,' ',DelMuHi,' ',DelKLo,' ',DelKHi);
    Writeln(OutFile,ChangeRand);
    Writeln(OutFile,NPoints,' ',FromMu,' ',ToMu,' ',FromK,' ',ToK);
    Writeln(OutFile,MuRand);
    Writeln(OutFile,Weight6,' ',Weight7,' ',R_Coeff,' ',Weight8,' ',R_Coeff8);
    CloseFile(OutFile);
    ParamFileName := ExtractFileName(SaveDialog1.FileName);
    ParamFilePath := ExtractFilePath(SaveDialog1.FileName);
    SaveDialog1.InitialDir := ParamFilePath;
    SetLength(ParamFileName,Length(ParamFileName) - 4);
  end else
  begin
    ParamFileName := '';
  end;
end;

procedure TfmPbLoss3.sbFileSaveGEODATEClick(Sender: TObject);
begin
  if (Title <> '') then
  begin
    SaveDialogGEODATE.FileName := Title+SaveDialogGEODATE.DefaultExt;
    if SaveDialogGEODATE.Execute then
    begin
      AssignFile(York_file, SaveDialogGEODATE.FileName);
      Rewrite(York_file);
      Write_File_Data;
      CloseFile(York_file);
    end;
  end;
end;

procedure TfmPbLoss3.eT1Change(Sender: TObject);
begin
  Title := '';
end;

function TfmPbLoss3.CalcMu (Age, StartAge,
                            Model64, Model74,
                            R64, R74 : double) : double;
{Mu for individual points}
var
  A, B, C, G,
  U, V, W : double;

function Age5 ( Age : double) : double;
begin
  Age5:=Exp(Decay235*Age);
end;

function Age8 ( Age : double) : double;
begin
  Age8:=Exp(Decay238*Age);
end;

begin
  Age:=Age*1.0e6;
  StartAge := StartAge*1.0e6;
  A:=Age5(Age)-1.0;
  B:=Age8(Age)-1.0;
  C:=Age5(StartAge)-Age5(Age);
  G:=Age8(StartAge)-Age8(Age);
  V:=(A/B)*(Model64-R64)/137.88 + R74-Model74;
  W:=C/137.88 - (A/B)*G/137.88;
  U:=V/W;
  Result:=U;
end;


procedure TfmPbLoss3.ChUPbDblClick(Sender: TObject);
begin
  ChUPb.PrintProportional := true;
end;

procedure TfmPbLoss3.rbUniformStartClick(Sender: TObject);
begin
  UpdateModelParameters;
  gbStartUniform.Visible := true;
  gbStartMix.Visible := false;
end;

procedure TfmPbLoss3.rbMixedStartClick(Sender: TObject);
begin
  UpdateModelParameters;
  gbStartMix.Visible := true;
  gbStartUniform.Visible := false;
end;

procedure TfmPbLoss3.rbTwoStageClick(Sender: TObject);
begin
    ModelStart := 3700.0;
    ModelMu := 9.74;
    Model64_0 := 11.152;
    Model74_0 := 12.998;
    ModelK := 3.782;
    Model84_0 := 31.230;
end;

procedure TfmPbLoss3.rbSingleStageClick(Sender: TObject);
begin
    ModelStart := 4570.0;
    ModelMu := 7.19;
    Model64_0 := 9.307;
    Model74_0 := 10.294;
    ModelK := 4.619;
    Model84_0 := 29.487;
end;

procedure TfmPbLoss3.sbCalculateModelValuesClick(Sender: TObject);
var
  Model64AtAge, Model74AtAge, Model84AtAge : double;
  iCode : integer;
begin
  Val(eT1.Text,T0,iCode);
  Model64AtAge := Model64_0 + ModelMu*(exp(Decay238*ModelStart*1.0e6)-exp(Decay238*T0*1.0e6));
  Model74AtAge := Model74_0 + (ModelMu/Factor)*(exp(Decay235*ModelStart*1.0e6)-exp(Decay235*T0*1.0e6));
  Model84AtAge := Model84_0 + (ModelMu*ModelK)*(exp(Decay232*ModelStart*1.0e6)-exp(Decay232*T0*1.0e6));
  eR64_0.Text :=   FormatFloat('  ##0.000',Model64AtAge);
  eR74_0.Text :=   FormatFloat('  ##0.000',Model74AtAge);
  eR84_0.Text :=   FormatFloat('  ##0.000',Model84AtAge);
end;

procedure TfmPbLoss3.FormClose(Sender: TObject; var MyAction: TCloseAction);
begin
  dmUPb.cdsPbModel.FileName := cdsPath+'PbModel.xml';
  dmUPb.cdsPbModel.SaveToFile(PbModelFileName,dfXML);
end;

procedure TfmPbLoss3.FormCreate(Sender: TObject);
var
  Style: String;
  Item: TMenuItem;
begin
  //Add child menu items based on available styles.
  GetIniFile;
  TStyleManager.TrySetStyle(GlobalChosenStyle);
  for Style in TStyleManager.StyleNames do
  begin
    Item := TMenuItem.Create(Styles1);
    Item.Caption := Style;
    Item.OnClick := StyleClick;
    if TStyleManager.ActiveStyle.Name = Style then
      Item.Checked := true;
    Styles1.Add(Item);
  end;
end;

procedure TfmPbLoss3.GetIniFile;
var
  AppIni   : TIniFile;
  PublicPath : string;
begin
  //PublicPath := TPath.GetPublicPath;
  //Used to use CSIDL_COMMON_APPDATA but some users do not have access to this
  //and don't know how to change their system settings and permissions to all
  //software to write to this path.
  //Now changed to use CSIDL_COMMON_DOCUMENTS which automatically permits
  //all users to have both read and write permission
  PublicPath := TPath.GetHomePath;
  dmUPb.CommonFilePath := IncludeTrailingPathDelimiter(PublicPath) + 'EggSoft\';
  dmUPb.IniFilename := dmUPb.CommonFilePath + 'UPbModel.INI';
  AppIni := TIniFile.Create(dmUPb.IniFilename);
  try
    GlobalChosenStyle := 'Windows';
    cdsPath := AppIni.ReadString('FilePath','XMLfile',dmUPb.CommonFilePath+'UPbModel\Data\');
    DataPath := AppIni.ReadString('FilePath','Sample file',dmUPb.CommonFilePath+'UPbModel\Data\');
    GlobalChosenStyle := AppIni.ReadString('Styles','Chosen style','Windows');
    if (GlobalChosenStyle = '') then GlobalChosenStyle := 'Windows';
    dmUPb.ChosenStyle := GlobalChosenStyle;
    SampleNoStr := AppIni.ReadString('Sample Spreadsheet Columns','SampleNo','A');
    ShowDataStr := AppIni.ReadString('Sample Spreadsheet Columns','ShowData','B');
    s238204Str := AppIni.ReadString('Sample Spreadsheet Columns','238U204Pb','C');
    s232238Str := AppIni.ReadString('Sample Spreadsheet Columns','232Th238U','D');
    s206204Str := AppIni.ReadString('Sample Spreadsheet Columns','206Pb204Pb','E');
    s207204Str := AppIni.ReadString('Sample Spreadsheet Columns','207Pb204Pb','F');
    s208204Str := AppIni.ReadString('Sample Spreadsheet Columns','208Pb204Pb','G');
    s207206Str := AppIni.ReadString('Sample Spreadsheet Columns','207Pb206Pb','H');
    s208206Str := AppIni.ReadString('Sample Spreadsheet Columns','208Pb206Pb','I');
  finally
    AppIni.Free;
  end;
end;

procedure TfmPbLoss3.SetIniFile;
var
  AppIni   : TIniFile;
  PublicPath : string;
  //zpath: array [0..MAX_PATH] of char;
begin
  //PublicPath := TPath.GetPublicPath;
  //Used to use CSIDL_COMMON_APPDATA but some users do not have access to this
  //and don't know how to change their system settings and permissions to all
  //software to write to this path.
  //Now changed to use CSIDL_COMMON_DOCUMENTS which automatically permits
  //all users to have both read and write permission
  PublicPath := TPath.GetHomePath;
  dmUPb.CommonFilePath := IncludeTrailingPathDelimiter(PublicPath) + 'EggSoft\';
  dmUPb.IniFilename := dmUPb.CommonFilePath + 'UPbModel.INI';
  AppIni := TIniFile.Create(dmUPb.IniFilename);
  try
    AppIni.WriteString('FilePath','XMLfile',cdsPath);
    AppIni.WriteString('FilePath','Sample file',DataPath);
    AppIni.WriteString('Styles','Chosen style',GlobalChosenStyle);
    AppIni.WriteString('Sample Spreadsheet Columns','SampleNo',SampleNoStr);
    AppIni.WriteString('Sample Spreadsheet Columns','ShowData',ShowDataStr);
    AppIni.WriteString('Sample Spreadsheet Columns','238U204Pb',s238204Str);
    AppIni.WriteString('Sample Spreadsheet Columns','232Th238U',s232238Str);
    AppIni.WriteString('Sample Spreadsheet Columns','206Pb204Pb',s206204Str);
    AppIni.WriteString('Sample Spreadsheet Columns','207Pb204Pb',s207204Str);
    AppIni.WriteString('Sample Spreadsheet Columns','208Pb204Pb',s208204Str);
    AppIni.WriteString('Sample Spreadsheet Columns','207Pb206Pb',s207206Str);
    AppIni.WriteString('Sample Spreadsheet Columns','208Pb206Pb',s208206Str);
  finally
    AppIni.Free;
  end;
end;

procedure TfmPbLoss3.About1Click(Sender: TObject);
begin
  AboutBox.ShowModal;
end;

procedure TfmPbLoss3.Importsampledata1Click(Sender: TObject);
var
  i : integer;
begin
  try
    ImportForm := TfmSheetImport.Create(Self);
    ImportForm.OpenDialogSprdSheet.FileName := 'Pbdata';
    ImportForm.ShowModal;
  finally
    ImportForm.Free;
  end;
  (*
  dmUPb.cdsPbModel.First;
  repeat
    ShowMessage(dmUPb.cdsPbModelSample_No.AsString+'***'+dmUPb.cdsPbModelPb206Pb204.AsString);
    dmUPb.cdsPbModel.Next;
  until dmUPb.cdsPbModel.Eof;
  dmUPb.cdsPbModel.First;
  *)
end;

procedure TfmPbLoss3.Printgraph1Click(Sender: TObject);
begin
  if (pc1.ActivePage = TabUPb) then ChUPb.PrintLandscape;
  if (pc1.ActivePage = TabUranogenic) then ChPbLossU.PrintLandscape;
  if (pc1.ActivePage = TabThorogenic) then ChPbLossTh.PrintLandscape;
  if (pc1.ActivePage = TabMu1Mu2) then ChMu1Mu2.PrintLandscape;
  if (pc1.ActivePage = TabK1K2) then ChK1K2.PrintLandscape;
  if (pc1.ActivePage = TabMu1MuSource) then ChMu1MuSource.PrintLandscape;
  if (pc1.ActivePage = tsPb46Pb76) then ChPb46Pb76.PrintLandscape;
  if (pc1.ActivePage = tsPb76Pb86) then ChPb76Pb86.PrintLandscape;
  if (pc1.ActivePage = tsThUPbrad) then ChThUPb86.PrintLandscape;
  if (pc1.ActivePage = ts6486) then ChPb64Pb86.PrintLandscape;
  if (pc1.ActivePage = tsAge64) then ChAgePb64.PrintLandscape;
  if (pc1.ActivePage = tsAge84) then ChAgePb84.PrintLandscape;
  if (pc1.ActivePage = tsAgeThU) then ChAgeThU.PrintLandscape;
end;

end.
