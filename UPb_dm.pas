unit UPb_dm;

interface

uses
  System.SysUtils, MidasLib, DB, System.Classes, Datasnap.DBClient,
  Vcl.BaseImageCollection, Vcl.ImageCollection,
  SVGIconImageCollection;

type
  TdmUPb = class(TDataModule)
    dsPbModel: TDataSource;
    cdsPbModel: TClientDataSet;
    cdsPbModelSample_No: TStringField;
    cdsPbModelShow: TStringField;
    cdsPbModelU238Pb204: TFloatField;
    cdsPbModelTh232U238: TFloatField;
    cdsPbModelPb206Pb204: TFloatField;
    cdsPbModelPb207Pb204: TFloatField;
    cdsPbModelPb208Pb204: TFloatField;
    cdsPbModelPb207Pb206: TFloatField;
    cdsPbModelPb208Pb206: TFloatField;
    ImageCollection1: TImageCollection;
    SVGIconImageCollection1: TSVGIconImageCollection;
  private
    { Private declarations }
  public
    { Public declarations }
    CommonFilePath : string;
    IniFilename : string;
    ChosenStyle : string;
  end;

var
  dmUPb: TdmUPb;

implementation

{$R *.dfm}

end.
