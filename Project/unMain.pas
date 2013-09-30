unit unMain;
  { TODO -oKAA -c2013.06.26 :
  - ���� ��� ������������� Delphi-AutoCAD: http://www.cadhouse.narod.ru/articles/acad/acad_connect.htm
  - �������� ������ (������ ���������� AutoCAD) http://forum.dwg.ru/showthread.php?t=12126
  }

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, AutoCAD_TLB, ComObj, kaaAcadConverter, Grids, ActiveX,
  ComCtrls;

type
  TfmMain = class(TForm)
    pnBtns: TPanel;
    btAutoCAD: TButton;
    OpenDialog: TOpenDialog;
    stgStatistics: TStringGrid;
    PageControl: TPageControl;
    tsGraph: TTabSheet;
    tsResults: TTabSheet;
    Memo: TMemo;
    PaintBox: TPaintBox;
    procedure btAutoCADClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure PaintBoxPaint(Sender: TObject);
  private
    FAutoCAD: TAutoCAD;
  private//Memo
    procedure AddMemoMsg(const AMessage: String; const APriorBlankLine: Boolean = False); overload;
    procedure AddMemoMsg(const AFormat: String; const AArgs: array of const; const APriorBlankLine: Boolean = False); overload;
  protected//���������� �������� AutoCAD
    //����������
    procedure _ExtractAutoCADFile();
    //�������
    procedure ExtractAcadClear();
    //Statistics
    procedure ExtractAcadStatistics();
    procedure ExtractAcadLayer(const AObjectNo: Integer; const ALayer: IAcadLayer);
    procedure ExtractAcadPoint(const AObjectNo: Integer; const AObject: IAcadPoint);
    //Objects
    procedure ExtractAcadBlock(const AObjectNo: Integer; const ABlock: IAcadBlockReference);
    procedure ExtractAcadHatch(const AObjectNo: Integer; const AObject: IAcadHatch);
    procedure ExtractAcadLine(const AObjectNo: Integer; const AObject: IAcadLine);
    procedure ExtractAcadLWPolyline(const AObjectNo: Integer; const AObject: IAcadLWPolyline);
    procedure ExtractAcad3dPolyline(const AObjectNo: Integer; const AObject: IAcad3dPolyline);
    procedure ExtractAcadCircle(const AObjectNo: Integer; const AObject: IAcadCircle);
    procedure ExtractAcadArc(const AObjectNo: Integer; const AObject: IAcadArc);
    procedure ExtractAcadEllipse(const AObjectNo: Integer; const AObject: IAcadEllipse);
    procedure ExtractAcadSpline(const AObjectNo: Integer; const AObject: IAcadSpline);
    procedure ExtractAcadText(const AObjectNo: Integer; const AObject: IAcadText);
    procedure ExtractAcadMText(const AObjectNo: Integer; const AObject: IAcadMText);
 end;{TfmMain}

var
  fmMain: TfmMain;
implementation

{$R *.dfm}
uses Math;

//Memo
procedure TfmMain.AddMemoMsg(const AMessage: String; const APriorBlankLine: Boolean = False);
begin
  if APriorBlankLine then Memo.Lines.Add('');
  Memo.Lines.Add(AMessage);
end;{AddMemoMsg}
procedure TfmMain.AddMemoMsg(const AFormat: String; const AArgs: array of const; const APriorBlankLine: Boolean = False); 
begin
  AddMemoMsg(Format(AFormat,AArgs), APriorBlankLine);
end;{AddMemoMsg}

//�������
procedure TfmMain.ExtractAcadClear();
var I: Integer;
begin
  Memo.Clear();
  FAutoCAD.Clear();
  stgStatistics.RowCount := 2;
  for I := 0 to stgStatistics.ColCount - 1 do
    stgStatistics.Cells[I, 1] := '';
end;{ExtractAcadClear}

//AutoCAD Statistics
procedure TfmMain.ExtractAcadStatistics();
var
  I: Integer;
begin
  stgStatistics.RowCount := 1 + Max(1, FAutoCAD.Statistics.Count);
  for I := 0 to FAutoCAD.Statistics.Count-1 do
  begin
    stgStatistics.Cells[0, I+1] := IntToStr(I + 1);
    stgStatistics.Cells[1, I+1] := IntToStr(FAutoCAD.Statistics[I].EntityType);
    stgStatistics.Cells[2, I+1] := FAutoCAD.Statistics[I].EntityName;
    stgStatistics.Cells[3, I+1] := IntToStr(FAutoCAD.Statistics[I].Count);
    stgStatistics.Cells[4, I+1] := IntToStr(FAutoCAD.Statistics[I].ImportedCount);
  end;{for}
end;{ExtractAcadStatistics}

//��������� c���� AutoCAD
procedure TfmMain.ExtractAcadLayer(const AObjectNo: Integer; const ALayer: IAcadLayer);
begin
  AddMemoMsg('���� Layer �%d', [AObjectNo], True);
  AddMemoMsg('�������� ���� Name: %s', [ALayer.Name], True);
  //����� ���������
  AddMemoMsg('������� LayerOn: %d', [Integer(ALayer.LayerOn = True)]);
  AddMemoMsg('���������� Freeze: %d', [Integer(ALayer.Freeze = True)]);
  AddMemoMsg('����������� Lock: %d', [Integer(ALayer.Lock = True)]);
  //����
  //AddMemoMsg('TrueColor.EntityColor: %d', [ALayer.TrueColor.EntityColor]);
  //AddMemoMsg('TrueColor.ColorName: %s', [ALayer.TrueColor.ColorName]);
  //AddMemoMsg('TrueColor.BookName: %s', [ALayer.TrueColor.BookName]);
  AddMemoMsg('���� ����� �� ��������� TrueColor.Red: %d', [ALayer.TrueColor.Red]);
  AddMemoMsg('���� ����� �� ��������� TrueColor.Green: %d', [ALayer.TrueColor.Green]);
  AddMemoMsg('���� ����� �� ��������� TrueColor.Blue: %d', [ALayer.TrueColor.Blue]);
  //AddMemoMsg('TrueColor.ColorMethod: %d', [ALayer.TrueColor.ColorMethod]);
  //AddMemoMsg('TrueColor.ColorIndex: %d', [ALayer.TrueColor.ColorIndex]);
  AddMemoMsg('��� ����� �� ��������� Linetype: %s', [ALayer.Linetype]);
  AddMemoMsg('��� ����� �� ��������� Lineweight: $%x', [ALayer.Lineweight]);
  //AddMemoMsg('����� ������ PlotStyleName: %s', [ALayer.PlotStyleName]);
  //AddMemoMsg('������ Plottable: %d', [Integer(ALayer.Plottable)]);
  AddMemoMsg('��������� Description: %s', [ALayer.Description]);
  //������ ���������
  //AddMemoMsg('ViewportDefault: %d', [Integer(ALayer.ViewportDefault)]);
  //AddMemoMsg('����������� Used: %d', [Integer(ALayer.Used)]);
  //AddMemoMsg('�������� Material: %s', [ALayer.Material]);
end;{ExtractAcadLayer}
procedure TfmMain.ExtractAcadPoint(const AObjectNo: Integer; const AObject: IAcadPoint);
begin
  AddMemoMsg('���� Layer: %s', [AObject.Layer]);
  AddMemoMsg('������ �%d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('��� ������� EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('������������ ������� EntityName: %s', [AObject.EntityName]);
  AddMemoMsg('��������� Visible: %d', [Integer(AObject.Visible)]);
  //
  AddMemoMsg('���������� Coords: (%f, %f, %f)', [Double(AObject.Coordinates[0]), Double(AObject.Coordinates[1]), Double(AObject.Coordinates[2])]);
  //AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  //AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  //AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  //AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  //AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //AddMemoMsg('������ Lineweight: %d', [AObject.Lineweight]);
  AddMemoMsg('������ Thickness: %f', [AObject.Thickness]);

  //AddMemoMsg('��� ����� LineType: %s', [AObject.Linetype]);
  //AddMemoMsg('������� ���� ����� LinetypeScale: %f', [AObject.LinetypeScale]);
  //AddMemoMsg('�������� Material: %s', [AObject.Material]);
  //AddMemoMsg('Normal: (%f; %f; %f)', [Double(AObject.Normal[0]), Double(AObject.Normal[1]), Double(AObject.Normal[2])]);
  //AddMemoMsg('����� ������ PlotStyleName: %s', [AObject.PlotStyleName]);
  //FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadPoint}

//��������� ������ AutoCAD
procedure TfmMain.ExtractAcadBlock(const AObjectNo: Integer; const ABlock: IAcadBlockReference);
begin
  AddMemoMsg('���� Block %d "%s"', [AObjectNo, ABlock.Name], True);
  AddMemoMsg('��� ������� EntityType: %d', [ABlock.EntityType]);
  AddMemoMsg('������������ ������� EntityName: %s', [ABlock.EntityName]);
  
  //����� ���������
  //����
  AddMemoMsg('TrueColor.EntityColor: %d', [ABlock.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [ABlock.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [ABlock.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [ABlock.TrueColor.Red]);
  AddMemoMsg('TrueColor.Green: %d', [ABlock.TrueColor.Green]);
  AddMemoMsg('TrueColor.Blue: %d', [ABlock.TrueColor.Blue]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [ABlock.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [ABlock.TrueColor.ColorIndex]);
  //����
  AddMemoMsg('���� Layer: %s', [ABlock.Layer]);
  //��� �����
  AddMemoMsg('��� ����� Linetype: %s', [ABlock.Linetype]);
  //������� ���� �����
  AddMemoMsg('������� ���� ����� LinetypeScale: %f', [ABlock.LinetypeScale]);
  //��� �����
  AddMemoMsg('��� ����� Lineweight: %d', [ABlock.Lineweight]);

  //3D ������������
  //��������
  AddMemoMsg('�������� Material: %s', [ABlock.Material]);

  //���������
  //���������
  AddMemoMsg('��������� InsertionPoint: (%f, %f, %f)', [Double(ABlock.InsertionPoint[0]), Double(ABlock.InsertionPoint[1]), Double(ABlock.InsertionPoint[2])]);
  //�������
  AddMemoMsg('������� ScaleFactor: (%f, %f, %f)', [ABlock.XScaleFactor, ABlock.YScaleFactor, ABlock.ZScaleFactor]);

  //������
  //���
  AddMemoMsg('��� EffectiveName: %s', [ABlock.EffectiveName]);
  //�������
  AddMemoMsg('������� Rotation: %f', [ABlock.Rotation]);
  //������������
  AddMemoMsg('������������ IsDynamicBlock: %d', [Integer(ABlock.IsDynamicBlock)]);
  //������� �����
  AddMemoMsg('������� ����� InsUnits: %s', [ABlock.InsUnits]);
  //����������� ������
  AddMemoMsg('����������� ������ InsUnitsFactor: %f', [ABlock.InsUnitsFactor]);

  //������ ��������
  //�������
  AddMemoMsg('Normal: (%f; %f; %f)', [Double(ABlock.Normal[0]), Double(ABlock.Normal[1]), Double(ABlock.Normal[2])]);
  //���������
  AddMemoMsg('��������� Visible: %d', [Integer(ABlock.Visible)]);
  //����� ������
  AddMemoMsg('����� ������ PlotStyleName: %s', [ABlock.PlotStyleName]);
  //��������
  AddMemoMsg('�������� HasAttributes: %d', [Integer(ABlock.HasAttributes)]);


  //
  AddMemoMsg('��� ������� ObjectName: %s', [ABlock.ObjectName]);
  AddMemoMsg('ID ������� ObjectID: %d', [ABlock.ObjectID]);

  FAutoCAD.Statistics.Add(ABlock.EntityType, ABlock.EntityName);
end;{ExtractAcadBlock}

//��������� �������� AutoCAD
procedure TfmMain.ExtractAcadHatch(const AObjectNo: Integer; const AObject: IAcadHatch);
begin
  AddMemoMsg('������ %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('��� ������� EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('������������ ������� EntityName: %s', [AObject.EntityName]);
  //����� ��������
  //����
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //����
  AddMemoMsg('���� Layer: %s', [AObject.Layer]);
  //��� �����
  AddMemoMsg('��� ����� LineType: %s', [AObject.Linetype]);
  //������� ���� �����
  AddMemoMsg('������� ���� ����� LinetypeScale: %f', [AObject.LinetypeScale]);
  //��� �����
  AddMemoMsg('��� ����� Lineweight: %d', [AObject.Lineweight]);

  //�������
  //��� �������
  AddMemoMsg('��� ������� PatternType: %d', [AObject.PatternType]);
  //��� �������
  AddMemoMsg('��� ������� PatternName: %s', [AObject.PatternName]);
  //������������

  //����
  AddMemoMsg('���� PatternAngle: %f', [AObject.PatternAngle]);
  //�������
  AddMemoMsg('������� PatternScale: %f', [AObject.PatternScale]);
  //�������� �����
  AddMemoMsg('�������� ����� Origin: (%f, %f)', [Double(AObject.Origin[0]), Double(AObject.Origin[1])]);
  //�������������

  //����� ������� ���������
  AddMemoMsg('����� ������� ��������� HatchStyle: %d', [AObject.HatchStyle]);
  //���������
  //�������
  AddMemoMsg('������� Elevation: %f', [AObject.Elevation]);

  //��������� ��������
  //�������
  AddMemoMsg('������� Area: %f', [AObject.Area]);
  //������� ���� �� ISO
  AddMemoMsg('������� ���� �� ISO ISOPenWidth: %d', [AObject.ISOPenWidth]);

  //������ ��������

  AddMemoMsg(' NumberOfLoops: %d', [AObject.NumberOfLoops]);
  AddMemoMsg(' PatternSpace: %f', [AObject.PatternSpace]);
  AddMemoMsg(' PatternDouble: %d', [Integer(AObject.PatternDouble)]);
  AddMemoMsg(' AssociativeHatch: %d', [Integer(AObject.AssociativeHatch)]);

  //AObject.GradientColor1.
  //AObject.GradientColor2
  //AObject.GradientAngle
  //AObject.GradientCentered
  //AObject.GradientName
  //AObject.BackgroundColor

  //��� ������� �������
  AddMemoMsg('��� ������� ������� HatchObjectType: %d', [AObject.HatchObjectType]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadHatch}

procedure TfmMain.ExtractAcadLine(const AObjectNo: Integer; const AObject: IAcadLine);
begin
  AddMemoMsg('������ %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('��� ������� EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('������������ ������� EntityName: %s', [AObject.EntityName]);

  //����� ��������
  //����
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //����
  AddMemoMsg('���� Layer: %s', [AObject.Layer]);
  //��� �����
  AddMemoMsg('��� ����� LineType: %s', [AObject.Linetype]);
  //������� ���� �����
  AddMemoMsg('������� ���� ����� LinetypeScale: %f', [AObject.LinetypeScale]);
  //��� �����
  AddMemoMsg('��� ����� Lineweight: %d', [AObject.Lineweight]);
  //������ 3D
  AddMemoMsg('������3D Thickness: %f', [AObject.Thickness]);
        
  //3D ������������
  //��������
  AddMemoMsg('�������� Material: %s', [AObject.Material]);

  //���������
  //��������� �����
  AddMemoMsg('��������� ����� StartPoint: (%f, %f, %f)', [Double(AObject.StartPoint[0]), Double(AObject.StartPoint[1]), Double(AObject.StartPoint[2])]);
  //�������� �����
  AddMemoMsg('�������� ����� EndPoint: (%f, %f, %f)', [Double(AObject.EndPoint[0]), Double(AObject.EndPoint[1]), Double(AObject.EndPoint[2])]);

  //��������� ��������
  //������
  AddMemoMsg('������ Delta: (%f, %f, %f)', [Double(AObject.Delta[0]), Double(AObject.Delta[1]), Double(AObject.Delta[2])]);
  //�����
  AddMemoMsg('����� Length: %f', [AObject.Length]);
  //����
  AddMemoMsg('���� Angle: %f', [AObject.Angle]);

  //������ ��������
  //�������
  AddMemoMsg('Normal: (%f; %f; %f)', [Double(AObject.Normal[0]), Double(AObject.Normal[1]), Double(AObject.Normal[2])]);
  //���������
  AddMemoMsg('��������� Visible: %d', [Integer(AObject.Visible)]);
  //����� ������
  AddMemoMsg('����� ������ PlotStyleName: %s', [AObject.PlotStyleName]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadLine}


procedure TfmMain.ExtractAcadLWPolyline(const AObjectNo: Integer; const AObject: IAcadLWPolyline);
var
  ACoords: OleVariant;
  ACoordsCount, I: Integer;
begin
  AddMemoMsg('������ %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('��� ������� EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('������������ ������� EntityName: %s', [AObject.EntityName]);

  //����� ��������
  //����
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);

  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //����
  AddMemoMsg('���� Layer: %s', [AObject.Layer]);
  //��� �����
  AddMemoMsg('��� ����� LineType: %s', [AObject.Linetype]);
  //������� ���� �����
  AddMemoMsg('������� ���� ����� LinetypeScale: %f', [AObject.LinetypeScale]);
  //��� �����
  AddMemoMsg('��� ����� Lineweight: %d', [AObject.Lineweight]);
  //������ 3D
  AddMemoMsg('������3D Thickness: %f', [AObject.Thickness]);

  //3D ������������
  //��������
  AddMemoMsg('�������� Material: %s', [AObject.Material]);

  //���������
  //���������� ������
  ACoords := AObject.Coordinates;
  ACoordsCount := (VarArrayHighBound(ACoords, 1) - VarArrayLowBound(ACoords, 1) + 1) div 2;
  AddMemoMsg('ACoordsCount: %d', [ACoordsCount]);
  for I := 0 to ACoordsCount - 1 do
    AddMemoMsg('Coordinates[%d]: (%f; %f)', [I+1, Double(ACoords[2*I]), Double(ACoords[2*I+1])]);
  //���������� ������
  AddMemoMsg('���������� ������ ConstantWidth: %f', [AObject.ConstantWidth]);
  //�������
  AddMemoMsg('������� Elevation: %f', [AObject.Elevation]);

  //��������� ��������
  //�������
  AddMemoMsg('������� Area: %f', [AObject.Area]);
  //�����
  AddMemoMsg('����� Length: %f', [AObject.Length]);

  //������
  //��������
  AddMemoMsg('�������� Closed: %d', [Integer(AObject.Closed)]);
  //��������� ���� �����
  AddMemoMsg('��������� ���� ����� LinetypeGeneration: %d', [Integer(AObject.LinetypeGeneration)]);

  //������ ��������
  //�������
  AddMemoMsg('Normal: (%f; %f; %f)', [Double(AObject.Normal[0]), Double(AObject.Normal[1]), Double(AObject.Normal[2])]);
  //���������
  AddMemoMsg('��������� Visible: %d', [Integer(AObject.Visible)]);
  //����� ������
  AddMemoMsg('����� ������ PlotStyleName: %s', [AObject.PlotStyleName]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadLWPolyline}

procedure TfmMain.ExtractAcad3dPolyline(const AObjectNo: Integer; const AObject: IAcad3dPolyline);
var
  ACoords: OleVariant;
  ACoordsCount, I: Integer;
begin
  AddMemoMsg('������ %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('��� ������� EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('������������ ������� EntityName: %s', [AObject.EntityName]);

  //����� ��������
  //����
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //����
  AddMemoMsg('���� Layer: %s', [AObject.Layer]);
  //��� �����
  AddMemoMsg('��� ����� LineType: %s', [AObject.Linetype]);
  //������� ���� �����
  AddMemoMsg('������� ���� ����� LinetypeScale: %f', [AObject.LinetypeScale]);
  //��� �����
  AddMemoMsg('��� ����� Lineweight: %d', [AObject.Lineweight]);

  //3D ������������
  //��������
  AddMemoMsg('�������� Material: %s', [AObject.Material]);

  //���������
  //���������� ������
  ACoords := AObject.Coordinates;
  ACoordsCount := (VarArrayHighBound(ACoords, 1) - VarArrayLowBound(ACoords, 1) + 1) div 3;
  AddMemoMsg('ACoordsCount: %d', [ACoordsCount]);
  for I := 0 to ACoordsCount - 1 do
    AddMemoMsg('Coordinates[%d]: (%f; %f; %f)', [I+1, Double(ACoords[3*I]), Double(ACoords[3*I+1]), Double(ACoords[3*I+2])]);

  //��������� ��������
  //�����
  AddMemoMsg('����� Length: %f', [AObject.Length]);

  //������
  //��������
  AddMemoMsg('�������� Closed: %d', [Integer(AObject.Closed)]);
  //���
  AddMemoMsg('��� Type: %d', [AObject.type_]);

  //������ ��������
  //���������
  AddMemoMsg('��������� Visible: %d', [Integer(AObject.Visible)]);
  //����� ������
  AddMemoMsg('����� ������ PlotStyleName: %s', [AObject.PlotStyleName]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcad3dPolyline}

procedure TfmMain.ExtractAcadCircle(const AObjectNo: Integer; const AObject: IAcadCircle);
begin
  AddMemoMsg('������ %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('��� ������� EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('������������ ������� EntityName: %s', [AObject.EntityName]);

  //����� ��������
  //����
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //����
  AddMemoMsg('���� Layer: %s', [AObject.Layer]);
  //��� �����
  AddMemoMsg('��� ����� LineType: %s', [AObject.Linetype]);
  //������� ���� �����
  AddMemoMsg('������� ���� ����� LinetypeScale: %f', [AObject.LinetypeScale]);
  //��� �����
  AddMemoMsg('��� ����� Lineweight: %d', [AObject.Lineweight]);
  //������ 3D
  AddMemoMsg('������3D Thickness: %f', [AObject.Thickness]);

  //3D ������������
  //��������
  AddMemoMsg('�������� Material: %s', [AObject.Material]);

  //���������
  //���������� ������
  AddMemoMsg('���������� ������ Center: (%f, %f, %f)', [Double(AObject.Center[0]), Double(AObject.Center[1]), Double(AObject.Center[2])]);
  //������
  AddMemoMsg('Radius: %f', [AObject.Radius]);
  //�������
  AddMemoMsg('Diameter: %f', [AObject.Diameter]);
  //����� ����������
  AddMemoMsg('����� ���������� Circumference: %f', [AObject.Circumference]);
  //�������
  AddMemoMsg('������� Area: %f', [AObject.Area]);

  //��������� ��������
  //�������
  AddMemoMsg('Normal: (%f; %f; %f)', [Double(AObject.Normal[0]), Double(AObject.Normal[1]), Double(AObject.Normal[2])]);

  //������ ��������
  //���������
  AddMemoMsg('��������� Visible: %d', [Integer(AObject.Visible)]);
  //����� ������
  AddMemoMsg('����� ������ PlotStyleName: %s', [AObject.PlotStyleName]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadCircle}

procedure TfmMain.ExtractAcadArc(const AObjectNo: Integer; const AObject: IAcadArc);
begin
  AddMemoMsg('������ %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('��� ������� EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('������������ ������� EntityName: %s', [AObject.EntityName]);

  //����� ��������
  //����
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //����
  AddMemoMsg('���� Layer: %s', [AObject.Layer]);
  //��� �����
  AddMemoMsg('��� ����� LineType: %s', [AObject.Linetype]);
  //������� ���� �����
  AddMemoMsg('������� ���� ����� LinetypeScale: %f', [AObject.LinetypeScale]);
  //��� �����
  AddMemoMsg('��� ����� Lineweight: %d', [AObject.Lineweight]);
  //������ 3D
  AddMemoMsg('������3D Thickness: %f', [AObject.Thickness]);

  //3D ������������
  //��������
  AddMemoMsg('�������� Material: %s', [AObject.Material]);

  //���������
  //���������� ������
  AddMemoMsg('���������� ������ Center: (%f, %f, %f)', [Double(AObject.Center[0]), Double(AObject.Center[1]), Double(AObject.Center[2])]);
  //������
  AddMemoMsg('������ Radius: %f', [AObject.Radius]);
  //��������� ����
  AddMemoMsg('��������� ���� StartAngle: %f', [AObject.StartAngle]);
  //�������� ����
  AddMemoMsg('�������� ���� EndAngle: %f', [AObject.EndAngle]);

  //��������� ��������
  //��������� �����
  AddMemoMsg('��������� ����� StartPoint: (%f, %f, %f)', [Double(AObject.StartPoint[0]), Double(AObject.StartPoint[1]), Double(AObject.StartPoint[2])]);
  //�������� �����
  AddMemoMsg('�������� ����� EndPoint: (%f, %f, %f)', [Double(AObject.EndPoint[0]), Double(AObject.EndPoint[1]), Double(AObject.EndPoint[2])]);
  //������ ����
  AddMemoMsg('������ ���� TotalAngle: %f', [AObject.TotalAngle]);
  //����� ����
  AddMemoMsg('����� ���� ArcLength: %f', [AObject.ArcLength]);
  //�������
  AddMemoMsg('������� Area: %f', [AObject.Area]);
  //�������
  AddMemoMsg('Normal: (%f; %f; %f)', [Double(AObject.Normal[0]), Double(AObject.Normal[1]), Double(AObject.Normal[2])]);

  //������ ��������
  //���������
  AddMemoMsg('��������� Visible: %d', [Integer(AObject.Visible)]);
  //����� ������
  AddMemoMsg('����� ������ PlotStyleName: %s', [AObject.PlotStyleName]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadArc}

procedure TfmMain.ExtractAcadEllipse(const AObjectNo: Integer; const AObject: IAcadEllipse);
begin
  AddMemoMsg('������ %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('��� ������� EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('������������ ������� EntityName: %s', [AObject.EntityName]);

  //����� ��������
  //����
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //����
  AddMemoMsg('���� Layer: %s', [AObject.Layer]);
  //��� �����
  AddMemoMsg('��� ����� LineType: %s', [AObject.Linetype]);
  //������� ���� �����
  AddMemoMsg('������� ���� ����� LinetypeScale: %f', [AObject.LinetypeScale]);
  //��� �����
  AddMemoMsg('��� ����� Lineweight: %d', [AObject.Lineweight]);

  //3D ������������
  //��������
  AddMemoMsg('�������� Material: %s', [AObject.Material]);

  //���������
  //���������� ������
  AddMemoMsg('���������� ������ Center: (%f, %f, %f)', [Double(AObject.Center[0]), Double(AObject.Center[1]), Double(AObject.Center[2])]);
  //������� �������
  AddMemoMsg('������� ������� MajorRadius: %f', [AObject.MajorRadius]);
  //����� �������
  AddMemoMsg('����� ������� MinorRadius: %f', [AObject.MinorRadius]);
  //��������� �������� - ��������������
  AddMemoMsg('��������� �������� RadiusRatio: %f', [AObject.RadiusRatio]);
  //��������� ����
  AddMemoMsg('��������� ���� StartAngle: %f', [AObject.StartAngle]);
  //�������� ����
  AddMemoMsg('�������� ���� EndAngle: %f', [AObject.EndAngle]);

  //��������� ��������
  //��������� �����
  AddMemoMsg('��������� ����� StartPoint: (%f, %f, %f)', [Double(AObject.StartPoint[0]), Double(AObject.StartPoint[1]), Double(AObject.StartPoint[2])]);
  //�������� �����
  AddMemoMsg('�������� ����� EndPoint: (%f, %f, %f)', [Double(AObject.EndPoint[0]), Double(AObject.EndPoint[1]), Double(AObject.EndPoint[2])]);
  //������ ������� ����
  AddMemoMsg('������ ������� ���� MajorAxis: (%f, %f, %f)', [Double(AObject.MajorAxis[0]), Double(AObject.MajorAxis[1]), Double(AObject.MajorAxis[2])]);
  //������ ����� ����
  AddMemoMsg('������ ����� ���� MinorAxis: (%f, %f, %f)', [Double(AObject.MinorAxis[0]), Double(AObject.MinorAxis[1]), Double(AObject.MinorAxis[2])]);
  //�������
  AddMemoMsg('������� Area: %f', [AObject.Area]);

  //������ ��������
  //�������
  AddMemoMsg('Normal: (%f; %f; %f)', [Double(AObject.Normal[0]), Double(AObject.Normal[1]), Double(AObject.Normal[2])]);
  //���������
  AddMemoMsg('��������� Visible: %d', [Integer(AObject.Visible)]);
  //����� ������
  AddMemoMsg('����� ������ PlotStyleName: %s', [AObject.PlotStyleName]);
  //��������� ��������
  AddMemoMsg('��������� �������� StartParameter: %f', [AObject.StartParameter]);
  //�������� ��������
  AddMemoMsg('�������� �������� EndParameter: %f', [AObject.EndParameter]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadEllipse}

procedure TfmMain.ExtractAcadSpline(const AObjectNo: Integer; const AObject: IAcadSpline);
var
  AControlPoints, AFitPoints: OleVariant;
  AControlPointsCount, AFitPointsCount, I: Integer;
begin
  AddMemoMsg('������ %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('��� ������� EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('������������ ������� EntityName: %s', [AObject.EntityName]);

  //����� ��������
  //����
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //����
  AddMemoMsg('���� Layer: %s', [AObject.Layer]);
  //��� �����
  AddMemoMsg('��� ����� LineType: %s', [AObject.Linetype]);
  //������� ���� �����
  AddMemoMsg('������� ���� ����� LinetypeScale: %f', [AObject.LinetypeScale]);
  //��� �����
  AddMemoMsg('��� ����� Lineweight: %d', [AObject.Lineweight]);

  //3D ������������
  //��������
  AddMemoMsg('�������� Material: %s', [AObject.Material]);

  //������������ �����
  //���������� ����������� �����
  AControlPoints := AObject.ControlPoints;
  AControlPointsCount := AObject.NumberOfControlPoints;
  AddMemoMsg('AControlPointsCount: %d', [AControlPointsCount]);
  for I := 0 to AControlPointsCount - 1 do
    AddMemoMsg('����������� ����� AControlPoints[%d]: (%f; %f; %f)', [I+1, Double(AControlPoints[3*I]), Double(AControlPoints[3*I+1]), Double(AControlPoints[3*I+2])]);
  //���������� ������������ �����
  AFitPoints := AObject.FitPoints;
  AFitPointsCount := AObject.NumberOfFitPoints;
  AddMemoMsg('AFitPointsCount: %d', [AFitPointsCount]);
  for I := 0 to AFitPointsCount - 1 do
    AddMemoMsg('������������ �����  AFitPoints[%d]: (%f; %f; %f)', [I+1, Double(AFitPoints[3*I]), Double(AFitPoints[3*I+1]), Double(AFitPoints[3*I+2])]);
  //������� ��������������
  //AddMemoMsg('������� �������������� KnotParameterization: %d', [AObject.KnotParameterization]);
  //������ ��
  //AddMemoMsg('������ �� SplineFrame: %d', [AObject.SplineFrame]);
  
  //������
  //������
  //AddMemoMsg('������ SplineMethod: %d', [AObject.SplineMethod]);
  //��������
  AddMemoMsg('�������� Closed: %d', [Integer(AObject.Closed)]);
  //if AObject.SplineMethod = 0 then
  //begin
  //������ ����������� � ������ - X,Y,Z
  AddMemoMsg('������ ����������� � ������ StartTangent: (%f; %f; %f)', [Double(AObject.StartTangent[0]), Double(AObject.StartTangent[1]), Double(AObject.StartTangent[2])]);
  //������ ����������� � ����� - X,Y,Z
  AddMemoMsg('������ ����������� � ����� EndTangent: (%f; %f; %f)', [Double(AObject.EndTangent[0]), Double(AObject.EndTangent[1]), Double(AObject.EndTangent[2])]);
  //end;{if}
  //������
  AddMemoMsg('������ FitTolerance: %f', [AObject.FitTolerance]);

  //��������� ��������
  //�������
  AddMemoMsg('������� Degree: %d', [AObject.Degree]);
  //�������������
  AddMemoMsg('������������� IsPeriodic: %d', [Integer(AObject.IsPeriodic)]);
  //���������
  AddMemoMsg('��������� IsPlanar: %d', [Integer(AObject.IsPlanar)]);
  //�������
  AddMemoMsg('������� Area: %f', [AObject.Area]);
  //
  AddMemoMsg('IsRational: %d', [Integer(AObject.IsRational)]);

  //������ ��������
  //���������
  AddMemoMsg('��������� Visible: %d', [Integer(AObject.Visible)]);
  //����� ������
  AddMemoMsg('����� ������ PlotStyleName: %s', [AObject.PlotStyleName]);
  //����
  AddMemoMsg('���� Knots: (%f; %f; %f)', [Double(AObject.Knots[0]), Double(AObject.Knots[1]), Double(AObject.Knots[2])]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadSpline}

procedure TfmMain.ExtractAcadText(const AObjectNo: Integer; const AObject: IAcadText);
begin
  AddMemoMsg('������ %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('��� ������� EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('������������ ������� EntityName: %s', [AObject.EntityName]);

  //����� ��������
  //����
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //����
  AddMemoMsg('���� Layer: %s', [AObject.Layer]);
  //��� �����
  AddMemoMsg('��� ����� LineType: %s', [AObject.Linetype]);
  //������� ���� �����
  AddMemoMsg('������� ���� ����� LinetypeScale: %f', [AObject.LinetypeScale]);
  //��� �����
  AddMemoMsg('��� ����� Lineweight: %d', [AObject.Lineweight]);

  //3D ������������
  //��������
  AddMemoMsg('�������� Material: %s', [AObject.Material]);
  
  //�����
  //����������
  AddMemoMsg('���������� TextString: %s', [AObject.TextString]);
  //�����
  AddMemoMsg('����� StyleName: %s', [AObject.StyleName]);
  //������������
  AddMemoMsg('������������ TextGenerationFlag: %d', [AObject.TextGenerationFlag]);
  //������������
  AddMemoMsg('������������� Alignment: %d', [AObject.Alignment]);
  AddMemoMsg('HorizontalAlignment: %d', [AObject.HorizontalAlignment]);
  AddMemoMsg('VerticalAlignment: %d', [AObject.VerticalAlignment]);
  //������
  AddMemoMsg('������ Height: %f', [AObject.Height]);
  //�������
  AddMemoMsg('������� Rotation: %f', [AObject.Rotation]);
  //����������� ������
  AddMemoMsg('����������� ������ ScaleFactor: %f', [AObject.ScaleFactor]);
  //���� �������
  AddMemoMsg('���� ������� ObliqueAngle: %f', [AObject.ObliqueAngle]);

  //���������
  //���������
  AddMemoMsg('��������� InsertionPoint: (%f, %f, %f)', [Double(AObject.InsertionPoint[0]), Double(AObject.InsertionPoint[1]), Double(AObject.InsertionPoint[2])]);

  //������
  //������������
  AddMemoMsg('������������ UpsideDown: %d', [Integer(AObject.UpsideDown)]);
  //������ ������
  AddMemoMsg('������ ������ Backward: %d', [Integer(AObject.Backward)]);

  //��������� ��������
  //�������� ������ �� ���� X, Y, Z
  AddMemoMsg('�������� ������ TextAlignmentPoint: (%f, %f, %f)', [Double(AObject.TextAlignmentPoint[0]), Double(AObject.TextAlignmentPoint[1]), Double(AObject.TextAlignmentPoint[2])]);

  //������ ��������
  //�������
  AddMemoMsg('Normal: (%f; %f; %f)', [Double(AObject.Normal[0]), Double(AObject.Normal[1]), Double(AObject.Normal[2])]);
  //���������
  AddMemoMsg('��������� Visible: %d', [Integer(AObject.Visible)]);
  //����� ������
  AddMemoMsg('����� ������ PlotStyleName: %s', [AObject.PlotStyleName]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadText}

procedure TfmMain.ExtractAcadMText(const AObjectNo: Integer; const AObject: IAcadMText);
begin
  AddMemoMsg('������ %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('��� ������� EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('������������ ������� EntityName: %s', [AObject.EntityName]);

  //����� ��������
  //����
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //����
  AddMemoMsg('���� Layer: %s', [AObject.Layer]);
  //��� �����
  AddMemoMsg('��� ����� LineType: %s', [AObject.Linetype]);
  //������� ���� �����
  AddMemoMsg('������� ���� ����� LinetypeScale: %f', [AObject.LinetypeScale]);
  //��� �����
  AddMemoMsg('��� ����� Lineweight: %d', [AObject.Lineweight]);

  //3D ������������
  //��������
  AddMemoMsg('�������� Material: %s', [AObject.Material]);

  //�����
  //����������
  AddMemoMsg('���������� TextString: %s', [AObject.TextString]);
  //�����
  AddMemoMsg('����� StyleName: %s', [AObject.StyleName]);
  //������������
  AddMemoMsg('������������ AttachmentPoint: %d', [AObject.AttachmentPoint]);
  //�����������
  AddMemoMsg('����������� DrawingDirection: %d', [AObject.DrawingDirection]);
  //������ ������
  AddMemoMsg('������ Height: %f', [AObject.Height]);
  //�������
  AddMemoMsg('������� Rotation: %f', [AObject.Rotation]);
  //����������� ��������
  AddMemoMsg('����������� �������� LineSpacingDistance: %f', [AObject.LineSpacingDistance]);
  //�������� ����� �������
  AddMemoMsg('�������� ����� ������� LineSpacingFactor: %f', [AObject.LineSpacingFactor]);
  //����� ���.��������� ���������
  AddMemoMsg('����� ���.��������� ��������� LineSpacingStyle: %d', [AObject.LineSpacingStyle]);
  //������� ������� �����
  AddMemoMsg('������� ������� ����� BackgroundFill: %d', [Integer(AObject.BackgroundFill)]);
  //�������
  AddMemoMsg('������� Width: %f', [AObject.Width]);
  
  //���������
  //���������
  AddMemoMsg('��������� InsertionPoint: (%f, %f, %f)', [Double(AObject.InsertionPoint[0]), Double(AObject.InsertionPoint[1]), Double(AObject.InsertionPoint[2])]);

  //������ ��������
  //�������
  AddMemoMsg('Normal: (%f; %f; %f)', [Double(AObject.Normal[0]), Double(AObject.Normal[1]), Double(AObject.Normal[2])]);
  //���������
  AddMemoMsg('��������� Visible: %d', [Integer(AObject.Visible)]);
  //����� ������
  AddMemoMsg('����� ������ PlotStyleName: %s', [AObject.PlotStyleName]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadMText}

//����� � �������������
procedure esaPause(const AMsecs: Cardinal = 1500);
var AStart: Cardinal;
begin
  AStart := GetTickCount();
  while not GetTickCount - AStart < AMsecs do
    Application.ProcessMessages();
end;{esaPause}

procedure TfmMain._ExtractAutoCADFile();
var
  AAutoCAD   : IAcadApplication;//���������� AutoCAD
  ADoc       : IAcadDocument;   //�������� �������� AutoCAD
  ALayers    : IAcadLayers;     //���� AutoCAD
  AObjects   : IAcadModelSpace; //������� AutoCAD
  //ABlocks    : IAcadBlocks;     //����� AutoCAD

  AObj       : IAcadEntity;     //������� ������ AutoCAD
  //ABlock     : IAcadBlock;      //������� ���� AutoCAD
  //ABlockObj  : IAcadEntity;     //������� ������ ����� AutoCAD

  //AUnknown   : IUnknown;        //����������� ��������� AutoCAD
  //AClassID   : TCLSID;          //����������������� ����� AutoCAD
  //AIsCreated : Boolean;         //������� �� ��������� AutoCAD
  //AErrorCode : HResult;         //������, ����������� ��� ��������������� AutoCAD
  
  I: Integer;
  //J: Integer;
begin
  if OpenDialog.Execute then
  begin
    ExtractAcadClear();
    //AutoCAD.Objects ----------------------------------------------------------
    try
      (*
      //���� AutoCAD �� ���������� � ���������� ������
      AErrorCode := CLSIDFromProgID(PWideChar(WideString('AutoCAD.Application')), AClassID); //�������� ������� AutoCAD �� ��
      if Succeeded(AErrorCode) then
        AIsCreated := Succeeded(GetActiveObject(AClassID, nil, AUnknown)) //�������� ������� ����������� AutoCAD
      else
        raise EOleSysError.Create('AutoCAD �� ������!', AErrorCode, 0);
      if not AIsCreated then
        //AAutoCAD := IDispatch(CreateComObject(AClassID));
        AAutoCAD := IDispatch(CreateOleObject('AutoCAD.Application')) as IAcadApplication //������ AutoCAD
      else
        //AAutoCAD := IDispatch(AUnknown);
        AAutoCAD := IDispatch(AUnknown) as IAcadApplication; //������������� � ����������� AutoCAD
      *)

      AAutoCAD := IDispatch(CreateOleObject('AutoCAD.Application')) as IAcadApplication;
      AAutoCAD.Visible := False;
      //acadPause(1500);
      //�������� �������� ------------------------------------------------------
      if AAutoCAD.Documents.Count > 0 then AAutoCAD.Documents.Close();
      ADoc := AAutoCAD.Documents.Open(OpenDialog.FileName, True, Null);

      (*
      esaPause(1500);
      AddMemoMsg('AAutoCAD.Name: %s', [AAutoCAD.Name]);
      AddMemoMsg('AAutoCAD.Caption: %s', [AAutoCAD.Caption]);
      AddMemoMsg('AAutoCAD.FullName: %s', [AAutoCAD.FullName]);
      AddMemoMsg('AAutoCAD.Version: %s', [AAutoCAD.Version]);
      *)

      AddMemoMsg('ADoc.FullName: %s', [ADoc.FullName], True);
      //����
      ALayers := ADoc.Layers;
      for I := 0 to ALayers.Count - 1 do
        ExtractAcadLayer(I+1, ALayers.Item(I) as IAcadLayer);
      //�������
      AObjects := ADoc.ModelSpace;
      AddMemoMsg('ModelSpace.Count: %d',[AObjects.Count], True);

     (*
      //LineTypes --------------------------------------------------------------
      for I := 0 to ADoc.Database.Linetypes.Count-1 do
      begin
        AddMemoMsg('# %d; ADoc.Database.Linetypes.Item(I).Name %s;  ADoc.Database.Linetypes.Item(I).Description %s',
        [I+1, ADoc.Database.Linetypes.Item(I).Name, ADoc.Database.Linetypes.Item(I).Description]);
      end;{for}

      AddMemoMsg('ADoc.ModelSpace.Database.Linetypes.Count: %d', [ADoc.ModelSpace.Database.Linetypes.Count]);
      for I := 0 to ADoc.ModelSpace.Database.Linetypes.Count-1 do
      begin
        AddMemoMsg('# %d; ADoc.ModelSpace.Database.Linetypes.Item(I).Name %s;  ADoc.ModelSpace.Database.Linetypes.Item(I).Description %s',
        [I+1, ADoc.ModelSpace.Database.Linetypes.Item(I).Name, ADoc.ModelSpace.Database.Linetypes.Item(I).Description]);
      end;{for}

      for I := 0 to ADoc.Linetypes.Count-1 do
      begin
        AddMemoMsg('# %d; ADoc.Linetypes.Item(I).Name %s;  ADoc.Linetypes.Item(I).Description %s',
        [I+1, ADoc.Linetypes.Item(I).Name, ADoc.Linetypes.Item(I).Description]);
      end;{for}

      //FontNames --------------------------------------------------------------
      for I := 0 to ADoc.TextStyles.Count-1 do
      begin
        AddMemoMsg('# %d; ADoc.TextStyles.Item(I).Name %s', [I+1, ADoc.TextStyles.Item(I).Name]);
      end;{for}

      for I := 0 to ADoc.TextStyles.Count-1 do
      begin
        AddMemoMsg('# %d; ADoc.TextStyles.Item(I).fontFile %s', [I+1, ADoc.TextStyles.Item(I).fontFile]);
      end;{for}   

      //Blocks -----------------------------------------------------------------
      AddMemoMsg('ADoc.Blocks.Count = %d', [ADoc.Blocks.Count], True);
      for I := 0 to ADoc.Blocks.Count-1 do
      begin
        AddMemoMsg('# %d; ADoc.Blocks.Item.Name = %s; ADoc.Blocks.Item.Units = %d', [I+1, ADoc.Blocks.Item(I).Name, ADoc.Blocks.Item(I).Units]);
        for J := 0 to ADoc.Blocks.Item(I).Count-1 do
        begin
          AddMemoMsg('# %d; ADoc.Blocks.Item(I).Item(J).EntityName = %s; ADoc.Blocks.Item(I).Item(J).EntityType = %d', [J+1, ADoc.Blocks.Item(I).Item(J).EntityName, ADoc.Blocks.Item(I).Item(J).EntityType]);
        end;{for}
      end;{for}           
      *)
      (*
      //Blocks -----------------------------------------------------------------  
      ABlocks := ADoc.Blocks;
      AddMemoMsg('����� ���������� ������ � ������� ���������: %d', [ABlocks.Count], True);
      for I := 0 to ABlocks.Count-1 do
      begin
        ABlock := ABlocks.Item(I);
        AddMemoMsg('��� ����� #%d : %s', [I+1, ABlock.Name]);
        AddMemoMsg('���������� �������� � �����: %d', [ABlock.Count]);
        for J := 0 to ABlock.Count-1 do
        begin
          ABlockObj := ABlock.Item(J);
          case ABlockObj.EntityType of
            acBlockReference: ExtractAcadBlock(J+1, ABlockObj as IAcadBlockReference);
            acLine: ExtractAcadLine(J+1, ABlockObj as IAcadLine);
            acCircle: ExtractAcadCircle(J+1, ABlockObj as IAcadCircle);
            acPolylineLight: ExtractAcadLWPolyline(J+1, ABlockObj as IAcadLWPolyline);
            ac3dPolyline: ExtractAcad3dPolyline(J+1, ABlockObj as IAcad3dPolyline);
            acArc: ExtractAcadArc(J+1, ABlockObj as IAcadArc);
          end;{case}
        end;{for}
      end;{for}    
      *)
      for I := 0 to AObjects.Count-1 do
      begin
        AObj := AObjects.Item(I);
        case AObj.EntityType of
          //acPoint: ExtractAcadPoint(I+1, AObj as IAcadPoint);
          //acBlockReference: ExtractAcadBlock(I+1, AObj as IAcadBlockReference);
          acHatch: ExtractAcadHatch(I+1, AObj as IAcadHatch);
          //acLine: ExtractAcadLine(I+1, AObj as IAcadLine);
          //acCircle: ExtractAcadCircle(I+1, AObj as IAcadCircle);
          //acArc: ExtractAcadArc(I+1, AObj as IAcadArc);
          //acSpline: ExtractAcadSpline(I+1, AObj as IAcadSpline);
          //acEllipse: ExtractAcadEllipse(I+1, AObj as IAcadEllipse);
          //acText: ExtractAcadText(I+1, AObj as IAcadText);
          //acMText: ExtractAcadMText(I+1, AObj as IAcadMText);
          //acPolylineLight: ExtractAcadLWPolyline(I+1, AObj as IAcadLWPolyline);
          //ac3dPolyline: ExtractAcad3dPolyline(I+1, AObj as IAcad3dPolyline);
          else FAutoCAD.Statistics.Add(AObj.EntityType, AObj.EntityName,True);
        end;{case}
      end;{for}

      
    finally
      AAutoCAD.Quit;
    end;{try}
  end;{if}
end;{_ExtractAutoCAD}

procedure TfmMain.btAutoCADClick(Sender: TObject);
begin
  //_ExtractAutoCADFile();
  if OpenDialog.Execute then
  begin
    FAutoCAD.ImportFromAutoCADFile(OpenDialog.FileName);
    FAutoCAD.Draw(PaintBox.Canvas, Rect(0,0,PaintBox.Width,PaintBox.Height));
  end;{if}
  ExtractAcadStatistics();
end;{btAutoCADClick}

procedure TfmMain.FormCreate(Sender: TObject);
begin
  OpenDialog.InitialDir := ExtractFileDir(Application.ExeName);
  stgStatistics.Cells[0, 0] := '�';
  stgStatistics.Cells[1, 0] := 'EntityType';
  stgStatistics.Cells[2, 0] := 'EntityName';
  stgStatistics.Cells[3, 0] := '�����';
  stgStatistics.Cells[4, 0] := '����������';
  
  FAutoCAD := TAutoCAD.Create();
end;{FormCreate}

procedure TfmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  FreeAndNil(FAutoCAD);
end;{FormClose}

procedure TfmMain.PaintBoxPaint(Sender: TObject);
begin
  FAutoCAD.Draw(PaintBox.Canvas, Rect(0,0,PaintBox.Width,PaintBox.Height));
end;{PaintBoxPaint}

end.
