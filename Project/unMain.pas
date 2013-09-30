unit unMain;
  { TODO -oKAA -c2013.06.26 :
  - сайт дла разработчиков Delphi-AutoCAD: http://www.cadhouse.narod.ru/articles/acad/acad_connect.htm
  - полезная ссылка (версии документов AutoCAD) http://forum.dwg.ru/showthread.php?t=12126
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
  protected//Извлечение объектов AutoCAD
    //Извлечение
    procedure _ExtractAutoCADFile();
    //Очистка
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

//Очистка
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

//Обработка cлоев AutoCAD
procedure TfmMain.ExtractAcadLayer(const AObjectNo: Integer; const ALayer: IAcadLayer);
begin
  AddMemoMsg('Слой Layer №%d', [AObjectNo], True);
  AddMemoMsg('Название слоя Name: %s', [ALayer.Name], True);
  //ОБЩИЕ НАСТРОЙКИ
  AddMemoMsg('Включен LayerOn: %d', [Integer(ALayer.LayerOn = True)]);
  AddMemoMsg('Заморозить Freeze: %d', [Integer(ALayer.Freeze = True)]);
  AddMemoMsg('Блокировать Lock: %d', [Integer(ALayer.Lock = True)]);
  //Цвет
  //AddMemoMsg('TrueColor.EntityColor: %d', [ALayer.TrueColor.EntityColor]);
  //AddMemoMsg('TrueColor.ColorName: %s', [ALayer.TrueColor.ColorName]);
  //AddMemoMsg('TrueColor.BookName: %s', [ALayer.TrueColor.BookName]);
  AddMemoMsg('Цвет линии по умолчанию TrueColor.Red: %d', [ALayer.TrueColor.Red]);
  AddMemoMsg('Цвет линии по умолчанию TrueColor.Green: %d', [ALayer.TrueColor.Green]);
  AddMemoMsg('Цвет линии по умолчанию TrueColor.Blue: %d', [ALayer.TrueColor.Blue]);
  //AddMemoMsg('TrueColor.ColorMethod: %d', [ALayer.TrueColor.ColorMethod]);
  //AddMemoMsg('TrueColor.ColorIndex: %d', [ALayer.TrueColor.ColorIndex]);
  AddMemoMsg('Тип линии по умолчанию Linetype: %s', [ALayer.Linetype]);
  AddMemoMsg('Вес линии по умолчанию Lineweight: $%x', [ALayer.Lineweight]);
  //AddMemoMsg('Стиль печати PlotStyleName: %s', [ALayer.PlotStyleName]);
  //AddMemoMsg('Печать Plottable: %d', [Integer(ALayer.Plottable)]);
  AddMemoMsg('Пояснение Description: %s', [ALayer.Description]);
  //ПРОЧИЕ НАСТРОЙКИ
  //AddMemoMsg('ViewportDefault: %d', [Integer(ALayer.ViewportDefault)]);
  //AddMemoMsg('Использован Used: %d', [Integer(ALayer.Used)]);
  //AddMemoMsg('Материал Material: %s', [ALayer.Material]);
end;{ExtractAcadLayer}
procedure TfmMain.ExtractAcadPoint(const AObjectNo: Integer; const AObject: IAcadPoint);
begin
  AddMemoMsg('Слой Layer: %s', [AObject.Layer]);
  AddMemoMsg('Объект №%d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('Тип объекта EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('Наименование объекта EntityName: %s', [AObject.EntityName]);
  AddMemoMsg('Видимость Visible: %d', [Integer(AObject.Visible)]);
  //
  AddMemoMsg('Координаты Coords: (%f, %f, %f)', [Double(AObject.Coordinates[0]), Double(AObject.Coordinates[1]), Double(AObject.Coordinates[2])]);
  //AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  //AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  //AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  //AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  //AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //AddMemoMsg('Размер Lineweight: %d', [AObject.Lineweight]);
  AddMemoMsg('Размер Thickness: %f', [AObject.Thickness]);

  //AddMemoMsg('Тип линий LineType: %s', [AObject.Linetype]);
  //AddMemoMsg('Масштаб типа линий LinetypeScale: %f', [AObject.LinetypeScale]);
  //AddMemoMsg('Материал Material: %s', [AObject.Material]);
  //AddMemoMsg('Normal: (%f; %f; %f)', [Double(AObject.Normal[0]), Double(AObject.Normal[1]), Double(AObject.Normal[2])]);
  //AddMemoMsg('Стиль печати PlotStyleName: %s', [AObject.PlotStyleName]);
  //FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadPoint}

//Обработка блоков AutoCAD
procedure TfmMain.ExtractAcadBlock(const AObjectNo: Integer; const ABlock: IAcadBlockReference);
begin
  AddMemoMsg('Блок Block %d "%s"', [AObjectNo, ABlock.Name], True);
  AddMemoMsg('Тип объекта EntityType: %d', [ABlock.EntityType]);
  AddMemoMsg('Наименование объекта EntityName: %s', [ABlock.EntityName]);
  
  //ОБЩИЕ НАСТРОЙКИ
  //Цвет
  AddMemoMsg('TrueColor.EntityColor: %d', [ABlock.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [ABlock.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [ABlock.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [ABlock.TrueColor.Red]);
  AddMemoMsg('TrueColor.Green: %d', [ABlock.TrueColor.Green]);
  AddMemoMsg('TrueColor.Blue: %d', [ABlock.TrueColor.Blue]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [ABlock.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [ABlock.TrueColor.ColorIndex]);
  //Слой
  AddMemoMsg('Слой Layer: %s', [ABlock.Layer]);
  //Тип линий
  AddMemoMsg('Тип линии Linetype: %s', [ABlock.Linetype]);
  //Масштаб типа линий
  AddMemoMsg('Масштаб типа линий LinetypeScale: %f', [ABlock.LinetypeScale]);
  //Вес линий
  AddMemoMsg('Вес линии Lineweight: %d', [ABlock.Lineweight]);

  //3D ВИЗУАЛИЗАЦИЯ
  //Материал
  AddMemoMsg('Материал Material: %s', [ABlock.Material]);

  //ГЕОМЕТРИЯ
  //Положение
  AddMemoMsg('Положение InsertionPoint: (%f, %f, %f)', [Double(ABlock.InsertionPoint[0]), Double(ABlock.InsertionPoint[1]), Double(ABlock.InsertionPoint[2])]);
  //Масштаб
  AddMemoMsg('Масштаб ScaleFactor: (%f, %f, %f)', [ABlock.XScaleFactor, ABlock.YScaleFactor, ABlock.ZScaleFactor]);

  //РАЗНОЕ
  //Имя
  AddMemoMsg('Имя EffectiveName: %s', [ABlock.EffectiveName]);
  //Поворот
  AddMemoMsg('Поворот Rotation: %f', [ABlock.Rotation]);
  //Аннотативный
  AddMemoMsg('Аннотативный IsDynamicBlock: %d', [Integer(ABlock.IsDynamicBlock)]);
  //Единицы блока
  AddMemoMsg('Единицы блока InsUnits: %s', [ABlock.InsUnits]);
  //Коэффициент единиц
  AddMemoMsg('Коэффициент единиц InsUnitsFactor: %f', [ABlock.InsUnitsFactor]);

  //ПРОЧИЕ СВОЙСТВА
  //Нормаль
  AddMemoMsg('Normal: (%f; %f; %f)', [Double(ABlock.Normal[0]), Double(ABlock.Normal[1]), Double(ABlock.Normal[2])]);
  //Видимость
  AddMemoMsg('Видимость Visible: %d', [Integer(ABlock.Visible)]);
  //Стиль печати
  AddMemoMsg('Стиль печати PlotStyleName: %s', [ABlock.PlotStyleName]);
  //Атрибуты
  AddMemoMsg('Атрибуты HasAttributes: %d', [Integer(ABlock.HasAttributes)]);


  //
  AddMemoMsg('Имя объекта ObjectName: %s', [ABlock.ObjectName]);
  AddMemoMsg('ID объекта ObjectID: %d', [ABlock.ObjectID]);

  FAutoCAD.Statistics.Add(ABlock.EntityType, ABlock.EntityName);
end;{ExtractAcadBlock}

//Обработка объектов AutoCAD
procedure TfmMain.ExtractAcadHatch(const AObjectNo: Integer; const AObject: IAcadHatch);
begin
  AddMemoMsg('Объект %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('Тип объекта EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('Наименование объекта EntityName: %s', [AObject.EntityName]);
  //ОБЩИЕ СВОЙСТВА
  //Цвет
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //Слой
  AddMemoMsg('Слой Layer: %s', [AObject.Layer]);
  //Тип линий
  AddMemoMsg('Тип линий LineType: %s', [AObject.Linetype]);
  //Масштаб типа линий
  AddMemoMsg('Масштаб типа линий LinetypeScale: %f', [AObject.LinetypeScale]);
  //Вес линий
  AddMemoMsg('Вес линий Lineweight: %d', [AObject.Lineweight]);

  //ОБРАЗЕЦ
  //Тип заливки
  AddMemoMsg('Тип заливки PatternType: %d', [AObject.PatternType]);
  //Имя заливки
  AddMemoMsg('Тип заливки PatternName: %s', [AObject.PatternName]);
  //Аннотативный

  //Угол
  AddMemoMsg('Угол PatternAngle: %f', [AObject.PatternAngle]);
  //Масштаб
  AddMemoMsg('Масштаб PatternScale: %f', [AObject.PatternScale]);
  //Исходная точка
  AddMemoMsg('Исходная точка Origin: (%f, %f)', [Double(AObject.Origin[0]), Double(AObject.Origin[1])]);
  //Ассоциативный

  //Стиль решения островков
  AddMemoMsg('Стиль решения островков HatchStyle: %d', [AObject.HatchStyle]);
  //ГЕОМЕТРИЯ
  //Уровень
  AddMemoMsg('Уровень Elevation: %f', [AObject.Elevation]);

  //РАСЧЕТНЫЕ СВОЙСТВА
  //Площадь
  AddMemoMsg('Площадь Area: %f', [AObject.Area]);
  //Толщина пера по ISO
  AddMemoMsg('Толщина пера по ISO ISOPenWidth: %d', [AObject.ISOPenWidth]);

  //ПРОЧИЕ СВОЙСТВА

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

  //Тип объекта заливки
  AddMemoMsg('Тип объекта заливки HatchObjectType: %d', [AObject.HatchObjectType]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadHatch}

procedure TfmMain.ExtractAcadLine(const AObjectNo: Integer; const AObject: IAcadLine);
begin
  AddMemoMsg('Объект %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('Тип объекта EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('Наименование объекта EntityName: %s', [AObject.EntityName]);

  //ОБЩИЕ СВОЙСТВА
  //Цвет
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //Слой
  AddMemoMsg('Слой Layer: %s', [AObject.Layer]);
  //Тип линий
  AddMemoMsg('Тип линий LineType: %s', [AObject.Linetype]);
  //Масштаб типа линий
  AddMemoMsg('Масштаб типа линий LinetypeScale: %f', [AObject.LinetypeScale]);
  //Вес линий
  AddMemoMsg('Вес линий Lineweight: %d', [AObject.Lineweight]);
  //Высота 3D
  AddMemoMsg('Высота3D Thickness: %f', [AObject.Thickness]);
        
  //3D ВИЗУАЛИЗАЦИЯ
  //Материал
  AddMemoMsg('Материал Material: %s', [AObject.Material]);

  //ГЕОМЕТРИЯ
  //Начальная точка
  AddMemoMsg('Начальная точка StartPoint: (%f, %f, %f)', [Double(AObject.StartPoint[0]), Double(AObject.StartPoint[1]), Double(AObject.StartPoint[2])]);
  //Конечная точка
  AddMemoMsg('Конечная точка EndPoint: (%f, %f, %f)', [Double(AObject.EndPoint[0]), Double(AObject.EndPoint[1]), Double(AObject.EndPoint[2])]);

  //РАСЧЕТНЫЕ СВОЙСТВА
  //Дельта
  AddMemoMsg('Дельта Delta: (%f, %f, %f)', [Double(AObject.Delta[0]), Double(AObject.Delta[1]), Double(AObject.Delta[2])]);
  //Длина
  AddMemoMsg('Длина Length: %f', [AObject.Length]);
  //Угол
  AddMemoMsg('Угол Angle: %f', [AObject.Angle]);

  //ПРОЧИЕ СВОЙСТВА
  //Нормаль
  AddMemoMsg('Normal: (%f; %f; %f)', [Double(AObject.Normal[0]), Double(AObject.Normal[1]), Double(AObject.Normal[2])]);
  //Видимость
  AddMemoMsg('Видимость Visible: %d', [Integer(AObject.Visible)]);
  //Стиль печати
  AddMemoMsg('Стиль печати PlotStyleName: %s', [AObject.PlotStyleName]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadLine}


procedure TfmMain.ExtractAcadLWPolyline(const AObjectNo: Integer; const AObject: IAcadLWPolyline);
var
  ACoords: OleVariant;
  ACoordsCount, I: Integer;
begin
  AddMemoMsg('Объект %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('Тип объекта EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('Наименование объекта EntityName: %s', [AObject.EntityName]);

  //ОБЩИЕ СВОЙСТВА
  //Цвет
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);

  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //Слой
  AddMemoMsg('Слой Layer: %s', [AObject.Layer]);
  //Тип линий
  AddMemoMsg('Тип линий LineType: %s', [AObject.Linetype]);
  //Масштаб типа линий
  AddMemoMsg('Масштаб типа линий LinetypeScale: %f', [AObject.LinetypeScale]);
  //Вес линий
  AddMemoMsg('Вес линий Lineweight: %d', [AObject.Lineweight]);
  //Высота 3D
  AddMemoMsg('Высота3D Thickness: %f', [AObject.Thickness]);

  //3D ВИЗУАЛИЗАЦИЯ
  //Материал
  AddMemoMsg('Материал Material: %s', [AObject.Material]);

  //ГЕОМЕТРИЯ
  //Координаты вершин
  ACoords := AObject.Coordinates;
  ACoordsCount := (VarArrayHighBound(ACoords, 1) - VarArrayLowBound(ACoords, 1) + 1) div 2;
  AddMemoMsg('ACoordsCount: %d', [ACoordsCount]);
  for I := 0 to ACoordsCount - 1 do
    AddMemoMsg('Coordinates[%d]: (%f; %f)', [I+1, Double(ACoords[2*I]), Double(ACoords[2*I+1])]);
  //Глобальная ширина
  AddMemoMsg('Глобальная ширина ConstantWidth: %f', [AObject.ConstantWidth]);
  //Уровень
  AddMemoMsg('Уровень Elevation: %f', [AObject.Elevation]);

  //РАСЧЕТНЫЕ СВОЙСТВА
  //Площадь
  AddMemoMsg('Площадь Area: %f', [AObject.Area]);
  //Длина
  AddMemoMsg('Длина Length: %f', [AObject.Length]);

  //РАЗНОЕ
  //Замкнуто
  AddMemoMsg('Замкнуто Closed: %d', [Integer(AObject.Closed)]);
  //Генерация типа линий
  AddMemoMsg('Генерация типа линий LinetypeGeneration: %d', [Integer(AObject.LinetypeGeneration)]);

  //ПРОЧИЕ СВОЙСТВА
  //Нормаль
  AddMemoMsg('Normal: (%f; %f; %f)', [Double(AObject.Normal[0]), Double(AObject.Normal[1]), Double(AObject.Normal[2])]);
  //Видимость
  AddMemoMsg('Видимость Visible: %d', [Integer(AObject.Visible)]);
  //Стиль печати
  AddMemoMsg('Стиль печати PlotStyleName: %s', [AObject.PlotStyleName]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadLWPolyline}

procedure TfmMain.ExtractAcad3dPolyline(const AObjectNo: Integer; const AObject: IAcad3dPolyline);
var
  ACoords: OleVariant;
  ACoordsCount, I: Integer;
begin
  AddMemoMsg('Объект %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('Тип объекта EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('Наименование объекта EntityName: %s', [AObject.EntityName]);

  //ОБЩИЕ СВОЙСТВА
  //Цвет
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //Слой
  AddMemoMsg('Слой Layer: %s', [AObject.Layer]);
  //Тип линий
  AddMemoMsg('Тип линий LineType: %s', [AObject.Linetype]);
  //Масштаб типа линий
  AddMemoMsg('Масштаб типа линий LinetypeScale: %f', [AObject.LinetypeScale]);
  //Вес линий
  AddMemoMsg('Вес линий Lineweight: %d', [AObject.Lineweight]);

  //3D ВИЗУАЛИЗАЦИЯ
  //Материал
  AddMemoMsg('Материал Material: %s', [AObject.Material]);

  //ГЕОМЕТРИЯ
  //Координаты вершин
  ACoords := AObject.Coordinates;
  ACoordsCount := (VarArrayHighBound(ACoords, 1) - VarArrayLowBound(ACoords, 1) + 1) div 3;
  AddMemoMsg('ACoordsCount: %d', [ACoordsCount]);
  for I := 0 to ACoordsCount - 1 do
    AddMemoMsg('Coordinates[%d]: (%f; %f; %f)', [I+1, Double(ACoords[3*I]), Double(ACoords[3*I+1]), Double(ACoords[3*I+2])]);

  //РАСЧЕТНЫЕ СВОЙСТВА
  //Длина
  AddMemoMsg('Длина Length: %f', [AObject.Length]);

  //РАЗНОЕ
  //Замкнуто
  AddMemoMsg('Замкнуто Closed: %d', [Integer(AObject.Closed)]);
  //Тип
  AddMemoMsg('Тип Type: %d', [AObject.type_]);

  //ПРОЧИЕ СВОЙСТВА
  //Видимость
  AddMemoMsg('Видимость Visible: %d', [Integer(AObject.Visible)]);
  //Стиль печати
  AddMemoMsg('Стиль печати PlotStyleName: %s', [AObject.PlotStyleName]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcad3dPolyline}

procedure TfmMain.ExtractAcadCircle(const AObjectNo: Integer; const AObject: IAcadCircle);
begin
  AddMemoMsg('Объект %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('Тип объекта EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('Наименование объекта EntityName: %s', [AObject.EntityName]);

  //ОБЩИЕ СВОЙСТВА
  //Цвет
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //Слой
  AddMemoMsg('Слой Layer: %s', [AObject.Layer]);
  //Тип линий
  AddMemoMsg('Тип линий LineType: %s', [AObject.Linetype]);
  //Масштаб типа линий
  AddMemoMsg('Масштаб типа линий LinetypeScale: %f', [AObject.LinetypeScale]);
  //Вес линий
  AddMemoMsg('Вес линий Lineweight: %d', [AObject.Lineweight]);
  //Высота 3D
  AddMemoMsg('Высота3D Thickness: %f', [AObject.Thickness]);

  //3D ВИЗУАЛИЗАЦИЯ
  //Материал
  AddMemoMsg('Материал Material: %s', [AObject.Material]);

  //ГЕОМЕТРИЯ
  //Координаты центра
  AddMemoMsg('Координаты центра Center: (%f, %f, %f)', [Double(AObject.Center[0]), Double(AObject.Center[1]), Double(AObject.Center[2])]);
  //Радиус
  AddMemoMsg('Radius: %f', [AObject.Radius]);
  //Диаметр
  AddMemoMsg('Diameter: %f', [AObject.Diameter]);
  //Длина окружности
  AddMemoMsg('Длина окружности Circumference: %f', [AObject.Circumference]);
  //Площадь
  AddMemoMsg('Площадь Area: %f', [AObject.Area]);

  //РАСЧЕТНЫЕ СВОЙСТВА
  //Нормаль
  AddMemoMsg('Normal: (%f; %f; %f)', [Double(AObject.Normal[0]), Double(AObject.Normal[1]), Double(AObject.Normal[2])]);

  //ПРОЧИЕ СВОЙСТВА
  //Видимость
  AddMemoMsg('Видимость Visible: %d', [Integer(AObject.Visible)]);
  //Стиль печати
  AddMemoMsg('Стиль печати PlotStyleName: %s', [AObject.PlotStyleName]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadCircle}

procedure TfmMain.ExtractAcadArc(const AObjectNo: Integer; const AObject: IAcadArc);
begin
  AddMemoMsg('Объект %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('Тип объекта EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('Наименование объекта EntityName: %s', [AObject.EntityName]);

  //ОБЩИЕ СВОЙСТВА
  //Цвет
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //Слой
  AddMemoMsg('Слой Layer: %s', [AObject.Layer]);
  //Тип линий
  AddMemoMsg('Тип линий LineType: %s', [AObject.Linetype]);
  //Масштаб типа линий
  AddMemoMsg('Масштаб типа линий LinetypeScale: %f', [AObject.LinetypeScale]);
  //Вес линий
  AddMemoMsg('Вес линий Lineweight: %d', [AObject.Lineweight]);
  //Высота 3D
  AddMemoMsg('Высота3D Thickness: %f', [AObject.Thickness]);

  //3D ВИЗУАЛИЗАЦИЯ
  //Материал
  AddMemoMsg('Материал Material: %s', [AObject.Material]);

  //ГЕОМЕТРИЯ
  //Координаты центра
  AddMemoMsg('Координаты центра Center: (%f, %f, %f)', [Double(AObject.Center[0]), Double(AObject.Center[1]), Double(AObject.Center[2])]);
  //Радиус
  AddMemoMsg('Радиус Radius: %f', [AObject.Radius]);
  //Начальный угол
  AddMemoMsg('Начальный угол StartAngle: %f', [AObject.StartAngle]);
  //Конечный угол
  AddMemoMsg('Конечный угол EndAngle: %f', [AObject.EndAngle]);

  //РАСЧЕТНЫЕ СВОЙСТВА
  //Начальная точка
  AddMemoMsg('Начальная точка StartPoint: (%f, %f, %f)', [Double(AObject.StartPoint[0]), Double(AObject.StartPoint[1]), Double(AObject.StartPoint[2])]);
  //Конечная точка
  AddMemoMsg('Конечная точка EndPoint: (%f, %f, %f)', [Double(AObject.EndPoint[0]), Double(AObject.EndPoint[1]), Double(AObject.EndPoint[2])]);
  //Полный угол
  AddMemoMsg('Полный угол TotalAngle: %f', [AObject.TotalAngle]);
  //Длина дуги
  AddMemoMsg('Длина дуги ArcLength: %f', [AObject.ArcLength]);
  //Площадь
  AddMemoMsg('Площадь Area: %f', [AObject.Area]);
  //Нормаль
  AddMemoMsg('Normal: (%f; %f; %f)', [Double(AObject.Normal[0]), Double(AObject.Normal[1]), Double(AObject.Normal[2])]);

  //ПРОЧИЕ СВОЙСТВА
  //Видимость
  AddMemoMsg('Видимость Visible: %d', [Integer(AObject.Visible)]);
  //Стиль печати
  AddMemoMsg('Стиль печати PlotStyleName: %s', [AObject.PlotStyleName]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadArc}

procedure TfmMain.ExtractAcadEllipse(const AObjectNo: Integer; const AObject: IAcadEllipse);
begin
  AddMemoMsg('Объект %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('Тип объекта EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('Наименование объекта EntityName: %s', [AObject.EntityName]);

  //ОБЩИЕ СВОЙСТВА
  //Цвет
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //Слой
  AddMemoMsg('Слой Layer: %s', [AObject.Layer]);
  //Тип линий
  AddMemoMsg('Тип линий LineType: %s', [AObject.Linetype]);
  //Масштаб типа линий
  AddMemoMsg('Масштаб типа линий LinetypeScale: %f', [AObject.LinetypeScale]);
  //Вес линий
  AddMemoMsg('Вес линий Lineweight: %d', [AObject.Lineweight]);

  //3D ВИЗУАЛИЗАЦИЯ
  //Материал
  AddMemoMsg('Материал Material: %s', [AObject.Material]);

  //ГЕОМЕТРИЯ
  //Координаты центра
  AddMemoMsg('Координаты центра Center: (%f, %f, %f)', [Double(AObject.Center[0]), Double(AObject.Center[1]), Double(AObject.Center[2])]);
  //Большая полуось
  AddMemoMsg('Большая полуось MajorRadius: %f', [AObject.MajorRadius]);
  //Малая полуось
  AddMemoMsg('Малая полуось MinorRadius: %f', [AObject.MinorRadius]);
  //Отношение полуосей - эксцентриситет
  AddMemoMsg('Отношение полуосей RadiusRatio: %f', [AObject.RadiusRatio]);
  //Начальный угол
  AddMemoMsg('Начальный угол StartAngle: %f', [AObject.StartAngle]);
  //Конечный угол
  AddMemoMsg('Конечный угол EndAngle: %f', [AObject.EndAngle]);

  //РАСЧЕТНЫЕ СВОЙСТВА
  //Начальная точка
  AddMemoMsg('Начальная точка StartPoint: (%f, %f, %f)', [Double(AObject.StartPoint[0]), Double(AObject.StartPoint[1]), Double(AObject.StartPoint[2])]);
  //Конечная точка
  AddMemoMsg('Конечная точка EndPoint: (%f, %f, %f)', [Double(AObject.EndPoint[0]), Double(AObject.EndPoint[1]), Double(AObject.EndPoint[2])]);
  //Вектор больших осей
  AddMemoMsg('Вектор больших осей MajorAxis: (%f, %f, %f)', [Double(AObject.MajorAxis[0]), Double(AObject.MajorAxis[1]), Double(AObject.MajorAxis[2])]);
  //Вектор малых осей
  AddMemoMsg('Вектор малых осей MinorAxis: (%f, %f, %f)', [Double(AObject.MinorAxis[0]), Double(AObject.MinorAxis[1]), Double(AObject.MinorAxis[2])]);
  //Площадь
  AddMemoMsg('Площадь Area: %f', [AObject.Area]);

  //ПРОЧИЕ СВОЙСТВА
  //Нормаль
  AddMemoMsg('Normal: (%f; %f; %f)', [Double(AObject.Normal[0]), Double(AObject.Normal[1]), Double(AObject.Normal[2])]);
  //Видимость
  AddMemoMsg('Видимость Visible: %d', [Integer(AObject.Visible)]);
  //Стиль печати
  AddMemoMsg('Стиль печати PlotStyleName: %s', [AObject.PlotStyleName]);
  //Начальный параметр
  AddMemoMsg('Начальный параметр StartParameter: %f', [AObject.StartParameter]);
  //Конечный параметр
  AddMemoMsg('Конечный параметр EndParameter: %f', [AObject.EndParameter]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadEllipse}

procedure TfmMain.ExtractAcadSpline(const AObjectNo: Integer; const AObject: IAcadSpline);
var
  AControlPoints, AFitPoints: OleVariant;
  AControlPointsCount, AFitPointsCount, I: Integer;
begin
  AddMemoMsg('Объект %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('Тип объекта EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('Наименование объекта EntityName: %s', [AObject.EntityName]);

  //ОБЩИЕ СВОЙСТВА
  //Цвет
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //Слой
  AddMemoMsg('Слой Layer: %s', [AObject.Layer]);
  //Тип линий
  AddMemoMsg('Тип линий LineType: %s', [AObject.Linetype]);
  //Масштаб типа линий
  AddMemoMsg('Масштаб типа линий LinetypeScale: %f', [AObject.LinetypeScale]);
  //Вес линий
  AddMemoMsg('Вес линий Lineweight: %d', [AObject.Lineweight]);

  //3D ВИЗУАЛИЗАЦИЯ
  //Материал
  AddMemoMsg('Материал Material: %s', [AObject.Material]);

  //ОПРЕДЕЛЯЮЩИЕ ТОЧКИ
  //Координаты управляющих точек
  AControlPoints := AObject.ControlPoints;
  AControlPointsCount := AObject.NumberOfControlPoints;
  AddMemoMsg('AControlPointsCount: %d', [AControlPointsCount]);
  for I := 0 to AControlPointsCount - 1 do
    AddMemoMsg('Управляющие точки AControlPoints[%d]: (%f; %f; %f)', [I+1, Double(AControlPoints[3*I]), Double(AControlPoints[3*I+1]), Double(AControlPoints[3*I+2])]);
  //Координаты определяющих точек
  AFitPoints := AObject.FitPoints;
  AFitPointsCount := AObject.NumberOfFitPoints;
  AddMemoMsg('AFitPointsCount: %d', [AFitPointsCount]);
  for I := 0 to AFitPointsCount - 1 do
    AddMemoMsg('Определяющие точки  AFitPoints[%d]: (%f; %f; %f)', [I+1, Double(AFitPoints[3*I]), Double(AFitPoints[3*I+1]), Double(AFitPoints[3*I+2])]);
  //Узловая параметризация
  //AddMemoMsg('Узловая параметризация KnotParameterization: %d', [AObject.KnotParameterization]);
  //Каркас УВ
  //AddMemoMsg('Каркас УВ SplineFrame: %d', [AObject.SplineFrame]);
  
  //РАЗНОЕ
  //Способ
  //AddMemoMsg('Способ SplineMethod: %d', [AObject.SplineMethod]);
  //Замкнуто
  AddMemoMsg('Замкнуто Closed: %d', [Integer(AObject.Closed)]);
  //if AObject.SplineMethod = 0 then
  //begin
  //Вектор касательной в начале - X,Y,Z
  AddMemoMsg('Вектор касательной в начале StartTangent: (%f; %f; %f)', [Double(AObject.StartTangent[0]), Double(AObject.StartTangent[1]), Double(AObject.StartTangent[2])]);
  //Вектор касательной в конце - X,Y,Z
  AddMemoMsg('Вектор касательной в конце EndTangent: (%f; %f; %f)', [Double(AObject.EndTangent[0]), Double(AObject.EndTangent[1]), Double(AObject.EndTangent[2])]);
  //end;{if}
  //Допуск
  AddMemoMsg('Допуск FitTolerance: %f', [AObject.FitTolerance]);

  //РАСЧЕТНЫЕ СВОЙСТВА
  //Порядок
  AddMemoMsg('Порядок Degree: %d', [AObject.Degree]);
  //Периодический
  AddMemoMsg('Периодический IsPeriodic: %d', [Integer(AObject.IsPeriodic)]);
  //Плоскость
  AddMemoMsg('Плоскость IsPlanar: %d', [Integer(AObject.IsPlanar)]);
  //Площадь
  AddMemoMsg('Площадь Area: %f', [AObject.Area]);
  //
  AddMemoMsg('IsRational: %d', [Integer(AObject.IsRational)]);

  //ПРОЧИЕ СВОЙСТВА
  //Видимость
  AddMemoMsg('Видимость Visible: %d', [Integer(AObject.Visible)]);
  //Стиль печати
  AddMemoMsg('Стиль печати PlotStyleName: %s', [AObject.PlotStyleName]);
  //Узлы
  AddMemoMsg('Узлы Knots: (%f; %f; %f)', [Double(AObject.Knots[0]), Double(AObject.Knots[1]), Double(AObject.Knots[2])]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadSpline}

procedure TfmMain.ExtractAcadText(const AObjectNo: Integer; const AObject: IAcadText);
begin
  AddMemoMsg('Объект %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('Тип объекта EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('Наименование объекта EntityName: %s', [AObject.EntityName]);

  //ОБЩИЕ СВОЙСТВА
  //Цвет
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //Слой
  AddMemoMsg('Слой Layer: %s', [AObject.Layer]);
  //Тип линий
  AddMemoMsg('Тип линий LineType: %s', [AObject.Linetype]);
  //Масштаб типа линий
  AddMemoMsg('Масштаб типа линий LinetypeScale: %f', [AObject.LinetypeScale]);
  //Вес линий
  AddMemoMsg('Вес линий Lineweight: %d', [AObject.Lineweight]);

  //3D ВИЗУАЛИЗАЦИЯ
  //Материал
  AddMemoMsg('Материал Material: %s', [AObject.Material]);
  
  //ТЕКСТ
  //Содержимое
  AddMemoMsg('Содержимое TextString: %s', [AObject.TextString]);
  //Стиль
  AddMemoMsg('Стиль StyleName: %s', [AObject.StyleName]);
  //Аннотативный
  AddMemoMsg('Аннотативный TextGenerationFlag: %d', [AObject.TextGenerationFlag]);
  //Выравнивание
  AddMemoMsg('Выравниевание Alignment: %d', [AObject.Alignment]);
  AddMemoMsg('HorizontalAlignment: %d', [AObject.HorizontalAlignment]);
  AddMemoMsg('VerticalAlignment: %d', [AObject.VerticalAlignment]);
  //Высота
  AddMemoMsg('Высота Height: %f', [AObject.Height]);
  //Поворот
  AddMemoMsg('Поворот Rotation: %f', [AObject.Rotation]);
  //Коэффициент сжатия
  AddMemoMsg('Коэффициент сжатия ScaleFactor: %f', [AObject.ScaleFactor]);
  //Угол наклона
  AddMemoMsg('Угол наклона ObliqueAngle: %f', [AObject.ObliqueAngle]);

  //ГЕОМЕТРИЯ
  //Положение
  AddMemoMsg('Положение InsertionPoint: (%f, %f, %f)', [Double(AObject.InsertionPoint[0]), Double(AObject.InsertionPoint[1]), Double(AObject.InsertionPoint[2])]);

  //РАЗНОЕ
  //Перевернутый
  AddMemoMsg('Перевернутый UpsideDown: %d', [Integer(AObject.UpsideDown)]);
  //Справа налево
  AddMemoMsg('Справа налево Backward: %d', [Integer(AObject.Backward)]);

  //РАСЧЕТНЫЕ СВОЙСТВА
  //Привязка текста по осям X, Y, Z
  AddMemoMsg('Привязка текста TextAlignmentPoint: (%f, %f, %f)', [Double(AObject.TextAlignmentPoint[0]), Double(AObject.TextAlignmentPoint[1]), Double(AObject.TextAlignmentPoint[2])]);

  //ПРОЧИЕ СВОЙСТВА
  //Нормаль
  AddMemoMsg('Normal: (%f; %f; %f)', [Double(AObject.Normal[0]), Double(AObject.Normal[1]), Double(AObject.Normal[2])]);
  //Видимость
  AddMemoMsg('Видимость Visible: %d', [Integer(AObject.Visible)]);
  //Стиль печати
  AddMemoMsg('Стиль печати PlotStyleName: %s', [AObject.PlotStyleName]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadText}

procedure TfmMain.ExtractAcadMText(const AObjectNo: Integer; const AObject: IAcadMText);
begin
  AddMemoMsg('Объект %d - %s', [AObjectNo, AObject.EntityName], True);
  AddMemoMsg('Тип объекта EntityType: %d', [AObject.EntityType]);
  AddMemoMsg('Наименование объекта EntityName: %s', [AObject.EntityName]);

  //ОБЩИЕ СВОЙСТВА
  //Цвет
  AddMemoMsg('TrueColor.EntityColor: %d', [AObject.TrueColor.EntityColor]);
  AddMemoMsg('TrueColor.ColorName: %s', [AObject.TrueColor.ColorName]);
  AddMemoMsg('TrueColor.BookName: %s', [AObject.TrueColor.BookName]);
  AddMemoMsg('TrueColor.Red: %d', [AObject.TrueColor.Red]);
  AddMemoMsg('TrueColor.Blue: %d', [AObject.TrueColor.Blue]);
  AddMemoMsg('TrueColor.Green: %d', [AObject.TrueColor.Green]);
  AddMemoMsg('TrueColor.ColorMethod: %d', [AObject.TrueColor.ColorMethod]);
  AddMemoMsg('TrueColor.ColorIndex: %d', [AObject.TrueColor.ColorIndex]);
  //Слой
  AddMemoMsg('Слой Layer: %s', [AObject.Layer]);
  //Тип линий
  AddMemoMsg('Тип линий LineType: %s', [AObject.Linetype]);
  //Масштаб типа линий
  AddMemoMsg('Масштаб типа линий LinetypeScale: %f', [AObject.LinetypeScale]);
  //Вес линий
  AddMemoMsg('Вес линий Lineweight: %d', [AObject.Lineweight]);

  //3D ВИЗУАЛИЗАЦИЯ
  //Материал
  AddMemoMsg('Материал Material: %s', [AObject.Material]);

  //ТЕКСТ
  //Содержимое
  AddMemoMsg('Содержимое TextString: %s', [AObject.TextString]);
  //Стиль
  AddMemoMsg('Стиль StyleName: %s', [AObject.StyleName]);
  //Аннотативный
  AddMemoMsg('Аннотативный AttachmentPoint: %d', [AObject.AttachmentPoint]);
  //Направление
  AddMemoMsg('Направление DrawingDirection: %d', [AObject.DrawingDirection]);
  //Высота текста
  AddMemoMsg('Высота Height: %f', [AObject.Height]);
  //Поворот
  AddMemoMsg('Поворот Rotation: %f', [AObject.Rotation]);
  //Межстрочный интервал
  AddMemoMsg('Межстрочный интервал LineSpacingDistance: %f', [AObject.LineSpacingDistance]);
  //Интервал между линиями
  AddMemoMsg('Интервал между линиями LineSpacingFactor: %f', [AObject.LineSpacingFactor]);
  //Стиль меж.строчного интервала
  AddMemoMsg('Стиль меж.строчного интервала LineSpacingStyle: %d', [AObject.LineSpacingStyle]);
  //Скрытие заднего плана
  AddMemoMsg('Скрытие заднего плана BackgroundFill: %d', [Integer(AObject.BackgroundFill)]);
  //Столбцы
  AddMemoMsg('Столбцы Width: %f', [AObject.Width]);
  
  //ГЕОМЕТРИЯ
  //Положение
  AddMemoMsg('Положение InsertionPoint: (%f, %f, %f)', [Double(AObject.InsertionPoint[0]), Double(AObject.InsertionPoint[1]), Double(AObject.InsertionPoint[2])]);

  //ПРОЧИЕ СВОЙСТВА
  //Нормаль
  AddMemoMsg('Normal: (%f; %f; %f)', [Double(AObject.Normal[0]), Double(AObject.Normal[1]), Double(AObject.Normal[2])]);
  //Видимость
  AddMemoMsg('Видимость Visible: %d', [Integer(AObject.Visible)]);
  //Стиль печати
  AddMemoMsg('Стиль печати PlotStyleName: %s', [AObject.PlotStyleName]);

  FAutoCAD.Statistics.Add(AObject.EntityType, AObject.EntityName);
end;{ExtractAcadMText}

//Пауза в миллисекундах
procedure esaPause(const AMsecs: Cardinal = 1500);
var AStart: Cardinal;
begin
  AStart := GetTickCount();
  while not GetTickCount - AStart < AMsecs do
    Application.ProcessMessages();
end;{esaPause}

procedure TfmMain._ExtractAutoCADFile();
var
  AAutoCAD   : IAcadApplication;//Приложение AutoCAD
  ADoc       : IAcadDocument;   //Активный документ AutoCAD
  ALayers    : IAcadLayers;     //Слои AutoCAD
  AObjects   : IAcadModelSpace; //Объекты AutoCAD
  //ABlocks    : IAcadBlocks;     //Блоки AutoCAD

  AObj       : IAcadEntity;     //Текущий объект AutoCAD
  //ABlock     : IAcadBlock;      //Текущий блок AutoCAD
  //ABlockObj  : IAcadEntity;     //Текущий объект блока AutoCAD

  //AUnknown   : IUnknown;        //Неизвестный экземпляр AutoCAD
  //AClassID   : TCLSID;          //Идентификационный номер AutoCAD
  //AIsCreated : Boolean;         //Запущен ли экземпляр AutoCAD
  //AErrorCode : HResult;         //Ошибка, возникающая при неустановленном AutoCAD
  
  I: Integer;
  //J: Integer;
begin
  if OpenDialog.Execute then
  begin
    ExtractAcadClear();
    //AutoCAD.Objects ----------------------------------------------------------
    try
      (*
      //Если AutoCAD не установлен — произойдет ошибка
      AErrorCode := CLSIDFromProgID(PWideChar(WideString('AutoCAD.Application')), AClassID); //Проверка наличия AutoCAD на ПК
      if Succeeded(AErrorCode) then
        AIsCreated := Succeeded(GetActiveObject(AClassID, nil, AUnknown)) //Проверка наличия запущенного AutoCAD
      else
        raise EOleSysError.Create('AutoCAD не найден!', AErrorCode, 0);
      if not AIsCreated then
        //AAutoCAD := IDispatch(CreateComObject(AClassID));
        AAutoCAD := IDispatch(CreateOleObject('AutoCAD.Application')) as IAcadApplication //Запуск AutoCAD
      else
        //AAutoCAD := IDispatch(AUnknown);
        AAutoCAD := IDispatch(AUnknown) as IAcadApplication; //Присоединение к запущенному AutoCAD
      *)

      AAutoCAD := IDispatch(CreateOleObject('AutoCAD.Application')) as IAcadApplication;
      AAutoCAD.Visible := False;
      //acadPause(1500);
      //Активный документ ------------------------------------------------------
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
      //Слои
      ALayers := ADoc.Layers;
      for I := 0 to ALayers.Count - 1 do
        ExtractAcadLayer(I+1, ALayers.Item(I) as IAcadLayer);
      //Объекты
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
      AddMemoMsg('Общее количество блоков в текущем документе: %d', [ABlocks.Count], True);
      for I := 0 to ABlocks.Count-1 do
      begin
        ABlock := ABlocks.Item(I);
        AddMemoMsg('Имя блока #%d : %s', [I+1, ABlock.Name]);
        AddMemoMsg('Количество объектов в блоке: %d', [ABlock.Count]);
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
  stgStatistics.Cells[0, 0] := '№';
  stgStatistics.Cells[1, 0] := 'EntityType';
  stgStatistics.Cells[2, 0] := 'EntityName';
  stgStatistics.Cells[3, 0] := 'Всего';
  stgStatistics.Cells[4, 0] := 'Обработано';
  
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
