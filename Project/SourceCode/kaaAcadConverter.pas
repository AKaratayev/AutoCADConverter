unit kaaAcadConverter;
  { TODO -oKAA -c2013.05.27 : Отключить ПАРАМЕТРЫ-СИСТЕМА-ИНФОЦЕНТР-ВСПЛЫВАЮЩИЕ УВЕДОМЛЕНИЯ-ВКЛЮЧИТЬ ВСПЛЫВАЮЩИЕ УВЕДОМЛЕНИЯ ДЛЯ СЛЕДУЮЩИХ ИСТОЧНИКОВ }
  
interface
uses Classes, SysUtils, Graphics, Types, AutoCAD_TLB;
resourcestring
  EacadInvalidIndex = 'Неверное значение индекса %d. Допустимый диапазон: 0..%d.';
type
  //Исключительные ситуации ----------------------------------------------------
  TacadException = class(Exception)
  public
    //Неверное значение индекса
    constructor InvalidIndex(const AIndex, ACount: Integer);
  end;{TacadException}

  //Тип графического объекта AutoCAD -------------------------------------------
  RacadEntityType = record
    EntityType   : Integer; //Код типа графических объектов
    EntityName   : String;  //Наименование типа графических объектов AutoCAD
    Count        : Integer; //Количество всех графических объектов AutoCAD данного типа
    ImportedCount: Integer; //Количество импортированных графических объектов AutoCAD данного типа
  end;{RacadEntityType}
  PacadEntityType = ^RacadEntityType;

  //Цвет AutoCAD ---------------------------------------------------------------
  RacadColor = record
    R,G,B: Byte;
  end;{RacadColor}

  //Тип линии AutoCAD ----------------------------------------------------------
  TacadPenStyle = (apsSolid, apsDash, apsDot, apsDashDot, apsDashDotDot, apsClear, apsInsideFrame);
  //Тип заливки AutoCAD --------------------------------------------------------
  TacadBrushStyle = (absSolid, absClear, absHorizontal, absVertical, absFDiagonal, absBDiagonal, absCross, absDiagCross);
  //Стиль шрифта AutoCAD -------------------------------------------------------
  TacadFontStyle = (afsNormal, afsBold, afsItalic, afsUnderline, afsStrikeout);
  //Имя шрифта AutoCAD ---------------------------------------------------------
  TacadFontName = (afnArial, afnCalibri, afnCambria, afnCourierNew, afnISOCPEUR, afnISOCTEUR, afnTimesNewRoman, afnVerdana);
  //Горизонтальное выравнивание шрифта AutoCAD ---------------------------------
  TacadHAlign = (ahaLeft, ahaCenter, ahaRight);
  //Вертикальное выравнивание шрифта AutoCAD -----------------------------------
  TacadVAlign = (avaTop, avaCenter, avaBottom);
  //Направление шрифта AutoCAD -------------------------------------------------
  TacadTextDirection = (atdHorizontal, atdVertical);
                           
  //Стили шрифта AutoCAD -------------------------------------------------------
  TacadFontStyles = set of TacadFontStyle;

  //Перо AutoCAD ---------------------------------------------------------------
  RacadPen = record
    Color: RacadColor;   //Цвет
    Width: Double;       //Толщина линии, мм
    Style: TacadPenStyle;//Тип линии
  end;{RacadPen}
  //Заливка AutoCAD ------------------------------------------------------------
  RacadBrush = record
    Color: RacadColor;     //Цвет
    Style: TacadBrushStyle;//Тип заливки
  end;{RacadBrush}
  //Шрифт AutoCAD --------------------------------------------------------------
  RacadFont = record
    Color: RacadColor;     //Цвет
    Size : Double;         //Размер
    Style: TacadFontStyles;//Стиль шрифта
    Name : TacadFontName;  //Имя шрифта
  end;{RacadFont}

  //Координата AutoCAD ---------------------------------------------------------
  TacadCoord3D = array[0..2] of Double;
  //Координаты AutoCAD ---------------------------------------------------------
  TacadCoords3D = array of TacadCoord3D;
  //Треугольник AutoCAD --------------------------------------------------------
  RacadTriangle = record
    p0, p1, p2: TacadCoord3D;//Вершины треугольника
  end;{RacadTriangle}
  //Треугольники AutoCAD -------------------------------------------------------
  TacadTriangles = array of RacadTriangle;
  //Контур AutoCAD -------------------------------------------------------------
  RacadBound = record
    Min: TacadCoord3D;
    Max: TacadCoord3D;
  end;{RacadBound}

  //Предопределенные классы ----------------------------------------------------
  TacadLayer = class; 
  TacadBlock = class;

  //Глобальный объект AutoCAD --------------------------------------------------
  TacadObject = class
  public//Конструктор/деструктор
    constructor Create(); virtual;
  end;{TacadObject}

  //Графический объект AutoCAD -------------------------------------------------
  TacadGraphObject = class(TacadObject)
  private
    FBound  : RacadBound;  //Контур
    FCenter : TacadCoord3D;//Центр
    FVisible: Boolean;        //Признак видимости
  public//Конструктор/деструктор
    constructor Create(); override;
  protected//Методы
    //Прорисовка на канвасе
    procedure DrawToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound); virtual;
    //Определение контура и центра
    procedure _DefineBound(); virtual;
  protected//Свойства
    //Общие
    property Bound  : RacadBound read FBound;
    property Center : TacadCoord3D read FCenter;
    property Visible: Boolean read FVisible;
  end;{TacadGraphObject}

  //Объект слоя AutoCAD --------------------------------------------------------
  TacadLayerObject = class(TacadGraphObject)
  private
    FEntityType: Integer; //Тип объекта AutoCAD
  public//Конструктор/деструктор
    constructor Create(); override;
  protected//Методы
    //Наименование класса объекта AutoCAD
    function GetEntityName(): String;
    //Обработка объекта AutoCAD
    procedure ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock); virtual;
  protected//Свойства
    //Общие
    property EntityType: Integer read FEntityType;
    //Графический объект
    property Bound;
    property Center;
    property Visible;
  end;{TacadLayerObject}
  TacadLayerObjectRef = class of TacadLayerObject;

  //Объект AutoCAD "Точка" -----------------------------------------------------
  TacadPoint = class(TacadLayerObject)
  private
    FColor: RacadColor;  //Цвет
    FSize : Double;         //Толщина, мм
  public//Конструктор/деструктор
    constructor Create(); override;
  public//Методы
    //Обработка точки AutoCAD
    procedure ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock); override;
    //Прорисовка на канвасе
    procedure DrawToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound); override;
    //Определение контура и центра
    procedure _DefineBound(); override;
  public//Свойства
    //Общие
    property EntityType;
    //Графический объект
    property Bound;
    property Center;
    property Visible;
    //Point
    property Color: RacadColor read FColor;
    property Size : Double read FSize;
  end;{TacadPoint}

  //Пользовательский объект CustomPolyline -------------------------------------
  TacadCustomPolyline = class(TacadLayerObject)
  private
    FPen        : RacadPen;     //Перо
    FCoords     : TacadCoords3D;//Вершины
    FCoordsCount: Integer;         //Количество вершин
    FPerimeter  : Double;          //Периметр
  private//Методы диспетчеризации свойств
    function GetCoord(const AIndex: Integer): TacadCoord3D;
    function _GetCoord(const AIndex: Integer): TacadCoord3D;
  protected//Методы
    //Прорисовка на канвасе
    procedure _DrawPlineToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound; const AClosed: Boolean = False);
    //Определение контура и центра
    procedure _DefineBound(); override;
    //Расчет точек контура
    procedure _DefineAdditional(); virtual;
  protected//Методы
    //Прорисовка на канвасе
    procedure DrawToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound); override;
  public//Конструктор/деструктор
    constructor Create(); override;
    destructor Destroy(); override;
  protected//Свойства
    property _Coords[const AIndex: Integer]: TacadCoord3D read _GetCoord;
  protected//Свойства
    //Общие
    property EntityType;
    //Графический объект
    property Bound;
    property Center;
    property Visible;
    //CustomPolyline
    property Pen        : RacadPen read FPen;
    property Coords[const AIndex: Integer]: TacadCoord3D read GetCoord; default;
    property CoordsCount: Integer read FCoordsCount;
    property Perimeter  : Double read FPerimeter;
  end;{TacadCustomPolyline}

  //Блок AutoCAD ---------------------------------------------------------------
  TacadBlock = class(TacadLayerObject)
  private
    FName      : String;     //Имя блока
    FDefaultPen: RacadPen;//Перо по умолчанию
    FItems     : TList;      //Объекты
    FCount     : Integer;    //Количество объектов
  private//Методы диспетчеризации свойств
    function GetItem(const AIndex: Integer): TacadLayerObject;
    function _GetItem(const AIndex: Integer): TacadLayerObject;
  protected//Методы
    //Прорисовка на канвасе
    procedure DrawToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound); override;
    //Очистка
    procedure Clear();
    //Добавление объекта
    procedure Add(const AObject: TacadLayerObject);
    //Определение контура
    procedure _DefineBound(); override;
  public//Методы
    //Обработка блока AutoCAD
    procedure ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock); override;
    //Обработка объектов блока AutoCAD
    procedure ExtractAcadEntityDn(const ALayer: TacadLayer; const AAcadBlocks: IAcadBlocks);
  public//Конструктор/деструктор
    constructor Create(); override;
    destructor Destroy(); override;
  protected//Свойства
    property _Items[const AIndex: Integer]: TacadLayerObject read _GetItem;
  public//Свойства
    //Графический объект
    property Bound;
    property Center;
    property Visible;
    //Block
    property Name: String read FName;
    property DefaultPen: RacadPen read FDefaultPen;
    property Items[const AIndex: Integer]: TacadLayerObject read GetItem; default;
    property Count: Integer read FCount;
  end;{TacadBlock}

  //Объект AutoCAD "Линия" -----------------------------------------------------
  TacadLine = class(TacadCustomPolyline)
  private//Методы диспетчеризации свойств
    function GetStartPoint(): TacadCoord3D;
    function GetEndPoint(): TacadCoord3D;
  public//Методы
    //Обработка линии AutoCAD
    procedure ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock); override;
  public//Свойства
    //Общие
    property EntityType;
    //Графический объект
    property Bound;
    property Center;
    property Visible;
    //CustomPolyline
    property Pen;
    property Coords;
    property CoordsCount;
    property Perimeter;
    //Line
    property StartPoint: TacadCoord3D read GetStartPoint;
    property EndPoint  : TacadCoord3D read GetEndPoint;
  end;{TacadLine}

  //Объект AutoCAD "Полилиния" -------------------------------------------------
  TacadPolyline = class(TacadCustomPolyline)
  private
    FClosed: Boolean;//Признак замкнутости полилинии
  public//Конструктор/деструктор
    constructor Create(); override;
  public//Методы
    //Обработка полилинии AutoCAD
    procedure ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock); override;
    //Прорисовка на канвасе
    procedure DrawToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound); override;
    //Определение контура и центра
    procedure _DefineBound(); override;
  public//Свойства
    //Общие
    property EntityType;
    //Графический объект
    property Bound;
    property Center;
    property Visible;
    //CustomPolyline
    property Pen;
    property Coords;
    property CoordsCount;
    property Perimeter;
    //Polyline
    property Closed: Boolean read FClosed;
  end;{TacadPolyline}

  //Объект AutoCAD "3D Полилиния" ----------------------------------------------
  Tacad3DPolyline = class(TacadCustomPolyline)
  private
    FClosed: Boolean;//Признак замкнутости полилинии
  public//Конструктор/деструктор
    constructor Create(); override;
  public//Методы
    //Обработка полилинии AutoCAD
    procedure ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock); override;
    //Прорисовка на канвасе
    procedure DrawToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound); override;
    //Определение контура и центра
    procedure _DefineBound(); override;
  public//Свойства
    //Общие
    property EntityType;
    //Графический объект
    property Bound;
    property Center;
    property Visible;
    //CustomPolyline
    property Pen;
    property Coords;
    property CoordsCount;
    property Perimeter;
    //Polyline
    property Closed: Boolean read FClosed;
  end;{TacadPolyline}

  //Объект AutoCAD "Дуга" ------------------------------------------------------
  TacadArc = class(TacadCustomPolyline)
  private
    FRadius    : Double;         //Радиус
    FStartPoint: TacadCoord3D;//Начальная точка
    FEndPoint  : TacadCoord3D;//Конечная точка
    FStartAngle: Double;         //Начальный угол, рад.
    FEndAngle  : Double;         //Конечный угол, рад.
  public//Конструктор/деструктор
    constructor Create(); override;
  public//Методы
    //Обработка дуги AutoCAD
    procedure ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock); override;
  protected//Внутренние методы
    //Расчет точек контура
    procedure _DefineAdditional(); override;
  public//Свойства
    //Общие
    property EntityType;
    //Графический объект
    property Bound;
    property Center;
    property Visible;
    //CustomPolyline
    property Pen;
    property Coords;
    property CoordsCount;
    property Perimeter;
    //Arc
    property Radius    : Double read FRadius;
    property StartPoint: TacadCoord3D read FStartPoint;
    property EndPoint  : TacadCoord3D read FEndPoint;
    property StartAngle: Double read FStartAngle;
    property EndAngle  : Double read FEndAngle;
  end;{TacadArc}

  //Тип сплайна ----------------------------------------------------------------
  TacadSplineKind = (askQuadratic, ascCubic);//? 
  //Объект AutoCAD "Сплайн" ----------------------------------------------------
  TacadSpline = class(TacadCustomPolyline)
  private
    FClosed    : Boolean;           //Признак замкнутости сплайна
    FNodes     : TacadCoords3D;  //Определяющие вершины
    FNodesCount: Integer;           //Количество определяющих вершин
    FKind      : TacadSplineKind;//Тип сплайна
  private//Методы диспетчеризации свойств
    function GetNode(const AIndex: Integer): TacadCoord3D;
    function _GetNode(const AIndex: Integer): TacadCoord3D;
  public//Конструктор/деструктор
    constructor Create(); override;
    destructor Destroy(); override;
  public//Методы
    //Обработка сплайна AutoCAD
    procedure ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock); override;
  protected//Внутренние методы
    //Расчет точек контура
    procedure _DefineAdditional(); override;
  protected//Свойства
    property _Nodes[const AIndex: Integer]: TacadCoord3D read _GetNode;
  public//Свойства
    //Общие
    property EntityType;
    //Графический объект
    property Bound;
    property Center;
    property Visible;
    //CustomPolyline
    property Pen;
    property Coords;
    property CoordsCount;
    property Perimeter;
    //Spline
    property Closed    : Boolean read FClosed;
    property Nodes[const AIndex: Integer]: TacadCoord3D read GetNode;
    property NodesCount: Integer read FNodesCount;
  end;{TacadSpline}

  //Пользовательский объект CustomPolygon --------------------------------------
  TacadCustomPolygon = class(TacadCustomPolyline)
  private
    FBrush         : RacadBrush;    //Заливка
    FArea          : Double;           //Площадь
    FTriangles     : TacadTriangles;//Треугольники
    FTrianglesCount: Integer;          //Количество треугольников
  private//Методы диспетчеризации свойств
    function GetTriangle(const AIndex: Integer): RacadTriangle;
    function _GetTriangle(const AIndex: Integer): RacadTriangle;
  public//Конструктор/деструктор
    constructor Create(); override;
    destructor Destroy(); override;
  protected
    //Прорисовка на канвасе
    procedure _DrawPgonToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound);
  protected//Методы
    //Прорисовка на канвасе
    procedure DrawToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound); override;
    //Определение контура и центра
    procedure _DefineBound(); override;
    //Расчет точек контура и треугольников
    procedure _DefineAdditional(); override;
  protected//Свойства
    property _Triangles[const AIndex: Integer]: RacadTriangle read _GetTriangle;
  protected//Свойства
    //Общие
    property EntityType;
    //Графический объект
    property Bound;
    property Center;
    property Visible;
    //CustomPolyline
    property Pen;
    property Coords;
    property CoordsCount;
    property Perimeter;
    //CustomPolygon
    property Brush: RacadBrush read FBrush;
    property Area : Double read FArea;
    property Triangles[const AIndex: Integer]: RacadTriangle read GetTriangle;
    property TrianglesCount: Integer read FTrianglesCount;
  end;{TacadCustomPolygon}
  
  //Объект AutoCAD "Полигон" ---------------------------------------------------
  TacadPolygon = class(TacadCustomPolygon)
  public//Методы
    //Обработка полигона AutoCAD
    procedure ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock); override;
  protected//Внутренние методы
    //Расчет точек контура и треугольников
    procedure _DefineAdditional(); override;
  public//Свойства
    //Общие
    property EntityType;
    //Графический объект
    property Bound;
    property Center;
    property Visible;
    //CustomPolyline
    property Pen;
    property Perimeter;
    //CustomPolygon
    property Brush;
    property Area;
  protected//Свойства
    //CustomPolyline
    property Coords;
    property CoordsCount;
    property Triangles;
    property TrianglesCount;
  end;{TacadPolygon}

  //Объект AutoCAD "Круг" ------------------------------------------------------
  TacadCircle = class(TacadCustomPolygon)
  private
    FRadius: Double;//Радиус
  public//Конструктор/деструктор
    constructor Create(); override;
  public//Методы
    //Обработка круга AutoCAD
    procedure ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock); override;
  protected//Внутренние методы
    //Расчет точек контура и треугольников
    procedure _DefineAdditional(); override;
  public//Свойства
    //Общие
    property EntityType;
    //Графический объект
    property Bound;
    property Center;
    property Visible;
    //CustomPolyline
    property Pen;
    property Perimeter;
    //CustomPolygon
    property Brush;
    property Area;
    //Circle
    property Radius: Double read FRadius;
  protected//Свойства
    //CustomPolyline
    property Coords;
    property CoordsCount;
    property Triangles;
    property TrianglesCount;
  end;{TacadCircle}

  //Объект AutoCAD "Эллипс" ----------------------------------------------------
  TacadEllipse = class(TacadCustomPolygon)
  private
    FMajorRadius: Double;//Большая полуось
    FMinorRadius: Double;//Малая полуось
    FMajorAxis  : Double;//Вектор X большей полуоси
  public//Конструктор/деструктор
    constructor Create(); override;
  public//Методы
    //Обработка эллипса AutoCAD
    procedure ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock); override;
  protected//Внутренние методы
    //Расчет точек контура и треугольников
    procedure _DefineAdditional(); override;
  public//Свойства
    //Общие
    property EntityType;
    //Графический объект
    property Bound;
    property Center;
    property Visible;
    //CustomPolyline
    property Pen;
    property Perimeter;
    //CustomPolygon
    property Brush;
    property Area;
    //Ellipse
    property MajorRadius: Double read FMajorRadius;
    property MinorRadius: Double read FMinorRadius;
    property MajorAxis  : Double read FMajorAxis;
  protected//Свойства
    //CustomPolyline
    property Coords;
    property CoordsCount;
    property Triangles;
    property TrianglesCount;
  end;{TacadEllipse}

  //Объект AutoCAD "Текст" -----------------------------------------------------
  TacadText = class(TacadCustomPolygon)
  private
    FFont         : RacadFont;         //Шрифт
    FCaption      : String;               //Заголовок
    FHAlign       : TacadHAlign;       //Горизонтальное выравнивание
    FVAlign       : TacadVAlign;       //Вертикальное выравнивание
    FRotation     : Double;               //Поворот текста, град. 0-360
    FScaleFactor  : Double;               //Коэффициент плотности, 0-1
    FObliqueAngle : Double;               //Угол наклона
    FTextDirection: TacadTextDirection;//Направление
    FUpsideDown   : Boolean;              //Признак перевернутого текста
    FBackward     : Boolean;              //Признак обратного написания
  public//Конструктор/деструктор
    constructor Create(); override;
  public//Методы
    //Обработка текста AutoCAD
    procedure ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock); override;
    //Прорисовка на канвасе
    procedure DrawToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound); override;
  protected//Свойства
    //Общие
    property EntityType;
    //Графический объект
    property Bound;
    property Center;
    property Visible;
    //CustomPolyline
    property Pen;
    property Coords;
    property CoordsCount;
    property Perimeter;
    //CustomPolygon
    property Brush;
    property Area;
    //Text
    property Font         : RacadFont read FFont;
    property Caption      : String read FCaption;
    property HAlign       : TacadHAlign read FHAlign;
    property VAlign       : TacadVAlign read FVAlign;
    property Rotation     : Double read FRotation;
    property ScaleFactor  : Double read FScaleFactor;
    property ObliqueAngle : Double read FObliqueAngle;
    property TextDirection: TacadTextDirection read FTextDirection;
    property UpsideDown   : Boolean read FUpsideDown;
    property Backward     : Boolean read FBackward;
  end;{TacadText}

  //Слой AutoCAD ---------------------------------------------------------------
  TacadLayer = class(TacadGraphObject)
  private
    FName          : String;     //Название слоя
    FFreeze        : Boolean;    //Признак замораживания слоя - невидимы, исключены из процессов регенерации и печати
    FLock          : Boolean;    //Признак запрета редактирования объектов слоя
    FDefaultPen    : RacadPen;//Перо по умолчанию
    FDescription   : String;     //Примечание
    FItems         : TList;      //Объекты
    FCount         : Integer;    //Количество объектов
  private//Методы диспетчеризации свойств
    function GetItem(const AIndex: Integer): TacadLayerObject;
    function _GetItem(const AIndex: Integer): TacadLayerObject;
  protected//Методы
    //Прорисовка на канвасе
    procedure DrawToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound); override;
    //Очистка
    procedure Clear();
    //Добавление объекта
    procedure Add(const AObject: TacadLayerObject);
    //Определение контура
    procedure _DefineBound(); override;
  public//Конструктор/деструктор
    constructor Create(); override;
    destructor Destroy(); override;
  protected//Свойства
    property _Items[const AIndex: Integer]: TacadLayerObject read _GetItem;
  public//Свойства
    //Графический объект
    property Bound;
    property Center;
    property Visible;
    //Layer
    property Name        : String read FName;
    property Freeze      : Boolean read FFreeze;
    property Lock        : Boolean read FLock;
    property DefaultPen  : RacadPen read FDefaultPen;
    property Description : String read FDescription;
    property Items[const AIndex: Integer]: TacadLayerObject read GetItem; default;
    property Count: Integer read FCount;
  end;{TacadLayer}

  //Статистика типов объектов AutoCAD ------------------------------------------
  TacadEntitiesStatistic = class(TacadObject)
  private
    FItems: TList;
    FCount: Integer;
  private//Методы диспетчеризации свойств
    function _GetItem(const AIndex: Integer): RacadEntityType;
    function GetItem(const AIndex: Integer): RacadEntityType;
  protected//Методы
  public//Методы - временно public
    //Добавление
    procedure Add(const AEntityType: Integer; const AEntityName: String; const AUnknown: Boolean = False);
  protected//Методы
    //Очистка
    procedure Clear();
  public//Методы - временно public
    //Поиск
    function IndexOf(const AEntityType: Integer): Integer;
  public//Конструктор/деструктор
    constructor Create(); override;
    destructor Destroy(); override;
  protected//Свойства
    property _Items[const AIndex: Integer]: RacadEntityType read _GetItem;
  public//Свойства
    property Items[const AIndex: Integer]: RacadEntityType read GetItem; default;
    property Count: Integer read FCount;
  end;{TacadEntitiesStatistic}

  //Приложение AutoCAD ---------------------------------------------------------
  TAutoCAD = class
  private
    FCubeBound  : RacadBound;            //Контур слоев
    FCenter     : TacadCoord3D;          //Центр куба данных
    FStatistics : TacadEntitiesStatistic;//Статистика объектов
    FLayers     : TList;                    //Слои
    FLayersCount: Integer;                  //Количество слоев
  private//Методы диспетчеризации свойств
    function GetLayer(const AIndex: Integer): TacadLayer;
    function _GetLayer(const AIndex: Integer): TacadLayer;
  protected//Методы
    //Извлечение слоя AutoCAD
    procedure _ExtractAutoCADLayer(const AacadLayer: IAcadLayer);
    //Поиск слоя
    function _FindLayer(const ALayerName: String): Integer;
    //Уничтожение слоев
    procedure _ClearLayers();
    //Определение глобального контура
    procedure _DefineBound();
  public//Методы
    //Прорисовка
    procedure Draw(const ACanvas: TCanvas; const ACanvasBound: TRect);
    //Очистка
    procedure Clear();
    //Импорт из файла AutoCAD
    function ImportFromAutoCADFile(const AFileName: String): Boolean;
  public//Конструктор/деструктор
    constructor Create(); virtual;
    destructor Destroy(); override;
  protected//Свойства
    property _Layers[const AIndex: Integer]: TacadLayer read _GetLayer;
  public//Свойства
    property CubeBound: RacadBound read FCubeBound;
    property Center: TacadCoord3D read FCenter;
    property Layers[const AIndex: Integer]: TacadLayer read GetLayer; default;
    property LayersCount: Integer read FLayersCount;
    property Statistics: TacadEntitiesStatistic read FStatistics;
  end;{TAutoCAD}
  
//Цвет AutoCAD -----------------------------------------------------------------
function acadColor(const AR,AG,AB: Byte): RacadColor;
//Перо AutoCAD -----------------------------------------------------------------
function acadPen(const AColor: RacadColor; const AWidth: Double; const AStyle: TacadPenStyle): RacadPen;
//Заливка AutoCAD --------------------------------------------------------------
function acadBrush(const AColor: RacadColor; const AStyle: TacadBrushStyle): RacadBrush;
//Шрифт AutoCAD ----------------------------------------------------------------
function acadFont(const AColor: RacadColor; const ASize: Double; const AStyle: TacadFontStyles; const AName: TacadFontName): RacadFont;
//Координата AutoCAD -----------------------------------------------------------
function acadCoord3D(): TacadCoord3D; overload;
function acadCoord3D(const AX,AY,AZ: Double): TacadCoord3D; overload;
//Контур AutoCAD ---------------------------------------------------------------
function acadBound(): RacadBound; overload;
function acadBound(const AMin, AMax: TacadCoord3D): RacadBound; overload;
function acadCenter(const ABound: RacadBound): TacadCoord3D; 
//Перевод мировых координат в экранные
function acadCoordTo(const AX,AY: Double; const ACanvasBound: TRect; const ACubeBound: RacadBound): TPoint;
//Общие ------------------------------------------------------------------------
//Пауза в миллисекундах
procedure acadPause(const AMsecs: Cardinal = 1500);
//Функции перевода величин AutoCAD ---------------------------------------------
//Толщина линии
function acadLineSizeTo(const ALineSize: Cardinal; const ALayerLineSize, ABlockLineSize: Double): Double;
//Тип линии
function acadLineTypeTo(const ALineType: String; const ALayerLineType, ABlockLineType: TacadPenStyle): TacadPenStyle;
//Наименование шрифта
function acadFontNameTo(const AFontName: String): TacadFontName ;
//Тип заливки
function acadBrushStyleTo(const ABrushStyle: String): TacadBrushStyle ;

implementation
uses Windows, Math, ComObj, Forms, Variants;

//Цвет AutoCAD -----------------------------------------------------------------
function acadColor(const AR,AG,AB: Byte): RacadColor;
begin
  Result.R := AR;
  Result.G := AG;
  Result.B := AB;
end;{acadColor}
//Перо AutoCAD -----------------------------------------------------------------
function acadPen(const AColor: RacadColor; const AWidth: Double; const AStyle: TacadPenStyle): RacadPen;
begin
  Result.Color := AColor;
  Result.Width := AWidth;
  Result.Style := AStyle;
end;{acadPen}
//Заливка AutoCAD --------------------------------------------------------------
function acadBrush(const AColor: RacadColor; const AStyle: TacadBrushStyle): RacadBrush;
begin
  Result.Color := AColor;
  Result.Style := AStyle;
end;{acadBrush}
//Шрифт AutoCAD ----------------------------------------------------------------
function acadFont(const AColor: RacadColor; const ASize: Double; const AStyle: TacadFontStyles; const AName: TacadFontName): RacadFont;
begin
  Result.Color := AColor;
  Result.Size  := ASize;
  Result.Style := AStyle;
  Result.Name  := AName;
end;{acadFont}
//Координата AutoCAD -----------------------------------------------------------
function acadCoord3D(): TacadCoord3D;
begin
  Result := acadCoord3D(0.0,0.0,0.0);
end;{acadCoord3D}
function acadCoord3D(const AX,AY,AZ: Double): TacadCoord3D; 
begin
  Result[0] := AX;
  Result[1] := AY;
  Result[2] := AZ;
end;{acadCoord3D}
//Контур
function acadBound(): RacadBound;
begin
  Result := acadBound(acadCoord3D(), acadCoord3D());
end;{acadBound}
function acadBound(const AMin, AMax: TacadCoord3D): RacadBound;
begin
  Result.Min := AMin;
  Result.Max := AMax;
end;{acadBound}
function acadCenter(const ABound: RacadBound): TacadCoord3D;
begin
  Result := acadCoord3D(0.5*(ABound.Min[0]+ABound.Max[0]),
                           0.5*(ABound.Min[1]+ABound.Max[1]),
                           0.5*(ABound.Min[2]+ABound.Max[2]));
end;{acadCenter}
//Перевод мировых координат в экранные
function acadCoordTo(const AX,AY: Double; const ACanvasBound: TRect; const ACubeBound: RacadBound): TPoint;
begin
  Result.X := Round(ACanvasBound.Left+(ACanvasBound.Right-ACanvasBound.Left)*(AX-ACubeBound.Min[0])/(ACubeBound.Max[0]-ACubeBound.Min[0]));
  Result.Y := Round(ACanvasBound.Top+(ACanvasBound.Bottom-ACanvasBound.Top)*(1-(AY-ACubeBound.Min[1])/(ACubeBound.Max[1]-ACubeBound.Min[1])));
end;{acadCoordTo}
//Общие ------------------------------------------------------------------------
//Пауза в миллисекундах
procedure acadPause(const AMsecs: Cardinal = 1500);
var AStart: Cardinal;
begin
  AStart := GetTickCount();
  while not GetTickCount - AStart < AMsecs do
    Application.ProcessMessages();
end;{acadPause}
//Функции перевода величин AutoCAD ---------------------------------------------
//Толщина линии
function acadLineSizeTo(const ALineSize: Cardinal; const ALayerLineSize, ABlockLineSize: Double): Double;
begin
  case ALineSize of
    acLnWtByLayer    : Result := ALayerLineSize;
    acLnWtByBlock    : Result := ABlockLineSize;
    acLnWtByLwDefault: Result := 0.25;
    acLnWt000        : Result := 0.0;
    acLnWt005        : Result := 0.05;
    acLnWt009        : Result := 0.09;
    acLnWt013        : Result := 0.13;
    acLnWt015        : Result := 0.15;
    acLnWt018        : Result := 0.18;
    acLnWt020        : Result := 0.2;
    acLnWt025        : Result := 0.25;
    acLnWt030        : Result := 0.3;
    acLnWt035        : Result := 0.35;
    acLnWt040        : Result := 0.4;
    acLnWt050        : Result := 0.5;
    acLnWt053        : Result := 0.53;
    acLnWt060        : Result := 0.6;
    acLnWt070        : Result := 0.7;
    acLnWt080        : Result := 0.8;
    acLnWt090        : Result := 0.9;
    acLnWt100        : Result := 1.0;
    acLnWt106        : Result := 1.06;
    acLnWt120        : Result := 1.2;
    acLnWt140        : Result := 1.4;
    acLnWt158        : Result := 1.58;
    acLnWt200        : Result := 2.0;
    acLnWt211        : Result := 2.11;
  else Result := 0.0;
  end;{case}
end;{acadLineSizeTo}
//Тип линии
function acadLineTypeTo(const ALineType: String; const ALayerLineType, ABlockLineType: TacadPenStyle): TacadPenStyle;
begin
  if ALineType = 'ByLayer' then Result := ALayerLineType else
  if ALineType = 'ByBlock' then Result := ABlockLineType else

  if ALineType = 'Continuous'     then Result := apsSolid else
  if ALineType = 'ACAD_ISO02W100' then Result := apsDash else
  if ALineType = 'ACAD_ISO03W100' then Result := apsDash else
  if ALineType = 'ACAD_ISO04W100' then Result := apsDashDot else
  if ALineType = 'ACAD_ISO05W100' then Result := apsDashDotDot else
  if ALineType = 'ACAD_ISO06W100' then Result := apsDashDotDot else
  if ALineType = 'ACAD_ISO07W100' then Result := apsDot else
  if ALineType = 'ACAD_ISO08W100' then Result := apsDash else
  if ALineType = 'ACAD_ISO09W100' then Result := apsDash else
  if ALineType = 'ACAD_ISO10W100' then Result := apsDashDot else
  if ALineType = 'ACAD_ISO11W100' then Result := apsDashDot else
  if ALineType = 'ACAD_ISO12W100' then Result := apsDashDotDot else
  if ALineType = 'ACAD_ISO13W100' then Result := apsDashDot else
  if ALineType = 'ACAD_ISO14W100' then Result := apsDashDotDot else
  if ALineType = 'ACAD_ISO15W100' then Result := apsDashDotDot else
                                                 
  if ALineType = 'JIS_02_0.7' then Result := apsDash else
  if ALineType = 'JIS_02_1.0' then Result := apsDash else
  if ALineType = 'JIS_02_1.2' then Result := apsDash else
  if ALineType = 'JIS_02_2.0' then Result := apsDash else
  if ALineType = 'JIS_02_4.0' then Result := apsDash else
  if ALineType = 'JIS_08_11'  then Result := apsDash else
  if ALineType = 'JIS_08_15'  then Result := apsDash else
  if ALineType = 'JIS_08_25'  then Result := apsDash else
  if ALineType = 'JIS_08_37'  then Result := apsDash else
  if ALineType = 'JIS_08_50'  then Result := apsDash else
  if ALineType = 'JIS_09_08'  then Result := apsDash else
  if ALineType = 'JIS_09_15'  then Result := apsDash else
  if ALineType = 'JIS_09_29'  then Result := apsDash else
  if ALineType = 'JIS_09_50'  then Result := apsDash else

  if ALineType = 'линия_сгиба'       then Result := apsDashDotDot else
  if ALineType = 'линия_сгиба2'      then Result := apsDashDotDot else
  if ALineType = 'линия_сгибаX2'     then Result := apsDashDotDot else
  if ALineType = 'невидимая'         then Result := apsDash else
  if ALineType = 'невидимая2'        then Result := apsDash else
  if ALineType = 'невидимаяX2'       then Result := apsDash else
  if ALineType = 'осевая'            then Result := apsDash else
  if ALineType = 'осевая2'           then Result := apsDash else
  if ALineType = 'осеваяX2'          then Result := apsDash else
  if ALineType = 'пунктирная'        then Result := apsDot else
  if ALineType = 'пунктирная2'       then Result := apsDot else
  if ALineType = 'пунктирнаяX2'      then Result := apsDot else
  if ALineType = 'рант'              then Result := apsDashDot else
  if ALineType = 'рант2'             then Result := apsDashDot else
  if ALineType = 'рантX2'            then Result := apsDashDot else
  if ALineType = 'фантом'            then Result := apsDash else
  if ALineType = 'фантом2'           then Result := apsDash else
  if ALineType = 'фантомX2'          then Result := apsDash else
  if ALineType = 'штриховая'         then Result := apsDash else
  if ALineType = 'штриховая2'        then Result := apsDash else
  if ALineType = 'штриховаяX2'       then Result := apsDash else
  if ALineType = 'штрихпунктирная'   then Result := apsDashDot else
  if ALineType = 'штрихпунктирная2'  then Result := apsDashDot else
  if ALineType = 'штрихпунктирнаяX2' then Result := apsDashDot

  else Result := apsClear;
end;{acadLineTypeTo}
//Наименование шрифта
function acadFontNameTo(const AFontName: String): TacadFontName;
begin
  if AFontName = 'arial.ttf'      then Result := afnArial else
  if AFontName = 'calibri.ttf'    then Result := afnCalibri else
  if AFontName = 'cambria.ttf'    then Result := afnCambria else
  if AFontName = 'couriernew.ttf' then Result := afnCourierNew else
  if AFontName = 'isocpeur'       then Result := afnISOCPEUR else
  if AFontName = 'isocteur'       then Result := afnISOCTEUR else
  if AFontName = 'times.ttf'      then Result := afnTimesNewRoman else
  if AFontName = 'verdana.ttf'    then Result := afnVerdana
  //...
  else Result := afnArial;
end;{acadFontNameTo}
//Тип заливки
function acadBrushStyleTo(const ABrushStyle: String): TacadBrushStyle;
begin
  if ABrushStyle = 'SOLID' then Result := absSolid else

  if ABrushStyle = 'ANSI31' then Result := absBDiagonal else
  if ABrushStyle = 'ANSI32' then Result := absBDiagonal else
  if ABrushStyle = 'ANSI33' then Result := absBDiagonal else
  if ABrushStyle = 'ANSI34' then Result := absBDiagonal else
  if ABrushStyle = 'ANSI35' then Result := absBDiagonal else
  if ABrushStyle = 'ANSI36' then Result := absBDiagonal else
  if ABrushStyle = 'ANSI37' then Result := absDiagCross else
  if ABrushStyle = 'ANSI38' then Result := absDiagCross else

  if ABrushStyle = 'ISO02W100' then Result := absHorizontal else
  if ABrushStyle = 'ISO03W100' then Result := absHorizontal else
  if ABrushStyle = 'ISO04W100' then Result := absHorizontal else
  if ABrushStyle = 'ISO05W100' then Result := absHorizontal else
  if ABrushStyle = 'ISO06W100' then Result := absHorizontal else
  if ABrushStyle = 'ISO07W100' then Result := absHorizontal else
  if ABrushStyle = 'ISO08W100' then Result := absHorizontal else
  if ABrushStyle = 'ISO09W100' then Result := absHorizontal else
  if ABrushStyle = 'ISO10W100' then Result := absHorizontal else
  if ABrushStyle = 'ISO11W100' then Result := absHorizontal else
  if ABrushStyle = 'ISO12W100' then Result := absHorizontal else
  if ABrushStyle = 'ISO13W100' then Result := absHorizontal else
  if ABrushStyle = 'ISO14W100' then Result := absHorizontal else
  if ABrushStyle = 'ISO15W100' then Result := absHorizontal else

  if ABrushStyle = 'ANGLE'     then Result := absCross else
  if ABrushStyle = 'AR-B816'   then Result := absCross else
  if ABrushStyle = 'AR-B816C'  then Result := absCross else
  if ABrushStyle = 'AR-B88'    then Result := absCross else
  if ABrushStyle = 'AR-BRELM'  then Result := absHorizontal else
  if ABrushStyle = 'AR-BRSTD'  then Result := absCross else
  if ABrushStyle = 'AR-CONC'   then Result := absSolid else
  if ABrushStyle = 'AAR-HBONE' then Result := absDiagCross else
  if ABrushStyle = 'AR-PARQ1'  then Result := absCross else
  if ABrushStyle = 'AR-RROOF'  then Result := absSolid else
  if ABrushStyle = 'AR-RSHKE'  then Result := absCross else

  if ABrushStyle = 'AR-SAND'     then Result := absSolid else
  if ABrushStyle = 'BOX'         then Result := absVertical else
  if ABrushStyle = 'BRASS'       then Result := absHorizontal else
  if ABrushStyle = 'BRICK'       then Result := absCross else
  if ABrushStyle = 'BRSTONE'     then Result := absCross else
  if ABrushStyle = 'CLAY'        then Result := absHorizontal else
  if ABrushStyle = 'CORK'        then Result := absFDiagonal else
  if ABrushStyle = 'CROSS'       then Result := absCross else
  if ABrushStyle = 'DASH'        then Result := absHorizontal else
  if ABrushStyle = 'DOLMIT'      then Result := absBDiagonal else
  if ABrushStyle = 'DOTS'        then Result := absHorizontal else
  if ABrushStyle = 'EARTH'       then Result := absCross else
  if ABrushStyle = 'ESCHER'      then Result := absSolid else
  if ABrushStyle = 'FLEX'        then Result := absHorizontal else
  if ABrushStyle = 'GOST_GLASS'  then Result := absCross else
  if ABrushStyle = 'GOST_WOOD'   then Result := absVertical else
  if ABrushStyle = 'GOST_GROUND' then Result := absBDiagonal else
  if ABrushStyle = 'GRASS'       then Result := absCross else
  if ABrushStyle = 'GRATE'       then Result := absCross else
  if ABrushStyle = 'GRAVEL'      then Result := absSolid else
  if ABrushStyle = 'HEX'         then Result := absSolid else
  if ABrushStyle = 'HONEY'       then Result := absDiagCross else
  if ABrushStyle = 'HOUND'       then Result := absCross else
  if ABrushStyle = 'INSUL'       then Result := absHorizontal else

  if ABrushStyle = 'JIS_LC_20'  then Result := absBDiagonal else
  if ABrushStyle = 'JIS_LC_20A' then Result := absBDiagonal else
  if ABrushStyle = 'JIS_LC_8'   then Result := absBDiagonal else
  if ABrushStyle = 'JIS_LC_8A'  then Result := absBDiagonal else
  if ABrushStyle = 'JIS_RC_10'  then Result := absBDiagonal else
  if ABrushStyle = 'JIS_RC_15'  then Result := absBDiagonal else
  if ABrushStyle = 'JIS_RC_18'  then Result := absBDiagonal else
  if ABrushStyle = 'JIS_RC_30'  then Result := absBDiagonal else
  if ABrushStyle = 'JIS_STN_1E' then Result := absBDiagonal else
  if ABrushStyle = 'JIS_WOOD'   then Result := absBDiagonal else
  if ABrushStyle = 'LINE'       then Result := absHorizontal else
  if ABrushStyle = 'MUDST'      then Result := absHorizontal else
  if ABrushStyle = 'NET'        then Result := absCross else
  if ABrushStyle = 'NET3'       then Result := absDiagCross else
  if ABrushStyle = 'PLAST'      then Result := absHorizontal else
  if ABrushStyle = 'PLASTI'     then Result := absHorizontal else
  if ABrushStyle = 'SACNCR'     then Result := absBDiagonal else
  if ABrushStyle = 'SQUARE'     then Result := absSolid else
  if ABrushStyle = 'STARS'      then Result := absSolid else
  if ABrushStyle = 'STEEL'      then Result := absBDiagonal else
  if ABrushStyle = 'SWAMP'      then Result := absHorizontal else
  if ABrushStyle = 'TRANS'      then Result := absHorizontal else
  if ABrushStyle = 'TRIANG'     then Result := absSolid else
  if ABrushStyle = 'ZIGZAG'     then Result := absBDiagonal   

  else Result := absClear;
end;{acadBrushStyleTo}

//Исключительные ситуации ------------------------------------------------------
//Неверное значение индекса
constructor TacadException.InvalidIndex(const AIndex,ACount: Integer);
begin
  CreateResFmt(@EacadInvalidIndex,[AIndex,ACount-1]);
end;{InvalidIndex}

//Глобальный объект AutoCAD ----------------------------------------------------
//Конструктор/деструктор
constructor TacadObject.Create();
begin
  inherited;
end;{Create}

//Графический объект AutoCAD ---------------------------------------------------
//Конструктор/деструктор
constructor TacadGraphObject.Create();
begin
  inherited;
  FBound      := acadBound();
  FCenter     := acadCoord3D();
  FVisible    := True;
end;{Create}
//Прорисовка на канвасе
procedure TacadGraphObject.DrawToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound);
begin
end;{DrawToCanvas}
//Определение контура и центра
procedure TacadGraphObject._DefineBound();
begin
  FCenter := acadCenter(FBound);
end;{_DefineBound}

//Объект слоя AutoCAD ----------------------------------------------------------
//Конструктор/деструктор
constructor TacadLayerObject.Create();
begin
  inherited;
  FEntityType := 0;
end;{Create}
//Наименование класса объекта AutoCAD
function TacadLayerObject.GetEntityName(): String;
begin
  case EntityType of
    ac3dPolyline    : Result := 'IAcad3DPolyline';    //$00000002  (2)
    acArc           : Result := 'IAcadArc';           //$00000004  (4)
    acBlockReference: Result := 'IAcadBlockReference';//$00000007  (7)
    acCircle        : Result := 'IAcadCircle';        //$00000008  (8)
    acEllipse       : Result := 'IAcadEllipse';       //$00000010  (16)
    acHatch         : Result := 'IAcadHatch';         //$00000011  (17)
    acLine          : Result := 'IAcadLine';          //$00000013  (19)
    acMtext         : Result := 'IAcadMText';         //$00000015  (20)
    acPoint         : Result := 'IAcadPoint';         //$00000016  (21)
    acPolylineLight : Result := 'IAcadLWPolyline';    //$00000018  (24)
    acSpline        : Result := 'IAcadSpline';        //$0000001F
    acText          : Result := 'IAcadText';          //$00000020
    else Result := '';
  end;{case}
end;{GetEntityName}
//Обработка свойств объекта AutoCAD
procedure TacadLayerObject.ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock);
begin
end;{ExtractAcadEntity}

//Объект AutoCAD "Точка" -------------------------------------------------------
//Конструктор/деструктор
constructor TacadPoint.Create();
begin
  inherited;
  FColor := acadColor(255,255,255);
  FSize  := 1.0;
end;{Create}
//Извлечение точки AutoCAD
procedure TacadPoint.ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock);
var
  AAcadPoint    : IAcadPoint;
  ABlockPenWidth: Double;
  ABlockCenter  : TacadCoord3D;
begin
  AAcadPoint     := AAcadObj as IAcadPoint;
  //Block
  ABlockPenWidth := 0.0;
  ABlockCenter   := acadCoord3D();
  if ABlock <> nil then
  begin
    ABlockPenWidth := ABlock.DefaultPen.Width;
    ABlockCenter   := acadCoord3D(ABlock.Center[0], ABlock.Center[1], ABlock.Center[2]);
  end;{if}
  //Общие
  FEntityType := AAcadPoint.EntityType;
  //Графический объект

  FSize       := acadLineSizeTo(AAcadPoint.Lineweight, ALayer.DefaultPen.Width, ABlockPenWidth);
  FVisible    := AAcadPoint.Visible = True;
  //Point
  FColor      := acadColor(AAcadPoint.TrueColor.Red, AAcadPoint.TrueColor.Green, AAcadPoint.TrueColor.Blue);
  FCenter     := acadCoord3D(AAcadPoint.Coordinates[0]+ABlockCenter[0], AAcadPoint.Coordinates[1]+ABlockCenter[1], AAcadPoint.Coordinates[2]+ABlockCenter[2]);
  //Bound
  _DefineBound();
end;{ExtractAcadEntity}
//Прорисовка на канвасе
procedure TacadPoint.DrawToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound);
var p0,p1: TPoint;
begin
  if Visible then
  begin
    ACanvas.Brush.Color := RGB(Color.R,Color.G,Color.B);
    ACanvas.Pen.Color   := ACanvas.Brush.Color;
    p0 := acadCoordTo(Bound.Min[0], Bound.Min[1], ACanvasBound, ACubeBound);
    p1 := acadCoordTo(Bound.Max[0], Bound.Max[1], ACanvasBound, ACubeBound);
    ACanvas.Rectangle(p0.X, p0.Y, p1.X, p1.Y);
  end;{if}
end;{DrawToCanvas}
//Определение контура
procedure TacadPoint._DefineBound();
var AHalfSize: Double;
begin
  AHalfSize := 0.5 * Size;
  FBound := acadBound(acadCoord3D(Center[0]-AHalfSize, Center[1]-AHalfSize, Center[2]-AHalfSize),
                         acadCoord3D(Center[0]+AHalfSize, Center[1]+AHalfSize, Center[2]+AHalfSize));
end;{_DefineBound}

//Пользовательский объект CustomPolyline ---------------------------------------
function TacadCustomPolyline.GetCoord(const AIndex: Integer): TacadCoord3D;
begin
  if (AIndex < 0) or (AIndex >= CoordsCount)
  then raise TacadException.InvalidIndex(AIndex,CoordsCount-1);
  Result := FCoords[AIndex];
end;{GetCoord}
function TacadCustomPolyline._GetCoord(const AIndex: Integer): TacadCoord3D;
begin
  Result := FCoords[AIndex];
end;{_GetCoord}
//Конструктор/деструктор
constructor TacadCustomPolyline.Create();
begin
  inherited;
  FPen         := acadPen(acadColor(255,255,255), 1.0, apsSolid);
  FCoordsCount := 0;
  FCoords      := nil;
  FPerimeter   := 0.0;
end;{Create}
destructor TacadCustomPolyline.Destroy();
begin
  FCoordsCount := 0;
  FCoords      := nil;
  inherited;
end;{Destroy}
//Прорисовка на канвасе
procedure TacadCustomPolyline._DrawPlineToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound; const AClosed: Boolean = False);
var
  I    : Integer;
  p0,p1: TPoint;
begin
  if (CoordsCount > 0) and (Pen.Style <> apsClear) then
  begin
    ACanvas.Brush.Color := RGB(Pen.Color.R,Pen.Color.G,Pen.Color.B);
    ACanvas.Pen.Color   := ACanvas.Brush.Color;
    ACanvas.Pen.Width   := Max(1, Round(Pen.Width));
    ACanvas.Pen.Style   := TPenStyle(Pen.Style);
    p0 := acadCoordTo(_Coords[0][0], _Coords[0][1], ACanvasBound, ACubeBound);
    ACanvas.MoveTo(p0.X, p0.Y);
    for I := 1 to CoordsCount-1 do
    begin
      p1 := acadCoordTo(_Coords[I][0], _Coords[I][1], ACanvasBound, ACubeBound);
      ACanvas.LineTo(p1.X, p1.Y);
    end;{for}
    if AClosed then ACanvas.LineTo(p0.X, p0.Y);
  end;{if}
end;{_DrawPlineToCanvas}
//Прорисовка на канвасе
procedure TacadCustomPolyline.DrawToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound);
begin
  if Visible
  then _DrawPlineToCanvas(ACanvas, ACanvasBound, ACubeBound, False);
end;{DrawPlineToCanvas}
//Определение контура и центра
procedure TacadCustomPolyline._DefineBound();
var I, K: Integer;
begin
  FBound     := acadBound();
  FCenter    := acadCoord3D();
  FPerimeter := 0.0;
  if CoordsCount > 0 then
  begin
    //Bound & Perimeter
    FBound.Min := _Coords[0];
    FBound.Max := _Coords[0];
    for I := 1 to CoordsCount-1 do
    begin
      for K := 0 to 2 do
      begin
        if FBound.Min[K] > _Coords[I][K] then FBound.Min[K] := _Coords[I][K];
        if FBound.Max[K] < _Coords[I][K] then FBound.Max[K] := _Coords[I][K];
      end;{for}
      FPerimeter := FPerimeter + Sqrt(Sqr(_Coords[I][0]-_Coords[I-1][0]) + Sqr(_Coords[I][1]-_Coords[I-1][1]));
    end;{for}
    //Center
    for K := 0 to 2 do
      FCenter[K] := (FBound.Min[K] + FBound.Max[K])*0.5;
  end;{if}
end;{_DefineBound}
//Расчет точек контура
procedure TacadCustomPolyline._DefineAdditional();
begin
end;{_DefineAdditional}

//Блок AutoCAD -----------------------------------------------------------------
function TacadBlock.GetItem(const AIndex: Integer): TacadLayerObject;
begin
  if (AIndex < 0) or (AIndex >= Count)
  then raise TacadException.InvalidIndex(AIndex,Count-1);
  Result := TacadLayerObject(FItems.List^[AIndex]);
end;{GetItem}
function TacadBlock._GetItem(const AIndex: Integer): TacadLayerObject;
begin
  Result := TacadLayerObject(FItems.List^[AIndex]);
end;{_GetItem}
//Прорисовка на канвасе
procedure TacadBlock.DrawToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound);
var I: Integer;
begin
  if Visible then
  for I := 0 to Count-1 do
    _Items[I].DrawToCanvas(ACanvas, ACanvasBound, ACubeBound);
end;{DrawToCanvas}
//Очистка
procedure TacadBlock.Clear();
var I: Integer;
begin
  if Count > 0 then
  begin
    FCount := 0;
    for I := FItems.Count-1 downto 0 do
      _Items[I].Free;
    FreeAndNil(FItems);
    _DefineBound();
  end;{if}
end;{Clear}
//Добавление объекта
procedure TacadBlock.Add(const AObject: TacadLayerObject);
begin
  if Assigned(AObject) then
  begin
    //Add ----------------------------------------------------------------------
    if Count = 0 then FItems := TList.Create;
    FItems.Add(AObject);
    FCount := FItems.Count;
    //Update bound/center ------------------------------------------------------
    _DefineBound();
  end;{if}
end;{Add}
//Определение контура
procedure TacadBlock._DefineBound();
var I, K: Integer;
begin
  FBound := acadBound();
  if Count > 0 then
  begin
    //Bound
    FBound  := _Items[0].Bound;
    for I := 1 to Count-1 do
    begin
      for K := 0 to 2 do
      begin
        if FBound.Min[K] > _Items[I].Bound.Min[K] then FBound.Min[K] := _Items[I].Bound.Min[K];
        if FBound.Max[K] < _Items[I].Bound.Max[K] then FBound.Max[K] := _Items[I].Bound.Max[K];
      end;{for}
    end;{for}
  end;{if}
end;{_DefineBound}
//Обработка блока AutoCAD
procedure TacadBlock.ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock);
var AAcadBlockReference: IAcadBlockReference;
begin
  AAcadBlockReference := AAcadObj as IAcadBlockReference;
  //Общие
  FEntityType         := AAcadBlockReference.EntityType;
  FCenter             := acadCoord3D(AAcadBlockReference.InsertionPoint[0], AAcadBlockReference.InsertionPoint[1], AAcadBlockReference.InsertionPoint[2]);
  FVisible            := AAcadBlockReference.Visible = True;
  //Block
  FDefaultPen         := acadPen(acadColor(AAcadBlockReference.TrueColor.Red, AAcadBlockReference.TrueColor.Green, AAcadBlockReference.TrueColor.Blue),
                                    acadLineSizeTo(AAcadBlockReference.Lineweight, ALayer.DefaultPen.Width, 0.0),
                                    acadLineTypeTo(AAcadBlockReference.Linetype, ALayer.DefaultPen.Style, apsClear));
  FName               := AAcadBlockReference.Name;
end;{ExtractAcadEntity}
//Обработка объектов блока AutoCAD
procedure TacadBlock.ExtractAcadEntityDn(const ALayer: TacadLayer; const AAcadBlocks: IAcadBlocks);
var
  AAcadBlock   : IAcadBlock;           //Текущий блок AutoCAD
  AAcadBlockObj: IAcadEntity;          //Объект блока AutoCAD
  ABlockName   : String;               //Имя блока AutoCAD
  ABlockObject : TacadLayerObject;     //Пользовательский объект блока
  AClassRef    : TacadLayerObjectRef;  //Ссылка на тип класса TacadLayerObject
  //
  I, K         : Integer;
begin
  //Пробегаемся по блокам текущего документа
  for I := 0 to AAcadBlocks.Count-1 do
  begin
    ABlockName := AAcadBlocks.Item(I).Name;
    //Проверка на пользовательские блоки
    if Name = ABlockName then
    begin
      AAcadBlock := AAcadBlocks.Item(I);
      //Объекты текущего блока
      for K := 0 to AAcadBlock.Count-1 do
      begin
        AAcadBlockObj := AAcadBlock.Item(K);
        case AAcadBlockObj.EntityType of
          acPoint          : AClassRef := TacadPoint;
          acLine           : AClassRef := TacadLine;
          acCircle         : AClassRef := TacadCircle;
          acArc            : AClassRef := TacadArc;
          acSpline         : AClassRef := TacadSpline;
          acEllipse        : AClassRef := TacadEllipse;
          acPolylineLight  : AClassRef := TacadPolyline;
          ac3dPolyline     : AClassRef := Tacad3DPolyline;
          acBlockReference : AClassRef := TacadBlock;
          acHatch          : AClassRef := TacadPolygon;
          //acText           : AClassRef := TacadText;
          //acMText          : AClassRef := TacadMText;
          else AClassRef := nil;
        end;{case}
        if AClassRef <> nil then
        begin
          ABlockObject := AClassRef.Create();
          ABlockObject.ExtractAcadEntity(AAcadBlockObj, ALayer, Self);
          //Обработка вложенного блока рекурсией
          if ABlockObject is TacadBlock
          then TacadBlock(ABlockObject).ExtractAcadEntityDn(ALayer, AAcadBlocks);
          Add(ABlockObject);
        end;{if}
      end;{for}
      Break;
    end;{if}
  end;{for}
end;{ExtractAcadEntityDn}
//Конструктор/деструктор
constructor TacadBlock.Create();
begin
  inherited;
  FDefaultPen := acadPen(acadColor(255,255,255), 0.1, apsSolid);
  FCount      := 0;
  FItems      := nil;
  FName       := '';
end;{Create}
destructor TacadBlock.Destroy();
begin
  Clear();
  inherited;
end;{Destroy}

//Объект AutoCAD "Линия" -------------------------------------------------------
function TacadLine.GetStartPoint(): TacadCoord3D;
begin
  Result := Coords[0];
end;{GetStartPoint}
function TacadLine.GetEndPoint(): TacadCoord3D;
begin
  Result := Coords[1];
end;{GetEndPoint}
//Извеление линии AutoCAD
procedure TacadLine.ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock);
var
  AAcadLine     : IAcadLine;
  ABlockPenWidth: Double;
  ABlockCenter  : TacadCoord3D;
  ABlockPenStyle: TacadPenStyle;
begin
  AAcadLine      := AAcadObj as IAcadLine;
  //Block
  ABlockPenWidth := 0.0;
  ABlockCenter   := acadCoord3D();
  ABlockPenStyle := apsClear;
  if ABlock <> nil then
  begin
    ABlockPenWidth := ABlock.DefaultPen.Width;
    ABlockCenter   := acadCoord3D(ABlock.Center[0], ABlock.Center[1], ABlock.Center[2]);
    ABlockPenStyle := ABlock.DefaultPen.Style;
  end;{if}
  AAcadLine     := AAcadObj as IAcadLine;
  //Общие
  FEntityType   := AAcadLine.EntityType;
  //Графический объект
  FVisible      := AAcadLine.Visible = True;
  //CustomPolyline
  FPen          := acadPen(acadColor(AAcadLine.TrueColor.Red, AAcadLine.TrueColor.Green, AAcadLine.TrueColor.Blue), acadLineSizeTo(AAcadLine.Lineweight, ALayer.DefaultPen.Width, ABlockPenWidth),
                              acadLineTypeTo(AAcadLine.Linetype, ALayer.DefaultPen.Style, ABlockPenStyle));
  FCoordsCount  := 2;
  //Line
  SetLength(FCoords, CoordsCount);
  FCoords[0]    := acadCoord3D(AAcadLine.StartPoint[0]+ABlockCenter[0], AAcadLine.StartPoint[1]+ABlockCenter[1], AAcadLine.StartPoint[2]+ABlockCenter[2]);
  FCoords[1]    := acadCoord3D(AAcadLine.EndPoint[0]+ABlockCenter[0], AAcadLine.EndPoint[1]+ABlockCenter[1], AAcadLine.EndPoint[2]+ABlockCenter[2]);
  //Bound
  _DefineBound();
end;{ExtractAcadEntity}

//Объект AutoCAD "Полилиния" ---------------------------------------------------
//Конструктор/деструктор
constructor TacadPolyline.Create();
begin
  inherited;
  FClosed := False;
end;{Create}
//Обработка полилинии AutoCAD
procedure TacadPolyline.ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock);
var
  AAcadPolyline : IAcadLWPolyline;
  ABlockPenWidth: Double;
  ABlockCenter  : TacadCoord3D;
  ABlockPenStyle: TacadPenStyle;
  I, ACount     : Integer;
begin
  AAcadPolyline  := AAcadObj as IAcadLWPolyline;
  //Block
  ABlockPenWidth := 0.0;
  ABlockCenter   := acadCoord3D();
  ABlockPenStyle := apsClear;
  if ABlock <> nil then
  begin
    ABlockPenWidth := ABlock.DefaultPen.Width;
    ABlockCenter   := acadCoord3D(ABlock.Center[0], ABlock.Center[1], ABlock.Center[2]);
    ABlockPenStyle := ABlock.DefaultPen.Style;
  end;{if}
  AAcadPolyline := AAcadObj as IAcadLWPolyline;
  //Общие
  FEntityType   := AAcadPolyline.EntityType;
  //Графический объект
  FVisible      := AAcadPolyline.Visible = True;
  //CustomPolyline
  FPen          := acadPen(acadColor(AAcadPolyline.TrueColor.Red, AAcadPolyline.TrueColor.Green, AAcadPolyline.TrueColor.Blue), acadLineSizeTo(AAcadPolyline.Lineweight, ALayer.DefaultPen.Width, ABlockPenWidth),
                              acadLineTypeTo(AAcadPolyline.Linetype, ALayer.DefaultPen.Style, ABlockPenStyle));
  ACount        := (VarArrayHighBound(AAcadPolyline.Coordinates, 1) - VarArrayLowBound(AAcadPolyline.Coordinates, 1) + 1) div 2;
  SetLength(FCoords, ACount);
  for I := 0 to ACount-1 do
  begin
    FCoords[I][0] := AAcadPolyline.Coordinates[2*I]+ABlockCenter[0];
    FCoords[I][1] := AAcadPolyline.Coordinates[2*I+1]+ABlockCenter[1];
  end;{for}
  FCoordsCount := ACount;
  //Polyline
  FClosed      := AAcadPolyline.Closed;
  //Bound & Perimeter
  _DefineBound();
end;{ExtractAcadEntity}
//Прорисовка на канвасе
procedure TacadPolyline.DrawToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound);
begin
  if Visible
  then _DrawPlineToCanvas(ACanvas, ACanvasBound, ACubeBound, Closed);
end;{DrawToCanvas}
//Определение контура и центра
procedure TacadPolyline._DefineBound();
var I: Integer;
begin
  inherited;
  //Perimeter
  FPerimeter := 0.0;
  if CoordsCount > 0 then
  begin
    for I := 1 to CoordsCount-1 do
      FPerimeter := FPerimeter + Sqrt(Sqr(_Coords[I][0]-_Coords[I-1][0]) + Sqr(_Coords[I][1]-_Coords[I-1][1]));
    if Closed
    then FPerimeter := Perimeter + Sqrt(Sqr(_Coords[0][0]-_Coords[CoordsCount-1][0]) + Sqr(_Coords[0][1]-_Coords[CoordsCount-1][1]));
  end;{if}
end;{_DefineBound}

//Объект AutoCAD "3D Полилиния" ------------------------------------------------
//Конструктор/деструктор
constructor Tacad3DPolyline.Create();
begin
  inherited;
  FClosed := False;
end;{Create}
//Обработка полилинии AutoCAD
procedure Tacad3DPolyline.ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock);
var
  AAcad3DPolyline: IAcad3DPolyline;
  ABlockPenWidth : Double;
  ABlockCenter   : TacadCoord3D;
  ABlockPenStyle : TacadPenStyle;
  I, ACount      : Integer;
begin
  AAcad3DPolyline := AAcadObj as IAcad3DPolyline;
  //Block
  ABlockPenWidth  := 0.0;
  ABlockCenter    := acadCoord3D();
  ABlockPenStyle  := apsClear;
  if ABlock <> nil then
  begin
    ABlockPenWidth := ABlock.DefaultPen.Width;
    ABlockCenter   := acadCoord3D(ABlock.Center[0], ABlock.Center[1], ABlock.Center[2]);
    ABlockPenStyle := ABlock.DefaultPen.Style;
  end;{if}
  //Общие
  FEntityType     := AAcad3DPolyline.EntityType;
  //Графический объект
  FVisible        := AAcad3DPolyline.Visible = True;
  //CustomPolyline
  FPen := acadPen(acadColor(AAcad3DPolyline.TrueColor.Red, AAcad3DPolyline.TrueColor.Green, AAcad3DPolyline.TrueColor.Blue), acadLineSizeTo(AAcad3DPolyline.Lineweight, ALayer.DefaultPen.Width, ABlockPenWidth),
                     acadLineTypeTo(AAcad3DPolyline.Linetype, ALayer.DefaultPen.Style, ABlockPenStyle));
  ACount := (VarArrayHighBound(AAcad3DPolyline.Coordinates, 1) - VarArrayLowBound(AAcad3DPolyline.Coordinates, 1) + 1) div 3;
  SetLength(FCoords, ACount);
  for I := 0 to ACount-1 do
  begin
    FCoords[I][0] := AAcad3DPolyline.Coordinates[3*I]+ABlockCenter[0];
    FCoords[I][1] := AAcad3DPolyline.Coordinates[3*I+1]+ABlockCenter[1];
    FCoords[I][2] := AAcad3DPolyline.Coordinates[3*I+2]+ABlockCenter[2];
  end;{for}
  FCoordsCount := ACount;
  //Polyline
  FClosed := AAcad3DPolyline.Closed;
  //Bound & Perimeter
  _DefineBound();
end;{ExtractAcadEntity}
//Прорисовка на канвасе
procedure Tacad3DPolyline.DrawToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound);
begin
  if Visible
  then _DrawPlineToCanvas(ACanvas, ACanvasBound, ACubeBound, Closed);
end;{DrawToCanvas}
//Определение контура и центра
procedure Tacad3DPolyline._DefineBound();
var I: Integer;
begin
  inherited;
  //Perimeter
  FPerimeter := 0.0;
  if CoordsCount > 0 then
  begin
    for I := 1 to CoordsCount-1 do
      FPerimeter := FPerimeter + Sqrt(Sqr(_Coords[I][0]-_Coords[I-1][0]) + Sqr(_Coords[I][1]-_Coords[I-1][1]));
    if Closed
    then FPerimeter := Perimeter + Sqrt(Sqr(_Coords[0][0]-_Coords[CoordsCount-1][0]) + Sqr(_Coords[0][1]-_Coords[CoordsCount-1][1]));
  end;{if}
end;{_DefineBound}

//Объект AutoCAD "Дуга" --------------------------------------------------------
//Конструктор/деструктор
constructor TacadArc.Create();
begin
  inherited;
  FRadius     := 1.0;
  FStartPoint := acadCoord3D();
  FEndPoint   := acadCoord3D();
  FStartAngle := 0.0;
  FEndAngle   := 0.0;
end;{Create}
//Обработка дуги AutoCAD
procedure TacadArc.ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock);
var
  AAcadArc      : IAcadArc;
  ABlockPenWidth: Double;
  ABlockCenter  : TacadCoord3D;
  ABlockPenStyle: TacadPenStyle;
begin
  AAcadArc        := AAcadObj as IAcadArc;
  //Block
  ABlockPenWidth  := 0.0;
  ABlockCenter    := acadCoord3D();
  ABlockPenStyle  := apsClear;
  if ABlock <> nil then
  begin
    ABlockPenWidth := ABlock.DefaultPen.Width;
    ABlockCenter   := acadCoord3D(ABlock.Center[0], ABlock.Center[1], ABlock.Center[2]);
    ABlockPenStyle := ABlock.DefaultPen.Style;
  end;{if}
  //Общие
  FEntityType  := AAcadArc.EntityType;
  //Графический объект
  FCenter      := acadCoord3D(AAcadArc.Center[0]+ABlockCenter[0], AAcadArc.Center[1]+ABlockCenter[1], AAcadArc.Center[2]+ABlockCenter[2]);
  FVisible     := AAcadArc.Visible = True;
  //CustomPolyline
  FPen          := acadPen(acadColor(AAcadArc.TrueColor.Red, AAcadArc.TrueColor.Green, AAcadArc.TrueColor.Blue), acadLineSizeTo(AAcadArc.Lineweight, ALayer.DefaultPen.Width, ABlockPenWidth),
                              acadLineTypeTo(AAcadArc.Linetype, ALayer.DefaultPen.Style, ABlockPenStyle));
  FStartPoint  := acadCoord3D(AAcadArc.StartPoint[0]+ABlockCenter[0], AAcadArc.StartPoint[1]+ABlockCenter[1], AAcadArc.StartPoint[2]+ABlockCenter[2]);
  FEndPoint    := acadCoord3D(AAcadArc.EndPoint[0]+ABlockCenter[0], AAcadArc.EndPoint[1]+ABlockCenter[1], AAcadArc.EndPoint[2]+ABlockCenter[2]);
  //Arc
  FRadius      := AAcadArc.Radius;
  FStartAngle  := AAcadArc.StartAngle;
  FEndAngle    := AAcadArc.EndAngle;
  _DefineAdditional();
  //Bound
  _DefineBound();
end;{ExtractAcadEntity}
//Расчет точек контура
procedure TacadArc._DefineAdditional();
var I, K, ACount: Integer;
const
  RAD_TO_DEG = 180.0/PI;
  DEG_TO_RAD = PI/180.0;
begin
  K := 1;
  if (StartAngle < EndAngle)
  then ACount := 2 + Trunc(abs((StartAngle*RAD_TO_DEG) - (EndAngle*RAD_TO_DEG)))
  else ACount := 2 + Trunc((2*PI*RAD_TO_DEG) - (StartAngle*RAD_TO_DEG) + (EndAngle*RAD_TO_DEG));
  SetLength(FCoords, ACount);
  FCoords[0]        := StartPoint;
  FCoords[ACount-1] := EndPoint;
  for I := 1 to ACount-1 do
  begin
    FCoords[K] := acadCoord3D(Center[0]+Radius*Cos(StartAngle+(I*DEG_TO_RAD)), Center[1]+Radius*Sin(StartAngle+(I*DEG_TO_RAD)), 0.0);
    Inc(K);
  end;{for}
  FCoordsCount := ACount;
end;{_DefineAdditional}

//Объект AutoCAD "Сплайн" ------------------------------------------------------
function TacadSpline.GetNode(const AIndex: Integer): TacadCoord3D;
begin
  if (AIndex < 0) or (AIndex >= NodesCount)
  then raise TacadException.InvalidIndex(AIndex,NodesCount-1);
  Result := FNodes[AIndex];
end;{GetNode}
function TacadSpline._GetNode(const AIndex: Integer): TacadCoord3D;
begin
  Result := FNodes[AIndex];
end;{_GetNode}
//Конструктор/деструктор
constructor TacadSpline.Create();
begin
  inherited;
  FClosed     := False;
  FKind       := askQuadratic;
  FNodesCount := 0;
  FNodes      := nil;
end;{Create}
destructor TacadSpline.Destroy();
begin
  FNodesCount := 0;
  FNodes      := nil;
  inherited;
end;{Destroy}
//Обработка сплайна AutoCAD
procedure TacadSpline.ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock);
var
  AAcadSpline   : IAcadSpline;
  ABlockPenWidth: Double;
  ABlockCenter  : TacadCoord3D;
  ABlockPenStyle: TacadPenStyle;
  I, ACount     : Integer;
begin
  AAcadSpline     := AAcadObj as IAcadSpline;
  //Block
  ABlockPenWidth  := 0.0;
  ABlockCenter    := acadCoord3D();
  ABlockPenStyle  := apsClear;
  if ABlock <> nil then
  begin
    ABlockPenWidth := ABlock.DefaultPen.Width;
    ABlockCenter   := acadCoord3D(ABlock.Center[0], ABlock.Center[1], ABlock.Center[2]);
    ABlockPenStyle := ABlock.DefaultPen.Style;
  end;{if}
  //Общие
  FEntityType := AAcadSpline.EntityType;
  //Графический объект
  FVisible    := AAcadSpline.Visible = True;
  //CustomPolyline
  FPen        := acadPen(acadColor(AAcadSpline.TrueColor.Red, AAcadSpline.TrueColor.Green, AAcadSpline.TrueColor.Blue), acadLineSizeTo(AAcadSpline.Lineweight, ALayer.DefaultPen.Width, ABlockPenWidth),
                            acadLineTypeTo(AAcadSpline.Linetype, ALayer.DefaultPen.Style, ABlockPenStyle));
  //Spline
  FClosed     := AAcadSpline.Closed;
  //ACount := (VarArrayHighBound(AAcadSpline.ControlPoints, 1) - VarArrayLowBound(AAcadSpline.ControlPoints, 1) + 1) div 3;
  ACount := AAcadSpline.NumberOfControlPoints;
  SetLength(FNodes, ACount);
  for I := 0 to ACount-1 do
  begin
    FNodes[I][0] := AAcadSpline.ControlPoints[3*I]+ABlockCenter[0];
    FNodes[I][1] := AAcadSpline.ControlPoints[3*I+1]+ABlockCenter[1];
    FNodes[I][2] := AAcadSpline.ControlPoints[3*I+2]+ABlockCenter[2];
  end;{for}
  FNodesCount := ACount;
  //FKind := ;
  _DefineAdditional();
  //Bound
  _DefineBound();
end;{ExtractAcadEntity}
//Расчет точек контура
procedure TacadSpline._DefineAdditional();
var I, ACount: Integer;
begin
  ACount := NodesCount;
  SetLength(FCoords, ACount);
  for I := 0 to ACount-1 do
    FCoords[I] := _Nodes[I];
  FCoordsCount := ACount;
end;{_DefineAdditional}

//Пользовательский объект CustomPolygon ----------------------------------------
//Конструктор/деструктор
constructor TacadCustomPolygon.Create();
begin
  inherited;
  FBrush          := acadBrush(acadColor(255,255,255), absSolid);
  FArea           := 0.0;
  FTrianglesCount := 0;
  FTriangles      := nil;
end;{Create}
destructor TacadCustomPolygon.Destroy();
begin
  FTrianglesCount := 0;
  FTriangles      := nil;
  inherited;
end;{Destroy}
//Методы диспетчеризации свойств
function TacadCustomPolygon.GetTriangle(const AIndex: Integer): RacadTriangle;
begin
  if (AIndex < 0) or (AIndex >= TrianglesCount)
  then raise TacadException.InvalidIndex(AIndex, TrianglesCount-1);
  Result := FTriangles[AIndex];
end;{GetTriangle}
function TacadCustomPolygon._GetTriangle(const AIndex: Integer): RacadTriangle;
begin
  Result := FTriangles[AIndex];
end;{_GetTriangle}
//Прорисовка на канвасе
procedure TacadCustomPolygon._DrawPgonToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound);
var
  p0,p1,p2: TPoint;
  I       : Integer;
begin
  if (TrianglesCount > 0) and (Brush.Style <> absClear) then
  begin
    //Прорисовка треугольников
    ACanvas.Brush.Color := RGB(Pen.Color.R,Pen.Color.G,Pen.Color.B);
    ACanvas.Pen.Color   := ACanvas.Brush.Color;
    for I := 0 to TrianglesCount-1 do
    begin
      p0 := acadCoordTo(Triangles[I].p0[0], Triangles[I].p0[1], ACanvasBound, ACubeBound);
      p1 := acadCoordTo(Triangles[I].p1[0], Triangles[I].p1[1], ACanvasBound, ACubeBound);
      p2 := acadCoordTo(Triangles[I].p2[0], Triangles[I].p2[1], ACanvasBound, ACubeBound);
      ACanvas.Polygon([p0, p1, p2]);
    end;{for}
  end;{if}
end;{_DrawPgonToCanvas}
procedure TacadCustomPolygon.DrawToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound);
begin
  if Visible then
  begin
    _DrawPgonToCanvas(ACanvas, ACanvasBound, ACubeBound);
    _DrawPlineToCanvas(ACanvas, ACanvasBound, ACubeBound, True);
  end;{if}
end;{DrawToCanvas}
//Определение контура и центра
procedure TacadCustomPolygon._DefineBound();
var I: Integer;
begin
  inherited;
  //Расчет суммы площадей треугольников
  FArea := 0.0;
  if TrianglesCount > 0 then
  begin
    for I := 0 to TrianglesCount-1 do
      FArea := Area + ( (_Triangles[I].p0[0]-_Triangles[I].p2[0])*(_Triangles[I].p1[1]-_Triangles[I].p2[1])-(_Triangles[I].p1[0]-_Triangles[I].p2[0])*(_Triangles[I].p0[1]-_Triangles[I].p2[1]) ) * 0.5;
  end;{if}
end;{_DefineBound}
//Расчет точек контура и треугольников
procedure TacadCustomPolygon._DefineAdditional();
begin
  FTrianglesCount := 0;
  FTriangles      := nil;
end;{_DefineAdditional}

//Объект AutoCAD "Полигон" -----------------------------------------------------
//Обработка полигона AutoCAD
procedure TacadPolygon.ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock);
var
  AAcadHatch    : IAcadHatch;
  ABlockPenWidth: Double;
  ABlockCenter  : TacadCoord3D;
  ABlockPenStyle: TacadPenStyle;
begin
  AAcadHatch      := AAcadObj as IAcadHatch;
  //Block
  ABlockPenWidth  := 0.0;
  ABlockCenter    := acadCoord3D();
  ABlockPenStyle  := apsClear;
  if ABlock <> nil then
  begin
    ABlockPenWidth := ABlock.DefaultPen.Width;
    ABlockCenter   := acadCoord3D(ABlock.Center[0], ABlock.Center[1], ABlock.Center[2]);
    ABlockPenStyle := ABlock.DefaultPen.Style;
  end;{if}
  //Общие
  FEntityType  := AAcadHatch.EntityType;
  //Графический объект
  FCenter      := acadCoord3D(AAcadHatch.Origin[0]+ABlockCenter[0], AAcadHatch.Origin[1]+ABlockCenter[1], 0.0);
  FVisible     := AAcadHatch.Visible = True;
  //CustomPolyline
  FPen         := acadPen(acadColor(AAcadHatch.TrueColor.Red, AAcadHatch.TrueColor.Green, AAcadHatch.TrueColor.Blue), acadLineSizeTo(AAcadHatch.Lineweight, ALayer.DefaultPen.Width, ABlockPenWidth),
                             acadLineTypeTo(AAcadHatch.Linetype, ALayer.DefaultPen.Style, ABlockPenStyle));
  //CustomPolygon
  FBrush       := acadBrush(acadColor(AAcadHatch.TrueColor.Red, AAcadHatch.TrueColor.Green, AAcadHatch.TrueColor.Blue), acadBrushStyleTo(AAcadHatch.PatternName));
  //Polygon
  _DefineAdditional();
  //Bound
  _DefineBound();
end;{ExtractAcadEntity}
//Расчет точек контура и треугольников
procedure TacadPolygon._DefineAdditional();
begin
end;{_DefineAdditional}

//Объект AutoCAD "Круг" --------------------------------------------------------
//Конструктор/деструктор
constructor TacadCircle.Create();
begin
  inherited;
  FRadius := 1.0;
end;{Create}
//Обработка круга AutoCAD
procedure TacadCircle.ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock);
var
  AAcadCircle   : IAcadCircle;
  ABlockPenWidth: Double;
  ABlockCenter  : TacadCoord3D;
  ABlockPenStyle: TacadPenStyle;
begin
  AAcadCircle     := AAcadObj as IAcadCircle;
  //Block
  ABlockPenWidth  := 0.0;
  ABlockCenter    := acadCoord3D();
  ABlockPenStyle  := apsClear;
  if ABlock <> nil then
  begin
    ABlockPenWidth := ABlock.DefaultPen.Width;
    ABlockCenter   := acadCoord3D(ABlock.Center[0], ABlock.Center[1], ABlock.Center[2]);
    ABlockPenStyle := ABlock.DefaultPen.Style;
  end;{if}
  //Общие
  FEntityType  := AAcadCircle.EntityType;
  //Графический объект
  FCenter      := acadCoord3D(AAcadCircle.Center[0]+ABlockCenter[0], AAcadCircle.Center[1]+ABlockCenter[1], AAcadCircle.Center[2]+ABlockCenter[2]);
  FVisible     := AAcadCircle.Visible = True;
  //CustomPolyline
  FPen         := acadPen(acadColor(AAcadCircle.TrueColor.Red, AAcadCircle.TrueColor.Green, AAcadCircle.TrueColor.Blue), acadLineSizeTo(AAcadCircle.Lineweight, ALayer.DefaultPen.Width, ABlockPenWidth),
                             acadLineTypeTo(AAcadCircle.Linetype, ALayer.DefaultPen.Style, ABlockPenStyle));
  //CustomPolygon
  FBrush       := acadBrush(acadColor(AAcadCircle.TrueColor.Red, AAcadCircle.TrueColor.Green, AAcadCircle.TrueColor.Blue), absSolid);
  //Circle
  FRadius      := AAcadCircle.Radius;
  _DefineAdditional();
  //Bound
  _DefineBound();
end;{ExtractAcadEntity}
//Расчет точек контура и треугольников
procedure TacadCircle._DefineAdditional();
var I, ACount: Integer;
begin
  ACount := 36;
  SetLength(FCoords, ACount);
  for I := 0 to ACount-1 do
    FCoords[I] := acadCoord3D(Center[0]+Radius*Cos(I*10.0*PI/180.0), Center[1]+Radius*Sin(I*10.0*PI/180.0), 0.0);
  FCoordsCount := ACount;
  //Расчет координат треугольников
  //Количество треугольников
  ACount := 36;
  SetLength(FTriangles, ACount);
  FTriangles[0].p0[0] := _Coords[ACount-1][0];
  FTriangles[0].p0[1] := _Coords[ACount-1][1];
  FTriangles[0].p1[0] := Center[0];
  FTriangles[0].p1[1] := Center[1];
  FTriangles[0].p2[0] := _Coords[0][0];
  FTriangles[0].p2[1] := _Coords[0][1];
  for I := 1 to ACount-1 do
  begin
    FTriangles[I].p0[0] := _Coords[I][0];
    FTriangles[I].p0[1] := _Coords[I][1];
    FTriangles[I].p1[0] := Center[0];
    FTriangles[I].p1[1] := Center[1];
    FTriangles[I].p2[0] := _Coords[I-1][0];
    FTriangles[I].p2[1] := _Coords[I-1][1];
  end;{for}
  FTrianglesCount := ACount;
end;{_DefineAdditional}

//Объект AutoCAD "Эллипс" ------------------------------------------------------
//Конструктор/деструктор
constructor TacadEllipse.Create();
begin
  inherited;
  FMajorRadius := 0.0;
  FMinorRadius := 0.0;
  FMajorAxis   := 0.0;
end;{Create}
//Обработка эллипса AutoCAD
procedure TacadEllipse.ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock);
var
  AAcadEllipse  : IAcadEllipse;
  ABlockPenWidth: Double;
  ABlockCenter  : TacadCoord3D;
  ABlockPenStyle: TacadPenStyle;
begin
  AAcadEllipse    := AAcadObj as IAcadEllipse;
  //Block
  ABlockPenWidth  := 0.0;
  ABlockCenter    := acadCoord3D();
  ABlockPenStyle  := apsClear;
  if ABlock <> nil then
  begin
    ABlockPenWidth := ABlock.DefaultPen.Width;
    ABlockCenter   := acadCoord3D(ABlock.Center[0], ABlock.Center[1], ABlock.Center[2]);
    ABlockPenStyle := ABlock.DefaultPen.Style;
  end;{if}
  //Общие
  FEntityType  := AAcadEllipse.EntityType;
  //Графический объект
  FCenter      := acadCoord3D(AAcadEllipse.Center[0]+ABlockCenter[0], AAcadEllipse.Center[1]+ABlockCenter[1], AAcadEllipse.Center[2]+ABlockCenter[2]);
  FBound       := acadBound(FCenter, FCenter);
  FVisible     := AAcadEllipse.Visible = True;
  //CustomPolyline
  FPen         := acadPen(acadColor(AAcadEllipse.TrueColor.Red, AAcadEllipse.TrueColor.Green, AAcadEllipse.TrueColor.Blue), acadLineSizeTo(AAcadEllipse.Lineweight, ALayer.DefaultPen.Width, ABlockPenWidth),
                             acadLineTypeTo(AAcadEllipse.Linetype, ALayer.DefaultPen.Style, ABlockPenStyle));
  //CustomPolygon
  FBrush       := acadBrush(acadColor(AAcadEllipse.TrueColor.Red, AAcadEllipse.TrueColor.Green, AAcadEllipse.TrueColor.Blue), absSolid);
  FArea        := AAcadEllipse.Area;
  //Ellipse
  FMajorRadius := AAcadEllipse.MajorRadius;
  FMinorRadius := AAcadEllipse.MinorRadius;
  FMajorAxis   := AAcadEllipse.MajorAxis[0];
  _DefineAdditional();
  //Bound
  _DefineBound();
end;{ExtractAcadEntity}
//Расчет точек контура и треугольников
procedure TacadEllipse._DefineAdditional();
var
  AMajorRadius, AMinorRadius: Double;
  I, ACount                 : Integer;
begin
  //Расчет координат контура
  AMajorRadius := MajorRadius;
  AMinorRadius := MinorRadius;
  ACount := 36;
  SetLength(FCoords, ACount);
  if MajorAxis = 0 then
  begin
    AMajorRadius := MinorRadius;
    AMinorRadius := MajorRadius;
  end;{if}
  for I := 0 to ACount-1 do
    FCoords[I] := acadCoord3D(Center[0]+AMajorRadius*Cos(I*10.0*PI/180.0), Center[1]+AMinorRadius*Sin(I*10.0*PI/180.0), 0.0);
  FCoordsCount := ACount;
  //Расчет координат треугольников
  //Количество треугольников
  ACount := 36;
  SetLength(FTriangles, ACount);
  FTriangles[0].p0[0] := _Coords[ACount-1][0];
  FTriangles[0].p0[1] := _Coords[ACount-1][1];
  FTriangles[0].p1[0] := Center[0];
  FTriangles[0].p1[1] := Center[1];
  FTriangles[0].p2[0] := _Coords[0][0];
  FTriangles[0].p2[1] := _Coords[0][1];
  for I := 1 to ACount-1 do
  begin
    FTriangles[I].p0[0] := _Coords[I][0];
    FTriangles[I].p0[1] := _Coords[I][1];
    FTriangles[I].p1[0] := Center[0];
    FTriangles[I].p1[1] := Center[1];
    FTriangles[I].p2[0] := _Coords[I-1][0];
    FTriangles[I].p2[1] := _Coords[I-1][1];
  end;{for}
  FTrianglesCount := ACount;
end;{_DefineAdditional}

//Объект AutoCAD "Текст" -------------------------------------------------------
//Конструктор/деструктор
constructor TacadText.Create();
begin
  inherited;
  FFont          := acadFont(acadColor(255,255,255), 2.5, [afsNormal], afnTimesNewRoman);
  FCaption       := '';
  FHAlign        := ahaLeft;
  FVAlign        := avaCenter;
  FRotation      := 0.0;
  FScaleFactor   := 1.0;
  FObliqueAngle  := 15.0;
  FTextDirection := atdHorizontal;
  FUpsideDown    := False;
  FBackward      := False;
end;{Create}
//Обработка текста AutoCAD
procedure TacadText.ExtractAcadEntity(const AAcadObj: IAcadEntity; const ALayer: TacadLayer; const ABlock: TacadBlock);
begin

end;{ExtractAcadEntity}
//Прорисовка на канвасе
procedure TacadText.DrawToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound);
begin
  if Visible then
  begin
    ACanvas.Font.Color := RGB(Font.Color.R, Font.Color.G, Font.Color.B);
    //ACanvas.Font.Name 
    //ACanvas.Font.Style
    //ACanvas.Font.Size
    //ACanvas.Font.Pitch
    //ACanvas.Font.Height

    //Прорисовка...
  end;{if}
end;{DrawToCanvas}

//Слой AutoCAD ---------------------------------------------------------------
function TacadLayer.GetItem(const AIndex: Integer): TacadLayerObject;
begin
  if (AIndex < 0) or (AIndex >= Count)
  then raise TacadException.InvalidIndex(AIndex,Count-1);
  Result := TacadLayerObject(FItems.List^[AIndex]);
end;{GetItem}
function TacadLayer._GetItem(const AIndex: Integer): TacadLayerObject;
begin
  Result := TacadLayerObject(FItems.List^[AIndex]);
end;{_GetItem}
//Прорисовка на канвасе
procedure TacadLayer.DrawToCanvas(const ACanvas: TCanvas; const ACanvasBound: TRect; const ACubeBound: RacadBound);
var I: Integer;
begin
  if Visible and (Count > 0) then
  for I := 0 to Count-1 do
    _Items[I].DrawToCanvas(ACanvas, ACanvasBound, ACubeBound);
end;{DrawToCanvas}
//Очистка
procedure TacadLayer.Clear();
var I: Integer;
begin
  if Count > 0 then
  begin
    FCount := 0;
    for I := FItems.Count-1 downto 0 do
      _Items[I].Free;                             
    FreeAndNil(FItems);
    _DefineBound();
  end;{if}
end;{Clear}
//Добавление объекта
procedure TacadLayer.Add(const AObject: TacadLayerObject);
begin
  if Assigned(AObject) then
  begin
    //Add ----------------------------------------------------------------------
    if Count = 0 then FItems := TList.Create;
    FItems.Add(AObject);
    FCount := FItems.Count;
    //Update bound/center ------------------------------------------------------
    _DefineBound();
  end;{if}
end;{Add}
//Определение контура
procedure TacadLayer._DefineBound();
var I, K: Integer;
begin
  FBound  := acadBound();
  FCenter := acadCoord3D();
  if Count > 0 then
  begin
    //Bound
    FBound  := _Items[0].Bound;
    FCenter := _Items[0].Center;
    for I := 1 to Count-1 do
    begin
      for K := 0 to 2 do
      begin
        if FBound.Min[K] > _Items[I].Bound.Min[K] then FBound.Min[K] := _Items[I].Bound.Min[K];
        if FBound.Max[K] < _Items[I].Bound.Max[K] then FBound.Max[K] := _Items[I].Bound.Max[K];
      end;{for}
    end;{for}
    //Center
    for K := 0 to 2 do
      FCenter[K] := (FBound.Min[K] + FBound.Max[K])*0.5;
  end;{if}
end;{_DefineBound}
//Конструктор/деструктор
constructor TacadLayer.Create();
begin
  inherited;
  FName        := '';
  FVisible     := True;
  FFreeze      := False;
  FLock        := False;
  FDefaultPen  := acadPen(acadColor(255,255,255),1.0,apsSolid);
  FDescription := '';
  FCount       := 0;
  FItems       := nil;
end;{Create}
destructor TacadLayer.Destroy();
begin
  Clear();
  inherited;
end;{Destroy}

//Статистика объектов AutoCAD --------------------------------------------------
//Конструктор/деструктор
constructor TacadEntitiesStatistic.Create();
begin
  inherited;
  FItems := nil;
  FCount := 0;
end;{Create}
destructor TacadEntitiesStatistic.Destroy();
begin
  FItems := nil;
  FCount := 0;
  inherited;
end;{Destroy}
//Методы диспетчеризаций свойств
function TacadEntitiesStatistic._GetItem(const AIndex: Integer): RacadEntityType;
begin                                   
  Result := PacadEntityType(FItems.List^[AIndex])^;
end;{_GetItem}
function TacadEntitiesStatistic.GetItem(const AIndex: Integer): RacadEntityType;
begin                                   
  if (AIndex < 0) or (AIndex >= Count)
  then raise TacadException.InvalidIndex(AIndex,Count-1);
  Result := PacadEntityType(FItems.List^[AIndex])^;
end;{GetItem}
//Добавление
procedure TacadEntitiesStatistic.Add(const AEntityType: Integer; const AEntityName: String; const AUnknown: Boolean = False);
var
  AIndex: Integer;
  AItem : PacadEntityType;
begin
  AIndex := IndexOf(AEntityType);
  if AIndex = -1 then
  begin
    New(AItem);
    AItem^.EntityType := AEntityType;
    AItem^.EntityName := AEntityName;
    AItem^.Count      := 1;
    if AUnknown
    then AItem^.ImportedCount := 0
    else AItem^.ImportedCount := 1;
    if Count = 0 then FItems := TList.Create();
    FItems.Add(AItem);
  end{if}
  else
  begin
    PacadEntityType(FItems.List^[AIndex])^.Count := PacadEntityType(FItems.List^[AIndex])^.Count + 1;
    if not AUnknown
    then PacadEntityType(FItems.List^[AIndex])^.ImportedCount := PacadEntityType(FItems.List^[AIndex])^.ImportedCount + 1;
  end;{else}
  FCount := FItems.Count;
end;{Add}
//Очистка
procedure TacadEntitiesStatistic.Clear();
var I: Integer;
begin
  if Count > 0 then
  begin
    FCount := 0;
    for I := FItems.Count-1 downto 0 do
      Dispose(PacadEntityType(FItems.List^[I]));
    FreeAndNil(FItems);
    FCount := 0;
  end;{if}
end;{Clear}
//Поиск
function TacadEntitiesStatistic.IndexOf(const AEntityType: Integer): Integer;
var I: Integer;
begin
  Result := -1;
  for I := 0 to Count-1 do
  if _Items[I].EntityType = AEntityType then
  begin
    Result := I; Break;
  end;{for}
end;{IndexOf}

//Приложение AutoCAD -----------------------------------------------------------
function TAutoCAD.GetLayer(const AIndex: Integer): TacadLayer;
begin
  if (AIndex < 0) or (AIndex >= LayersCount)
  then raise TacadException.InvalidIndex(AIndex,LayersCount-1);
  Result := TacadLayer(FLayers.List^[AIndex]);
end;{GetCoord}
function TAutoCAD._GetLayer(const AIndex: Integer): TacadLayer;
begin
  Result := TacadLayer(FLayers.List^[AIndex]);
end;{_GetCoord}
//Конструктор/деструктор
constructor TAutoCAD.Create();
begin
  inherited;
  FLayersCount := 0;
  FLayers      := nil;
  FStatistics  := TacadEntitiesStatistic.Create();
  FCubeBound   := acadBound();
  FCenter      := acadCoord3D();
end;{Create}
destructor TAutoCAD.Destroy;
begin
  FreeAndNil(FStatistics);
  _ClearLayers();
  inherited;
end;{Destroy}
//Извлечение слоя AutoCAD
procedure TAutoCAD._ExtractAutoCADLayer(const AacadLayer: IAcadLayer);
var ANew: TacadLayer;
begin
  if LayersCount = 0 then FLayers := TList.Create();
  ANew := TacadLayer.Create();

  //Графический объект
  ANew.FBound       := acadBound();
  ANew.FCenter      := acadCoord3D();
  ANew.FVisible     := AacadLayer.LayerOn = True;
  //Layer
  ANew.FName        := AacadLayer.Name;
  ANew.FFreeze      := AacadLayer.Freeze = False;
  ANew.FLock        := AacadLayer.Lock = False;
  ANew.FDefaultPen  := acadPen(acadColor(AacadLayer.TrueColor.Red,AacadLayer.TrueColor.Green,AacadLayer.TrueColor.Blue),
                                 acadLineSizeTo(AacadLayer.Lineweight, 0.1, 0.1),
                                 acadLineTypeTo(AacadLayer.Linetype, apsClear, apsClear));
  ANew.FDescription := AacadLayer.Description;
  FLayers.Add(ANew);
  FLayersCount      := FLayers.Count;
end;{_ExtractAutoCADLayer}
//Поиск слоя
function TAutoCAD._FindLayer(const ALayerName: String): Integer;
var I: Integer;
begin
  Result := -1;
  for I := 0 to LayersCount-1 do
  if CompareText(ALayerName,_Layers[I].Name) = 0 then
  begin
    Result := I; Break;
  end;{for}
end;{_FindLayer}
//Уничтожение слоев
procedure TAutoCAD._ClearLayers();
var I: Integer;
begin
  if LayersCount > 0 then
  begin
    FLayersCount := 0;
    for I := LayersCount-1 downto 0 do
      _Layers[I].Free();
    FreeAndNil(FLayers);
  end;{if}
  //Bound
  _DefineBound();
end;{_ClearLayers}
//Определение глобального контура
procedure TAutoCAD._DefineBound();
var I, J: Integer;
begin
  FCubeBound   := acadBound();
  FCenter      := acadCoord3D();
  if LayersCount > 0 then
  begin
    FCubeBound := _Layers[0].Bound;
    FCenter    := _Layers[0].Center;
    for I := 1 to LayersCount - 1 do
    for J := 0 to 2 do
    begin
      if FCubeBound.Min[J] > _Layers[I].Bound.Min[J] then FCubeBound.Min[J] := _Layers[I].Bound.Min[J];
      if FCubeBound.Max[J] < _Layers[I].Bound.Max[J] then FCubeBound.Max[J] := _Layers[I].Bound.Max[J];
      FCenter[J] := (_Layers[I].Bound.Min[J] + _Layers[I].Bound.Max[J]) * 0.5;
    end;{for}
  end;{if}
end;{_DefineBound}
//Прорисовка
procedure TAutoCAD.Draw(const ACanvas: TCanvas; const ACanvasBound: TRect);
var
  I                     : Integer;
  ANewCanvasBound       : TRect;
  ANewCubeBound         : RacadBound;
  ARatio                : Double;
  ACubeHeight,ACubeWidth: Double;
begin
  if Assigned(ACanvas)and(ACanvasBound.Left < ACanvasBound.Right)and(ACanvasBound.Top < ACanvasBound.Bottom) then
  begin
    ACanvas.Pen.Width   := 1;
    ACanvas.Pen.Style   := psSolid;
    ACanvas.Brush.Style := bsSolid;
    //Foreground
    ACanvas.Brush.Color := clWhite;
    ACanvas.Pen.Color   := clBlack;
    ACanvas.Rectangle(ACanvasBound);
    //Inflate CanvasBound
    ANewCanvasBound := ACanvasBound;
    InflateRect(ANewCanvasBound,-10,-10);
    //Define CubeBound
    ARatio := (ACanvasBound.Right - ACanvasBound.Left)/(ACanvasBound.Bottom - ACanvasBound.Top);//=Canvas.Width/Canvas.Height
    ACubeHeight := CubeBound.Max[1]-CubeBound.Min[1];
    ACubeWidth := ACubeHeight * ARatio;
    if ACubeWidth < CubeBound.Max[0]-CubeBound.Min[0] then//если "обрезка" по ширине
    begin
      ACubeWidth := CubeBound.Max[0]-CubeBound.Min[0];
      ACubeHeight := ACubeWidth / ARatio;
    end;{if}
    ANewCubeBound := CubeBound;
    ANewCubeBound.Max[0] := ANewCubeBound.Min[0] + ACubeWidth;
    ANewCubeBound.Max[1] := ANewCubeBound.Min[1] + ACubeHeight;
    //Layers
    for I := 0 to LayersCount-1 do
      _Layers[I].DrawToCanvas(ACanvas, ANewCanvasBound, ANewCubeBound);
  end;{if}
end;{Draw}
//Очистка
procedure TAutoCAD.Clear();
begin
  Statistics.Clear();
  _ClearLayers();
end;{Clear}
//Импорт из файла AutoCAD
function TAutoCAD.ImportFromAutoCADFile(const AFileName: String): Boolean;
var
  AApp       : IAcadApplication;   //Приложение AutoCAD
  ADoc       : IAcadDocument;      //Активный документ AutoCAD
  ALayers    : IAcadLayers;        //Слои AutoCAD
  AObjects   : IAcadModelSpace;    //Объекты AutoCAD
  AAcadObj   : IAcadEntity;        //Текущий объект AutoCAD
  AObject    : TacadLayerObject;//Пользовательский объект
  AClassRef  : TacadLayerObjectRef;//Ссылка на тип класса TacadLayerObject
  ALayerIndex: Integer;            //Индекс слоя
  //
  I          : Integer;
begin
  Result := False;
  if FileExists(AFileName) then
  begin
    Clear();
    //AutoCAD.Objects ----------------------------------------------------------
    AApp := IDispatch(CreateOleObject('AutoCAD.Application')) as IAcadApplication;
    try
      AApp.Visible := False;
      //Активный документ ------------------------------------------------------
      if AApp.Documents.Count > 0 then AApp.Documents.Close();
      ADoc := AApp.Documents.Open(AFileName, True, Null);
      //Слои -------------------------------------------------------------------
      ALayers := ADoc.Layers;
      for I := 0 to ALayers.Count - 1 do
        _ExtractAutoCADLayer(ALayers.Item(I) as IAcadLayer);
      //Объекты слоя -----------------------------------------------------------
      AObjects := ADoc.ModelSpace;
      for I := 0 to AObjects.Count - 1 do
      begin
        AAcadObj := AObjects.Item(I);
        case AAcadObj.EntityType of
          acPoint          : AClassRef := TacadPoint;
          acLine           : AClassRef := TacadLine;
          acCircle         : AClassRef := TacadCircle;
          acArc            : AClassRef := TacadArc;
          acSpline         : AClassRef := TacadSpline;
          acEllipse        : AClassRef := TacadEllipse;
          acPolylineLight  : AClassRef := TacadPolyline;
          ac3dPolyline     : AClassRef := Tacad3DPolyline;
          acBlockReference : AClassRef := TacadBlock;
          acHatch          : AClassRef := TacadPolygon;
          //acText           : AClassRef := TacadText;
          //acMText          : AClassRef := TacadMText;
          else AClassRef := nil;
        end;{case}
        if AClassRef <> nil then
        begin
          //Определение принадлежности объекта слою
          ALayerIndex := _FindLayer(AAcadObj.Layer);
          if ALayerIndex <> -1 then
          begin
            //Создание соответствующего объекта
            AObject := AClassRef.Create();
            AObject.ExtractAcadEntity(AAcadObj, _Layers[ALayerIndex], nil);
            //Обработка блока
            if AObject is TacadBlock
            then TacadBlock(AObject).ExtractAcadEntityDn(_Layers[ALayerIndex], ADoc.Blocks);
            _Layers[ALayerIndex].Add(AObject);
          end;{if}
          Statistics.Add(AAcadObj.EntityType, AAcadObj.EntityName, ALayerIndex = -1);
        end{if}
        else Statistics.Add(AAcadObj.EntityType, AAcadObj.EntityName, True);
      end;{for}
      _DefineBound();
    finally
      AApp.Quit;
    end;{try}
  end;{if}
end;{ImportFromAutoCADFile}

end.

