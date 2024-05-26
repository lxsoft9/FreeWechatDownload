{***********************************************************}
{                                                           }
{    HotPDF PDF Component                                   }
{    2007-2018, https://www.loslab.com/                     }
{                                                           }
{***********************************************************}

unit HPDFDoc;

interface

{$I HotPDF.inc }

uses
{$IFDEF D16}System.Classes, System.SysUtils, System.Math, Winapi.Windows, Winapi.ShellAPI, Vcl.Graphics, Vcl.Imaging.jpeg,
{$ELSE}Classes, SysUtils, Math, Windows, ShellAPI, Graphics, jpeg,
{$ENDIF}HPDFTypes, HPDFObjs, HPDFBarcode, HPDFFonts, HPDFZLib, HPDFCrypt;

type
  THPDFJustificationType = (jtLeft, jtCenter, jtRight);

  THPDFTSh = record
    AValText: AnsiString;
    AValX: Single;
  end;

  THPDFCurrPoint = record
    X: Extended;
    Y: Extended;
  end;

  THPDFRect = record
    Left, Top, Right, Bottom: Extended;
  end;

  THPDFParagraph = record
    Justification: THPDFJustificationType;
    Indention: Single;
    LeftMargin: Single;
    RightMargin: Single;
    TopMargin: Single;
    BottomMargin: Single;
  end;

  TGenByteArray = array[0..9] of byte;
  TGenerationArray = array[0..9] of Integer;

  THPDFImageObject = record
    Index: Integer;
    Name: AnsiString;
    ImageObject: THPDFObject;
    Width: Integer;
    Height: Integer;
  end;

  XRefItem = record
    Loaded: boolean;
    Offset: Integer;
    ObjNumber: THPDFObjectNumber;
  end;

  THPDFKeyType = (k40, k128, aes128);
  TPDFVersType = (pdf10, pdf11, pdf12, pdf13, pdf14, pdf15, pdf16, pdf17, pdf18, pdf19, pdf110);
  THPDFActionScriptType = (astOpen, astClose, astWillSave, astDidSave, astWillPrint, astDidPrint);
  THPDFCompressionMethod = (cmNone, cmFlateDecode);
{$IFDEF BCB}
  THPDFPageOrientation = (vpoPortrait, vpoLandscape);
{$ELSE}
  THPDFPageOrientation = (poPortrait, poLandscape);
{$ENDIF}
  TLineCapStyle = (lcButtEnd, lcRoundEnd, lcProjectSquareEnd);
  TLineJoinStyle = (ljMiterJoin, ljRoundJoin, ljBevelJoin);
  TPDFTextRenderingMode = (trFill, trStroke, trFillThenStroke, trInvisible,
    trFillClipping, trStrokeClipping, trFillStrokeClipping, trClipping);
  THPDFImageCompressionType = (icFlate, icJpeg, icCCITT31, icCCITT32, icCCITT42);
  THPDFPageLayout = (plSinglePage, plOneColumn, plTwoColumnLeft, plTwoColumnRight);
  THPDFAnnotationSubType = (asTextNotes, asFreeText, asLine, asSquare, asCircle, asStamp, asFileAttachment, asSound, asMovie);
  THPDFTextAnnotationType = (taComment, taKey, taNote, taHelp, taNewParagraph, taParagraph, taInsert);
  THPDFFreeTextAnnotationJust = (ftLeftJust, ftCenter, ftRightJust);
  THPDFCSAnnotationType = (csCircle, csSquare);
  THPDFStampAnnotationType = (satApproved, satExperimental, satNotApproved,
    satAsIs, satExpired, satNotForPublicRelease, satConfidential, satFinal,
    satSold, satDepartmental, satForComment, satTopSecret, satDraft, satForPublicRelease);
  THPDFPageSize = (psUserDefined, psLetter, psA4, psA3, psLegal, psB5, psC5,
    ps8x11, psB4, psA5, psFolio, psExecutive, psEnvB4, psEnvB5, psEnvC6,
    psEnvDL, psEnvMonarch, psEnv9, psEnv10, psEnv11);
  THPDFBarcodeType =
    (bcCode_2_5_interleaved, bcCode_2_5_industrial, bcCode_2_5_matrix,
    bcCode39, bcCode39Extended, bcCode128A, bcCode128B, bcCode128C, bcCode93,
    bcCode93Extended, bcCodeMSI, bcCodePostNet, bcCodeCodabar, bcCodeEAN8, bcCodeEAN13,
    bcCodeUPC_A, bcCodeUPC_E0, bcCodeUPC_E1, bcCodeUPC_Supp2, bcCodeUPC_Supp5,
    bcCodeEAN128A, bcCodeEAN128B, bcCodeEAN128C);
  THPDFPageMode = (pmUseNone, pmUseOutlines, pmUseThumbs, pmFullScreen, pmUseAttachments);
  THPDFViewerPreference = (vpHideToolbar, vpHideMenubar, vpHideWindowUI,
    vpFitWindow, vpCenterWindow);
  THPDFProtection = (prPrint, prModifyStructure, prInformationCopy,
    prEditAnnotations, prPrint12bit, prFillAnnotations, prExtractContent, prAssemble);
  TTIFFPaintType = (tptResizePage, tptResizeImage);
  THPDFProtectOptions = set of THPDFProtection;
  THPDFViewerPreferences = set of THPDFViewerPreference;

  THPDFFontObj = class(TObject)
  private
    Name: AnsiString;
    Size: Single;
    ArrIndex: Integer;
    Saved: boolean;
    OldName: AnsiString;
    Ascent: Integer;
    FActive: boolean;
    IsUsed: boolean;
    UniLen: Integer;
    FontLen: Integer;
    IsUnicode: boolean;
    IsVertical: boolean;
    OrdinalName: AnsiString;
    IsStandard: boolean;
    FontStyle: TFontStyles;
    FontCharset: TFontCharset;
    IsMonospaced: boolean;
    OutTextM: OUTLINETEXTMETRIC;
    ABCArray: array[0..255] of ABC;
    Symbols: array of CDescript;
    UnicodeTable: array of IndexedChar;
    SymbolTable: array[32..255] of boolean;
    function GetCharWidth(AText: AnsiString; APos: Integer): Integer;
    procedure CopyFontFetures(InFnt: THPDFFontObj);
    procedure GetFontFeatures;
    procedure ParseFontName;
    procedure ClearTables;
  end;

  THotPDF = class;
  THPDFPage = class;

  THPDFAnnotation = class(TObject)
  private
    FPage: THPDFPage;
    FColor: array[0..3] of Byte;
    FType: THPDFAnnotationSubType;
    FTypeState: AnsiString;
    FContents: AnsiString;
    FQuadding: Integer;
    FOpen: boolean;
    FLeftTop: THPDFCurrPoint;
    FRightBottom: THPDFCurrPoint;
    FName: AnsiString;
  protected
    procedure AddAnnotationObject;
  end;

  THPDFDocOutlineObject = class(TObject)
  private
    FDoc: THotPDF;
    FParent: THPDFDocOutlineObject;
    FNext: THPDFDocOutlineObject;
    FPrev: THPDFDocOutlineObject;
    FFirst: THPDFDocOutlineObject;
    FLast: THPDFDocOutlineObject;
    FTitle: AnsiString;
    FOpened: boolean;
  protected
    LinkedObj: THPDFDictionaryObject;
    FCount: Integer;
    FTop: Integer;
    FLeft: Integer;
    procedure Init(AOwner: THotPDF);
  public
    function AddChild(Title: AnsiString; X: Single = 0; Y: Single = 0): THPDFDocOutlineObject;
    property Parent: THPDFDocOutlineObject read FParent;
    property Next: THPDFDocOutlineObject read FNext;
    property Prev: THPDFDocOutlineObject read FPrev;
    property First: THPDFDocOutlineObject read FFirst;
    property Last: THPDFDocOutlineObject read FLast;
    property Title: AnsiString read FTitle write FTitle;
    property Opened: boolean read FOpened write FOpened;
  end;

  THPDFPage = class(TObject)
  private
    DPI: Single;
    FWordSpace: Single;
    FCharSpace: Single;
    FLeading: Single;
    FHorizontalScaling: Single;
    ISStated: Integer;
    STextBegin: boolean;
    FKBegin: boolean;
    SKBegin: boolean;
    FFillColor: TColor;
    FStrokeColor: TColor;
    FHyperColor: TColor;
    FMinXVal: Single;
    FMinYVal: Single;
    FMaxXVal: Single;
    FMaxYVal: Single;
    FDescStream: TStream;
    PageObjStream: THPDFStreamObject;
    PageObj: THPDFDictionaryObject;
    FResolution: Integer;
    FUpdateSize: boolean;
    mWidth: Single;
    mHeight: Single;
    FWidth: Single;
    FHeight: Single;
    FAnnotsObj: THPDFArrayObject;
    FDocScale: Extended;
    FSize: THPDFPageSize;
    XObjectObj: THPDFObject;
    FontObjectObj: THPDFObject;
    XObjectNames: array of AnsiString;
    FontArr: array of THPDFFontObj;
    CurrentFontObj: THPDFFontObj;
    FMediaBoxArray: THPDFArrayObject;
    FOrientation: THPDFPageOrientation;
    FTopTextPosition: boolean;
    PageContent: TStringList;
    PageMeta: TMetafile;
    FCanvas: TCanvas;
    procedure CloseCanvas;
    procedure ClosePage;
    procedure CalculateWidthFormat;
    procedure CalculateHeightFormat;
    procedure CalculateFormat;
    procedure ConvertPageObject;
    procedure SetResolution(const Value: Integer);
    function GetPageHeight: Single;
    function GetPageWidth: Single;
    procedure SetPageHeight(const Value: Single);
    procedure SetPageWidth(const Value: Single);
    procedure SetPageSize(const Value: THPDFPageSize);
    procedure SetOrientation(const Value: THPDFPageOrientation);
    procedure TurnPage;
    procedure StoreFont(FontObj: THPDFFontObj);
    procedure SaveToPageStream(ValStr: AnsiString);
    function GetCanvas: TCanvas;
    function CompareCurrentFont: Integer;
    function XProjection(X: Single): Single;
    function YProjection(Y: Single): Single;
    function CompareResName(Index: Integer; Name: AnsiString): Integer;
    procedure CurveArc(CenterX, CenterY, RadiusX, RadiusY, StartAngle, SweepRange: Extended; UseMoveTo: boolean);
    procedure ExecuteXObject(XOName: AnsiString);
    procedure DrawXObjectEx(X, Y, AWidth, AHeight: Single; ClipX, ClipY, ClipWidth, ClipHeight: Single; AXObjectName: AnsiString; angle: Extended; DocScale: Extended);
  protected
    function ProcChar(ComposChar: WORD): AnsiString;
    procedure RMoveTo(X, Y: Single);
    procedure StoreCurrentFont;
    procedure InternUnicodeTextOut(X, Y: Single; angle: Extended; Text: PWORD; TextLength: Integer);
    procedure FEllipse(X, Y, Width, Height: Single);
    procedure RotateCoordinate(X, Y, angle: Extended; var XR, YR: Extended);
    procedure SetURIObject(UriLink: AnsiString; Left, Top, Right, Bottom: Extended);
  public
    FParent: THotPDF;
    constructor Create;
    destructor Destroy; override;
    procedure DrawPie(X1, Y1, X2, Y2, X3, Y3, X4, Y4: Extended);
    function DrawArc(X1, Y1, X2, Y2, X3, Y3, X4, Y4: Single): THPDFCurrPoint;
    procedure GStateSave;
    function GStateRestore: boolean;
    procedure Concat(A, B, C, D, E, F: Single);
    procedure SetFlat(Flatness: byte);
    procedure SetLineCap(Linecap: TLineCapStyle);
    procedure SetDash(DashArray: array of byte; Phase: byte);
    procedure NoDash;
    procedure AddTextAnnotation(Contents: AnsiString; Rectangle: TRect; Open: boolean; Name: THPDFTextAnnotationType; Color: TColor = clRed);
    procedure AddFreeTextAnnotation(Contents: AnsiString; Rectangle: TRect; Quadding: THPDFFreeTextAnnotationJust; Color: TColor = clRed);
    procedure AddLineAnnotation(Contents: AnsiString; BeginPoint: THPDFCurrPoint; EndPoint: THPDFCurrPoint; Color: TColor = clRed);
    procedure AddCircleSquareAnnotation(Contents: AnsiString; Rectangle: TRect; CSType: THPDFCSAnnotationType; Color: TColor = clRed);
    procedure AddStampAnnotation(Contents: AnsiString; Rectangle: TRect; StampType: THPDFStampAnnotationType; Color: TColor = clRed);
    procedure AddFileAttachmentAnnotation(Contents: AnsiString; FileName: AnsiString; Rectangle: TRect; Color: TColor = clRed);
    procedure AddSoundAnnotation(Contents: AnsiString; FileName: AnsiString; Rectangle: TRect; Color: TColor = clRed);
    procedure AddMovieAnnotation(Contents: AnsiString; FileName: AnsiString; Rectangle: TRect; Color: TColor = clRed);
    procedure SetLineJoin(LineJoin: TLineJoinStyle);
    procedure SetLineWidth(Width: Single);
    procedure SetMiterLimit(MiterLimit: byte);
    procedure CurveToC(X1, Y1, X2, Y2, X3, Y3: Single);
    procedure CurveToV(X2, Y2, X3, Y3: Single);
    procedure CurveToY(X1, Y1, X3, Y3: Single);
    procedure FMEllipse(X, Y, X1, Y1: Single);
    procedure MFRectangle(X, Y, X1, Y1: Single);
    procedure Rectangle(X, Y, Width, Height: Single);
    procedure RectangleRotate(X, Y, Width, Height: Single; angle: Extended);
    procedure SetCharacterSpacing(Spacing: Single);
    procedure SetWordSpacing(Spacing: Single);
    procedure SetHorizontalScaling(Scaling: Single);
    procedure SetLeading(Leading: Single);
    procedure SetTextRenderingMode(Mode: TPDFTextRenderingMode);
    procedure SetTextRise(Rise: SmallInt);
    procedure MoveTextPoint(X, Y: Single);
    procedure SetTextMatrix(A, B, C, D, X, Y: Single);
    procedure MoveToNextLine;
    procedure ShowText(Text: AnsiString; IsHexadecimal: boolean = false);
    procedure ShowUnicodeText(Text: WideString);
    procedure SetRGBHyperlinkColor(Value: TColor);
    procedure SetRGBColor(Value: TColor);
    procedure SetRGBFillColor(Value: TColor);
    procedure SetRGBStrokeColor(Value: TColor);
    procedure SetGrayColor(GrayColor: Extended);
    procedure SetGrayFillColor(GrayColor: Extended);
    procedure SetGrayStrokeColor(GrayColor: Extended);
    procedure RoundRect(X1, Y1, X2, Y2, X3, Y3: Integer);
    procedure Circle(X, Y, Radius: Single);
    procedure Ellipse(X, Y, Width, Height: Single);
    procedure DrawBarcode(BCType: THPDFBarcodeType; X, Y, Height, MUnit: Integer; Angle: Single; Info: AnsiString; 
      UseCheckSum: boolean; BarColor, Color: TColor);
    procedure ProcessFont(Unicoded: boolean);
    procedure PrintHyperlink(X, Y: Single; Text, Link: AnsiString);
    procedure SetFont(FontName: AnsiString; FontStyle: TFontStyles; ASize: Single;
      FontCharset: TFontCharset = ANSI_CHARSET; IsVertical: boolean = false);
{$IFDEF BCB}
    procedure UnicodeTextOut(X, Y: Single; Angle: Extended; Text: PWORD; TextLength: Integer);
    procedure UnicodeTextOutStr(X, Y: Single; Angle: Extended; Text: WideString);
    procedure PrintText(X, Y: Single; angle: Extended; Text: AnsiString);
    procedure OutText(X, Y: Single; angle: Extended; Text: AnsiString);
{$ELSE}
    procedure UnicodeTextOut(X, Y: Single; Angle: Extended; Text: PWORD; TextLength: Integer); overload;
    procedure UnicodeTextOut(X, Y: Single; Angle: Extended; Text: WideString); overload;
    procedure TextOut(X, Y: Single; angle: Extended; Text: AnsiString);
{$ENDIF}
    function TextWidth(Text: AnsiString): Single;
    function TextHeight(Text: AnsiString): Real;
    procedure SetFontAndSize(FontName: AnsiString; Size: Single);
{$IFDEF BCB}
    procedure ShowMetafileEx(MetaFile: TMetafile);
    procedure ShowMetafile(MetaFile: TMetafile; X, Y, HorScale, VertScale: Extended);
{$ELSE}
    procedure ShowMetafile(MetaFile: TMetafile); overload;
    procedure ShowMetafile(MetaFile: TMetafile; X, Y, HorScale, VertScale: Extended); overload;
{$ENDIF}
    procedure BeginText;
    procedure EndText;
    procedure ShowImage(ImageIndex: Integer; X, Y, w, h, angle: Extended);
    procedure MoveTo(X, Y: Single);
    procedure LineTo(X, Y: Single);
    procedure Stroke;
    procedure ClosePath;
    procedure NewPath;
    procedure ClosePathStroke;
    procedure Fill;
    procedure EoFill;
    procedure FillAndStroke;
    procedure ClosePathFillAndStroke;
    procedure EoFillAndStroke;
    procedure ClosePathEoFillAndStroke;
    procedure Clip;
    procedure EoClip;
    property Canvas: TCanvas read GetCanvas;
    property TopTextPosition: boolean read FTopTextPosition write FTopTextPosition;
    property DocScale: Extended read FDocScale write FDocScale;
    property Orientation: THPDFPageOrientation read FOrientation write SetOrientation;
    property Resolution: Integer read FResolution write SetResolution;
    property Size: THPDFPageSize read FSize write SetPageSize;
    property Width: Single read GetPageWidth write SetPageWidth;
    property Height: Single read GetPageHeight write SetPageHeight;
  end;

  THPDFPara = class(TObject)
  private
    FParent: THotPDF;
    FJustification: THPDFJustificationType;
    Indention: Single;
    LeftMargin: Single;
    RightMargin: Single;
    TopMargin: Single;
    BottomMargin: Single;
    procedure InternShowText(Text: AnsiString; var APrText: AnsiString; var APrVax: Single);
    procedure InternShowUnicodeText(Text: WideString; var APrText: WideString; var APrVax: Single);
    procedure SetJustification(const Value: THPDFJustificationType);
    constructor Create(Parent: THotPDF);
    function GetCurrentLine: Single;
    procedure SetCurrentLine(const Value: Single);
  public
{$IFDEF BCB}
    procedure PrintText(Text: AnsiString);
{$ELSE}
    procedure ShowText(Text: AnsiString);
{$ENDIF}
    procedure ShowUnicodeText(Text: WideString);
    procedure NewLine;
    property Justification: THPDFJustificationType read FJustification write SetJustification;
    property CurrentLine: Single read GetCurrentLine write SetCurrentLine;
  end;

  THotPDF = class(TComponent)
  private
  {private}
    FAutoAddPage: boolean;
    DocID: AnsiString;
    OwUsPass: AnsiString;
    UsPassStr: AnsiString;
    DocScale: Extended;
    ProtectFlags: Integer;
    FRevision: Integer;
    FCurrentKey: MD5Digest;
    FNewParaLine: boolean;
    FCurrentParagraph: THPDFPara;
    FPrevParaOffset: Single;
    FCurrentParaLine: Single;
    FCurrentImageIndex: Integer;
    FCurrentFontIndex: Integer;
    FVPChanged: boolean;
    FIsLoaded: boolean;
    FPagesCount: Integer;
    FCurrentPageNum: Integer;
    FDocStarted: boolean;
    FProgress: boolean;
    FIsEncrypted: boolean;
    FEncryprtLink: THPDFObjectNumber;
    FPageslink: THPDFObjectNumber;
    FRootLink: THPDFObjectNumber;
    FInfoLink: THPDFObjectNumber;
    TrailerObj: THPDFDictionaryObject;
    FastOpen: AnsiString;
    FastClose: AnsiString;
    FastWillSave: AnsiString;
    FastDidSave: AnsiString;
    FastWillPrint: AnsiString;
    FastDidPrint: AnsiString;
    FPagesIndex: Integer;
    FEncryprtIndex: Integer;
    FRootIndex: Integer;
    FInfoIndex: Integer;
    FInStream: TStream;
    FOutlineRoot: THPDFDocOutlineObject;
    FXrefLen: Integer;
    FMaxObjNum: Integer;
    FResolution: Integer;
    PageArrPosition: Integer;
    FActivePara: Integer;
    FParas: array of THPDFParagraph;
    FParaLen: Integer;
    FXref: array of XRefItem;
    PageArr: array of THPDFDictArrItem;
    CSArray: array of THPDFCopyStructItem;
    FParentMB: array[0..3] of Single;
    FIsParented: boolean;
    IndirectObjects: TList;
    FAutoLaunch: boolean;
    FShowInfo: boolean;
    FFontEmbedding: boolean;
    FStandardFontEmulation: boolean;
    FProtection: boolean;
    FCreationDate: TDateTime;
    FMemStream: boolean;
    FSubject: AnsiString;
    FTitle: AnsiString;
    FAuthor: AnsiString;
    FOwnerPassword: AnsiString;
    FKeywords: AnsiString;
    FUserPassword: AnsiString;
    FOutputStream: TStream;
    FontNames: array of AnsiString;
    FFileName: TFileName;
    FJpegQuality: TJPEGQualityRange;
    FVersion: TPDFVersType;
    FCurrentPage: THPDFPage;
    FNEmbeddedFonts: TStringList;
    FCompressionMethod: THPDFCompressionMethod;
    FCryptKeyType: THPDFKeyType;
    FPageLayout: THPDFPageLayout;
    FPageMode: THPDFPageMode;
    FProtectOption: THPDFProtectOptions;
    FViewerPreference: THPDFViewerPreferences;
    XImages: array of THPDFImageObject;
    FKeepImageAspectRatio: boolean;
    FImageCompressionType: THPDFImageCompressionType;
    FSizes: array of THPDF_SIZES;
    OutlineEnsemble: array of THPDFDocOutlineObject;
    OutlineEnsLen: Integer;
    SamCorrel: boolean;
    procedure LoadUnFlateLZW(RegionStream, StrumStream: TStream; NameObj: AnsiString);
    procedure CreateKeys;
    procedure EnableEncrypt;
    procedure PadTrunc(S: Pointer; SL: Integer; D: Pointer);
    procedure DeleteObj(InObj: THPDFObject; Recursive: boolean);
    function CrptStr(Data: AnsiString; Password: MD5Digest; PassLength: THPDFKeyType; ObjID: Integer): AnsiString;
    procedure CrptStrm(Data: TMemoryStream; Password: MD5Digest; PassLength: THPDFKeyType; ObjID: Integer);
    function CryptString(Data: AnsiString; ID: Integer): AnsiString;
    procedure CryptStream(Data: TMemoryStream; ID: Integer);
    function GetOutlineRoot: THPDFDocOutlineObject;
    function AddImageFromTIFF(Image: TBitmap; Compression: THPDFImageCompressionType): Integer;
    function GetCanvas: TCanvas;
    procedure SetOutputStream(const Value: TStream);
    procedure ListExtDictionary(PageObject: THPDFDictionaryObject; PageLink: THPDFObjectNumber);
    function CompareObjectID(IDL, IDR: THPDFObjectNumber): Integer;
    procedure SetAutoLaunch(const Value: boolean);
    procedure SetCompressionMethod(const Value: THPDFCompressionMethod);
    procedure SetCryptKeyType(const Value: THPDFKeyType);
    procedure SetNEmbeddedFont(const Value: TStringList);
    procedure SetPageLayout(const Value: THPDFPageLayout);
    procedure SetStandardFontEmulation(const Value: boolean);
    function GetObjectByLink(LinkObj: THPDFLink): THPDFObject;
    procedure SetCurrentPageNum(const Value: Integer);
  protected
   {protected}
    LinksLen: Integer;
    LinkTable: array of TGenerationArray;
    IsLinearized: boolean;
    function CopyObject(SourceDoc: THotPDF; SourceObj: THPDFObject): THPDFObject;
    procedure SetDocImagearray(Width: Integer; Height: Integer);
    procedure SetViewerPreferences(const Value: THPDFViewerPreferences);
    procedure StreamSaveString(DocStream: TStream; DocSting: AnsiString);
    procedure SetResolution(const Value: Integer);
    procedure CloseIndirectObjects;
    procedure LoadXrefArray;
    function LoadDocString: AnsiString;
    function LoadBackDocString: AnsiString;
    function LoadDocHeader: AnsiString;
    function FontIsEmbedded(FontNm: AnsiString): boolean;
    function CreateIndirectDictionary: THPDFDictionaryObject;
    function SaveTypeObject(ValObject: THPDFObject; ObjStream: TStream; IsArrayItem: boolean): Integer;
    function AddTypeObject(ObjType: THPDFObjectType; IsIndirect: boolean): THPDFObject;
    function LoadBooleanObject(ObjStream: TStream): THPDFBooleanObject;
    function LoadNumericObject(ObjStream: TStream): THPDFNumericObject;
    function LoadStringObject(ObjStream: TStream): THPDFStringObject;
    function LoadNameObject(ObjStream: TStream): THPDFNameObject;
    function LoadArrayObject(ObjStream: TStream): THPDFArrayObject;
    function LoadStreamObject(ObjStream: TStream): THPDFStreamObject;
    function LoadDictionaryObject(ObjStream: TStream): THPDFDictionaryObject;
    function LoadLinkObject(ObjStream: TStream): THPDFLink;
    function SaveObjectValue(ValObject: THPDFObject; Value: AnsiString; ObjStream: TStream): Integer;
    function SaveNullObject(ValObject: THPDFNullObject; ObjStream: TStream): Integer;
    function SaveBooleanObject(ValObject: THPDFBooleanObject; ObjStream: TStream): Integer;
    function SaveNumericObject(ValObject: THPDFNumericObject; ObjStream: TStream): Integer;
    function SaveStringObject(ValObject: THPDFStringObject; ObjStream: TStream): Integer;
    function SaveNameObject(ValObject: THPDFNameObject; ObjStream: TStream): Integer;
    function SaveArrayObject(ValObject: THPDFArrayObject; ObjStream: TStream): Integer;
    function SaveStreamObject(ValObject: THPDFStreamObject; ObjStream: TStream): Integer;
    function SaveDictionaryObject(ValObject: THPDFDictionaryObject; ObjStream: TStream; IsArrayItem: boolean): Integer;
    function SaveLinkObject(ValObject: THPDFLink; ObjStream: TStream): Integer;
    function MapImage(Image: TGraphic; Compression: THPDFImageCompressionType;
      IsMask: boolean; MaskIndex: Integer): Integer;
    function LoadIsLinearized: boolean;
    function GetIsHaveSimpleText: boolean;
    function GetObjectNumber: THPDFObjectNumber;
    function GetObjectType(ObjStream: TStream; UseIndirect: boolean): THPDFObjectType;
    procedure SaveToFile(FileName: TFileName);
    procedure SaveToStream(DocStream: TStream);
  public
    FCHandle: HDC;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
{$IFDEF BCB}
    function AddImage(Image: TGraphic; Compression: THPDFImageCompressionType;
      IsMask: boolean = false; MaskIndex: Integer = -1): Integer;
    function AddImageFromFile(FileName: TFileName; Compression:
      THPDFImageCompressionType; IsMask: boolean = False; MaskIndex: Integer = -1): Integer;
{$ELSE}
    function AddImage(Image: TGraphic; Compression: THPDFImageCompressionType;
      IsMask: boolean = false; MaskIndex: Integer = -1): Integer; overload;
    function AddImage(FileName: TFileName; Compression:
      THPDFImageCompressionType; IsMask: boolean = False; MaskIndex: Integer = -1): Integer; overload;
{$ENDIF}
    function AddPage: Integer;
    procedure AddDocumentAttachment(FileName: TFileName; Description: AnsiString);
    procedure CopyPageFromDocument(SourceDoc: THotPDF; SourceIndex: Integer; DestIndex: Integer);
    procedure SetActionScript(ActionType: THPDFActionScriptType; ActionText: AnsiString);
    procedure DeletePage(PageIndex: Integer);
    procedure SetCurrentPageNumber(const Value: Integer);
    function LoadFromFile(FileName: TFileName): Integer;
    function LoadFromStream(DocStream: TStream): Integer;
    function CreateParagraph(Indention: Single = 0; Justification: THPDFJustificationType = jtLeft; LeftMargin: Single = 0; RightMargin: Single = 0; TopMargin: Single = 0; BottomMargin: Single = 0): Integer;
    procedure BeginDoc(Initial: boolean = false);
    procedure EndDoc;
    procedure BeginParagraph(Index: Integer);
    procedure EndParagraph;
    procedure AddTiffFromFile(FileName: TFileName; Compression: THPDFImageCompressionType; PaintType: TTIFFPaintType);
    property IsHaveSimpleText: boolean read GetIsHaveSimpleText;
    property IsEncrypted: boolean read FIsEncrypted;
    property Canvas: TCanvas read GetCanvas; // Only for old version compatibility
    property CurrentParagraph: THPDFPara read FCurrentParagraph;
    property CurrentPage: THPDFPage read FCurrentPage;
    property OutputStream: TStream read FOutputStream write SetOutputStream;
    property CurrentPageNumber: Integer read FCurrentPageNum write SetCurrentPageNum;
    property OutlineRoot: THPDFDocOutlineObject read GetOutlineRoot;
  published
    property AutoLaunch: boolean read FAutoLaunch write SetAutoLaunch;
    property ShowInfo: boolean read FShowInfo write FShowInfo;
    property Author: AnsiString read FAuthor write FAuthor;
    property Keywords: AnsiString read FKeywords write FKeywords;
    property Subject: AnsiString read FSubject write FSubject;
    property Title: AnsiString read FTitle write FTitle;
    property FileName: TFileName read FFileName write FFileName;
    property Version: TPDFVersType read FVersion write FVersion;
    property UserPassword: AnsiString read FUserPassword write FUserPassword;
    property OwnerPassword: AnsiString read FOwnerPassword write FOwnerPassword;
    property ActivateProtection: boolean read FProtection write FProtection;
    property CryptKeyLength: THPDFKeyType read FCryptKeyType write SetCryptKeyType;
    property ProtectOptions: THPDFProtectOptions read FProtectOption write FProtectOption;
    property NotEmbeddedFonts: TStringList read FNEmbeddedFonts write SetNEmbeddedFont;
    property PageLayout: THPDFPageLayout read FPageLayout write SetPageLayout;
    property PageMode: THPDFPageMode read FPageMode write FPageMode;
    property PagesCount: Integer read FPagesCount;
    property StandardFontEmulation: boolean read FStandardFontEmulation write SetStandardFontEmulation;
    property FontEmbedding: boolean read FFontEmbedding write FFontEmbedding;
    property ViewerPreferences: THPDFViewerPreferences read FViewerPreference write SetViewerPreferences;
    property JpegQuality: TJPEGQualityRange read FJpegQuality write FJpegQuality;
    property Compression: THPDFCompressionMethod read FCompressionMethod write SetCompressionMethod;
    property KeepImageAspectRatio: boolean read FKeepImageAspectRatio write FKeepImageAspectRatio default true;
    property Resolution: Integer read FResolution write SetResolution;
    property ParaAutoAddPage: boolean read FAutoAddPage write FAutoAddPage;
    property ImageCompressionType: THPDFImageCompressionType read FImageCompressionType write FImageCompressionType default icJpeg;
  end;

procedure Register;

implementation

uses HPDFWmf, HPDFImage, HPDFTiff;

const
  EscapeChars = [' ', '/', '(', '<', '[', '>', ']', #10, #13];

{ THPDFAnnotation }

procedure THPDFAnnotation.AddAnnotationObject;
var
  MFS: TFileStream;
  AttS: THPDFStreamObject;
  AspectArray: THPDFArrayObject;
  ColorArray: THPDFArrayObject;
  RectArray: THPDFArrayObject;
  LineArray: THPDFArrayObject;
  MediaF: THPDFDictionaryObject;
  FilSpec: THPDFDictionaryObject;
  AnotDict: THPDFDictionaryObject;

  function ConvertFileName(FileName: TFileName): TFileName;
  var
    I: Integer;
    FNL: Integer;
  begin
    FNL := Length(FileName);
    for I := 1 to FNL do
    begin
      if (FileName[I] = '\') then FileName[I] := '/';
    end;
    result := FileName;
  end;

begin
  AnotDict := FPage.FParent.CreateIndirectDictionary;
  AnotDict.AddNameValue('Type', 'Annot');
  ColorArray := THPDFArrayObject.Create(nil);
  ColorArray.AddNumericValue(FColor[0] / 255);
  ColorArray.AddNumericValue(FColor[1] / 255);
  ColorArray.AddNumericValue(FColor[2] / 255);
  if (FPage.FAnnotsObj = nil) then
  begin
    FPage.FAnnotsObj := THPDFArrayObject.Create(nil);
    FPage.PageObj.AddValue('Annots', FPage.FAnnotsObj);
  end;
  case FType of
    asTextNotes:
      begin
        AnotDict.AddNameValue('Subtype', 'Text');
        AnotDict.AddNameValue('Name', FName);
        RectArray := THPDFArrayObject.Create(nil);
        RectArray.AddNumericValue(FPage.XProjection(FLeftTop.X));
        RectArray.AddNumericValue(FPage.YProjection(FRightBottom.Y));
        RectArray.AddNumericValue(FPage.XProjection(FRightBottom.X));
        RectArray.AddNumericValue(FPage.YProjection(FLeftTop.Y));
        AnotDict.AddValue('Rect', RectArray);
        AnotDict.AddStringValue('Contents', FContents);
        AnotDict.AddBooleanValue('Open', FOpen);
        AnotDict.AddValue('P', FPage.PageObj);
      end;
    asFreeText:
      begin
        AnotDict.AddNameValue('Subtype', 'FreeText');
        RectArray := THPDFArrayObject.Create(nil);
        RectArray.AddNumericValue(FPage.XProjection(FLeftTop.X));
        RectArray.AddNumericValue(FPage.YProjection(FRightBottom.Y));
        RectArray.AddNumericValue(FPage.XProjection(FRightBottom.X));
        RectArray.AddNumericValue(FPage.YProjection(FLeftTop.Y));
        AnotDict.AddValue('Rect', RectArray);
        AnotDict.AddStringValue('Contents', FContents);
        AnotDict.AddNumericValue('Q', FQuadding);
        AnotDict.AddValue('P', FPage.PageObj);
      end;
    asLine:
      begin
        AnotDict.AddNameValue('Subtype', 'Line');
        AnotDict.AddStringValue('Contents', FContents);
        LineArray := THPDFArrayObject.Create(nil);
        LineArray.AddNumericValue(FPage.XProjection(FLeftTop.X));
        LineArray.AddNumericValue(FPage.YProjection(FLeftTop.Y));
        LineArray.AddNumericValue(FPage.XProjection(FRightBottom.X));
        LineArray.AddNumericValue(FPage.YProjection(FRightBottom.Y));
        AnotDict.AddValue('L', LineArray);
        RectArray := THPDFArrayObject.Create(nil);
        RectArray.AddNumericValue(FPage.XProjection(FLeftTop.X));
        RectArray.AddNumericValue(FPage.YProjection(FLeftTop.Y));
        RectArray.AddNumericValue(FPage.XProjection(FRightBottom.X));
        RectArray.AddNumericValue(FPage.YProjection(FRightBottom.Y));
        AnotDict.AddValue('Rect', RectArray);
        AnotDict.AddValue('P', FPage.PageObj);
      end;
    asSquare:
      begin
        AnotDict.AddNameValue('Subtype', 'Square');
        AnotDict.AddStringValue('Contents', FContents);
        RectArray := THPDFArrayObject.Create(nil);
        RectArray.AddNumericValue(FPage.XProjection(FLeftTop.X));
        RectArray.AddNumericValue(FPage.YProjection(FLeftTop.Y));
        RectArray.AddNumericValue(FPage.XProjection(FRightBottom.X));
        RectArray.AddNumericValue(FPage.YProjection(FRightBottom.Y));
        AnotDict.AddValue('Rect', RectArray);
        AnotDict.AddValue('P', FPage.PageObj);
      end;
    asCircle:
      begin
        AnotDict.AddNameValue('Subtype', 'Circle');
        AnotDict.AddStringValue('Contents', FContents);
        RectArray := THPDFArrayObject.Create(nil);
        RectArray.AddNumericValue(FPage.XProjection(FLeftTop.X));
        RectArray.AddNumericValue(FPage.YProjection(FLeftTop.Y));
        RectArray.AddNumericValue(FPage.XProjection(FRightBottom.X));
        RectArray.AddNumericValue(FPage.YProjection(FRightBottom.Y));
        AnotDict.AddValue('Rect', RectArray);
        AnotDict.AddValue('P', FPage.PageObj);
      end;
    asStamp:
      begin
        AnotDict.AddNameValue('Subtype', 'Stamp');
        AnotDict.AddStringValue('Contents', FContents);
        AnotDict.AddNameValue('Name', FTypeState);
        RectArray := THPDFArrayObject.Create(nil);
        RectArray.AddNumericValue(FPage.XProjection(FLeftTop.X));
        RectArray.AddNumericValue(FPage.YProjection(FLeftTop.Y));
        RectArray.AddNumericValue(FPage.XProjection(FRightBottom.X));
        RectArray.AddNumericValue(FPage.YProjection(FRightBottom.Y));
        AnotDict.AddValue('Rect', RectArray);
        AnotDict.AddValue('P', FPage.PageObj);
      end;
    asFileAttachment:
      begin
        MFS := TFileStream.Create(String(FName), fmOpenRead);
        try
          MFS.Position := 0;
          Inc(FPage.FParent.FMaxObjNum);
          AttS := THPDFStreamObject.Create(nil);
          AttS.IsIndirect := true;
          AttS.ID.ObjectNumber := FPage.FParent.FMaxObjNum;
          FPage.FParent.IndirectObjects.Add(AttS);
          AttS.Stream.CopyFrom(MFS, 0);
          FilSpec := FPage.FParent.CreateIndirectDictionary;
          FilSpec.AddNameValue('Type', 'Filespec');
          FilSpec.AddStringValue('F', AnsiString(ExtractFileName(String(FName))));
          FilSpec.AddValue('EF', AttS);
          AnotDict.AddNameValue('Subtype', 'FileAttachment');
          AnotDict.AddStringValue('Contents', FContents);
          RectArray := THPDFArrayObject.Create(nil);
          RectArray.AddNumericValue(FPage.XProjection(FLeftTop.X));
          RectArray.AddNumericValue(FPage.YProjection(FLeftTop.Y));
          RectArray.AddNumericValue(FPage.XProjection(FRightBottom.X));
          RectArray.AddNumericValue(FPage.YProjection(FRightBottom.Y));
          AnotDict.AddValue('Rect', RectArray);
          AnotDict.AddValue('P', FPage.PageObj);
          AnotDict.AddValue('FS', FilSpec);
        finally
          MFS.Free;
        end;
      end;
    asSound:
      begin
        MFS := TFileStream.Create(String(FName), fmOpenRead);
        try
          MFS.Position := 0;
          Inc(FPage.FParent.FMaxObjNum);
          AttS := THPDFStreamObject.Create(nil);
          AttS.IsIndirect := true;
          AttS.ID.ObjectNumber := FPage.FParent.FMaxObjNum;
          FPage.FParent.IndirectObjects.Add(AttS);
          AttS.Stream.CopyFrom(MFS, 0);
          AttS.Dictionary.AddNameValue('Type', 'Sound');
          AttS.Dictionary.AddNumericValue('R', 22050);
          AttS.Dictionary.AddNumericValue('B', 16);
          AttS.Dictionary.AddNumericValue('C', 2);
          AttS.Dictionary.AddNameValue('E', 'Signed');
          AnotDict.AddNameValue('Subtype', 'Sound');
          AnotDict.AddStringValue('Contents', FContents);
          RectArray := THPDFArrayObject.Create(nil);
          RectArray.AddNumericValue(FPage.XProjection(FLeftTop.X));
          RectArray.AddNumericValue(FPage.YProjection(FLeftTop.Y));
          RectArray.AddNumericValue(FPage.XProjection(FRightBottom.X));
          RectArray.AddNumericValue(FPage.YProjection(FRightBottom.Y));
          AnotDict.AddValue('Rect', RectArray);
          AnotDict.AddValue('P', FPage.PageObj);
          AnotDict.AddValue('Sound', AttS);
        finally
          MFS.Free;
        end;
      end;
    asMovie:
      begin
        FilSpec := FPage.FParent.CreateIndirectDictionary;
        FilSpec.AddNameValue('Type', 'Filespec');
        FilSpec.AddStringValue('F', AnsiString(ConvertFileName(String(FName))));
        AnotDict.AddNameValue('Subtype', 'Movie');
        RectArray := THPDFArrayObject.Create(nil);
        RectArray.AddNumericValue(FPage.XProjection(FLeftTop.X));
        RectArray.AddNumericValue(FPage.YProjection(FLeftTop.Y));
        RectArray.AddNumericValue(FPage.XProjection(FRightBottom.X));
        RectArray.AddNumericValue(FPage.YProjection(FRightBottom.Y));
        AnotDict.AddValue('Rect', RectArray);
        AnotDict.AddBooleanValue('A', true);
        AnotDict.AddValue('FS', FilSpec);
        MediaF := FPage.FParent.CreateIndirectDictionary;
        MediaF.AddValue('F', FilSpec);
        AspectArray := THPDFArrayObject.Create(nil);
        AspectArray.AddNumericValue(352);
        AspectArray.AddNumericValue(18);
        MediaF.AddBooleanValue('Poster', true);
        MediaF.AddValue('Aspect', AspectArray);
        AnotDict.AddValue('Movie', MediaF);
      end;
  end;
  AnotDict.AddValue('C', ColorArray);
  FPage.FAnnotsObj.AddObject(AnotDict);
end;

{ THotPDF }

constructor THotPDF.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  DocScale := 1;
  FVPChanged := false;
  FMemStream := false;
  FIsLoaded := false;
  FDocStarted := false;
  FProgress := false;
  FAutoLaunch := true;
  FPageMode := pmUseNone;
  FCompressionMethod := cmNone;
  FStandardFontEmulation := true;
  FFontEmbedding := true;
  FResolution := 72; //72
  FJpegQuality := 100;
  FShowInfo := true;
  FastOpen := '';
  FastClose := '';
  FastWillSave := '';
  FastDidSave := '';
  FastWillPrint := '';
  FastDidPrint := '';
  FNEmbeddedFonts := TStringList.Create;
  FAuthor := 'losLab';
  FTitle := 'HotPDF Document';
  FKeywords := 'Delphi PDF Component';
  FCreationDate := Now;
  IndirectObjects := TList.Create;
  FCurrentPage := nil;
  FCurrentPageNum := -1;
  FCHandle := GetDC(0 );
  SamCorrel := false;
  FParas := nil;
  FParaLen := 0;
  FActivePara := 0;
  FAutoAddPage := true;
  FRevision := 2;
  FVersion := pdf14;
end;

destructor THotPDF.Destroy;
begin
  FXref := nil;
  PageArr := nil;
  FNEmbeddedFonts.Free;
  CloseIndirectObjects;
  ReleaseDC (0, FCHandle );
  inherited;
end;

function THotPDF.FontIsEmbedded(FontNm: AnsiString): boolean;
var
  I: Integer;
begin
  result := true;
  for I := 0 to NotEmbeddedFonts.Count - 1 do
  begin
    if (NotEmbeddedFonts.Strings[I] = String(FontNm)) then
    begin
      result := false;
      Exit;
    end;
  end;
end;

procedure THotPDF.AddDocumentAttachment(FileName: TFileName; Description: AnsiString);
var
  EFNamesInd, NmInd, EFSpecInd: Integer;
  EFSpec: THPDFDictionaryObject;
  EFArr: THPDFArrayObject;
  FS: TFileStream;
  EF: THPDFDictionaryObject;
  FSpec: THPDFDictionaryObject;
  EFNamesObj, EFSpecObj, NamesObj: THPDFObject;
  ParamsObj: THPDFDictionaryObject;
  AttachStream: THPDFStreamObject;
  RootDictObj: THPDFDictionaryObject;
  NamesDict: THPDFDictionaryObject;
begin
  RootDictObj := THPDFDictionaryObject(IndirectObjects.Items[FRootIndex]);
  NmInd := RootDictObj.FindValue('Names');
  if (NmInd >= 0) then
  begin
    NamesObj := RootDictObj.GetIndexedItem(NmInd);
    if (NamesObj.ObjectType = otLink) then
    begin
      NamesDict := THPDFDictionaryObject(GetObjectByLink(THPDFLink(NamesObj)));
    end
    else
    begin
      NamesDict := THPDFDictionaryObject(NamesObj);
    end;
  end
  else
  begin
    NamesDict := CreateIndirectDictionary;
    RootDictObj.AddValue('Names', NamesDict);
  end;
  FSpec := CreateIndirectDictionary;
  EFSpecInd := NamesDict.FindValue('EmbeddedFiles');
  if (EFSpecInd >= 0) then
  begin
    EFSpecObj := NamesDict.GetIndexedItem(EFSpecInd);
    if (EFSpecObj.ObjectType = otLink) then
    begin
      EFSpec := THPDFDictionaryObject(GetObjectByLink(THPDFLink(EFSpecObj)));
    end
    else
    begin
      EFSpec := THPDFDictionaryObject(EFSpecObj);
    end;
  end
  else
  begin
    EFSpec := CreateIndirectDictionary;
  end;
  EFNamesInd := EFSpec.FindValue('Names');
  if (EFNamesInd >= 0) then
  begin
    EFNamesObj := EFSpec.GetIndexedItem(EFNamesInd);
    if (EFNamesObj.ObjectType = otLink) then
    begin
      EFArr := THPDFArrayObject(GetObjectByLink(THPDFLink(EFNamesObj)));
    end
    else
    begin
      EFArr := THPDFArrayObject(EFNamesObj);
    end;
  end
  else
  begin
    EFArr := THPDFArrayObject.Create(nil);
    EFSpec.AddValue('Names', EFArr);
  end;
  EFArr.AddStringValue(AnsiString(ExtractFileName(String(FileName))));
  EFArr.AddObject(FSpec);
  NamesDict.AddValue('EmbeddedFiles', EFSpec);
  FS := TFileStream.Create(FileName, fmOpenRead);
  try
    Inc(FMaxObjNum);
    AttachStream := THPDFStreamObject.Create(nil);
    AttachStream.IsIndirect := true;
    AttachStream.ID.ObjectNumber := FMaxObjNum;
    IndirectObjects.Add(AttachStream);
    AttachStream.Dictionary.AddNumericValue('DL', FS.Size);
    ParamsObj := THPDFDictionaryObject.Create(nil);
    ParamsObj.AddStringValue('CreationDate', _DateTimeToPdfDate(Now));
    ParamsObj.AddStringValue('ModDate', _DateTimeToPdfDate(Now));
    ParamsObj.AddNumericValue('Size', FS.Size);
    AttachStream.Dictionary.AddValue('Params', ParamsObj);
    AttachStream.Stream.CopyFrom(FS, 0);
  finally
    FS.Free;
  end;
  FSpec.AddStringValue('F', AnsiString(ExtractFileName(String(FileName))));
  FSpec.AddStringValue('Desc', Description);
  FSpec.AddNameValue('Type', 'Filespec');
  EF := THPDFDictionaryObject.Create(nil);
  EF.AddValue('F', AttachStream);
  FSpec.AddValue('EF', EF);
end;

function THotPDF.AddPage: Integer;
var
  BlockInd: Integer;
  PagesObj: THPDFDictionaryObject;
  PageObj: THPDFDictionaryObject;
  ResourObj: THPDFDictionaryObject;
  PXObObj: THPDFDictionaryObject;
  PageContObj: THPDFStreamObject;
  MediaBoxObj: THPDFArrayObject;
  ProcSetObj: THPDFArrayObject;
  DecodeParm: THPDFArrayObject;
  FilterVal: THPDFArrayObject;

begin
  Inc(FPagesCount);
  PagesObj := THPDFDictionaryObject(IndirectObjects.Items[FPagesIndex]);
  BlockInd := PagesObj.FindValue('Count');
  if (BlockInd >= 0) then
    THPDFNumericObject(PagesObj.GetIndexedItem(BlockInd)).Value := FPagesCount;
  PageObj := CreateIndirectDictionary;
  PageObj.AddNameValue('Type', 'Page');
  PageObj.AddValue('Parent', PagesObj);
  FParentMB[0] := 0;
  FParentMB[1] := 0;
  FParentMB[2] := 595;
  FParentMB[3] := 842;
  FIsParented := false;
  MediaBoxObj := THPDFArrayObject.Create(nil);
  MediaBoxObj.AddNumericValue(0);
  MediaBoxObj.AddNumericValue(0);
  MediaBoxObj.AddNumericValue(595);
  MediaBoxObj.AddNumericValue(842);
  PageObj.AddValue('MediaBox', MediaBoxObj);
  PXObObj := THPDFDictionaryObject.Create(nil);
  ProcSetObj := THPDFArrayObject.Create(nil);
  ProcSetObj.AddNameValue('PDF');
  ProcSetObj.AddNameValue('Text');
  ProcSetObj.AddNameValue('ImageC');
  ResourObj := THPDFDictionaryObject.Create(nil);
  PageObj.AddValue('Resources', ResourObj);
  ResourObj.AddValue('ProcSet', ProcSetObj);
  ResourObj.AddValue('XObject', PXObObj);
  PageContObj := THPDFStreamObject.Create(nil);
  Inc(FMaxObjNum);
  PageContObj.IsIndirect := true;
  PageContObj.ID.ObjectNumber := FMaxObjNum;
  IndirectObjects.Add(PageContObj);
  PageObj.AddValue('Contents', PageContObj);
  PageContObj.Dictionary.AddNumericValue('Length', 0);
  FilterVal := THPDFArrayObject.Create(nil);
  PageContObj.Dictionary.AddValue('Filter', FilterVal);
  DecodeParm := THPDFArrayObject.Create(nil);
  PageContObj.Dictionary.AddValue('DecodeParms', DecodeParm);
  PageArrPosition := FPagesCount - 1;
  SetLength(PageArr, FPagesCount);
  PageArr[PageArrPosition].PageObj := PageObj;
  PageArr[PageArrPosition].PageLink.ObjectNumber := PageObj.ID.ObjectNumber;
  PageArr[PageArrPosition].PageLink.GenerationNumber := PageObj.ID.GenerationNumber;
  Inc(PageArrPosition);
  SetCurrentPageNum(FPagesCount - 1);
  FCurrentPage.Size := psA4;
  FCurrentPage.SetFont('Arial', [], 10);
  result := (FPagesCount - 1);
end;

procedure THotPDF.AddTiffFromFile(FileName: TFileName; Compression: THPDFImageCompressionType; PaintType: TTIFFPaintType);
var
  ImIndex: Integer;
  TiffImage: PTIFF;
  ImPageWidth, ImPageHeight: Cardinal;
  PageBitmap: TBitmap;

  procedure ShowTiffImage;
  var
    SCLineStr: PCardn;
    ImBuff: PCardn;
    ScanLen: Cardinal;
    NewBMLength: Cardinal;
    ImCont, ImRow, ImPos: Cardinal;
    PLConfig, ImageLength: Cardinal;
  const
    PrintTempHeight = 1200;
  begin
{$IFDEF LegacyC}
    TIFFGetField(TiffImage, TIFFTAG_IMAGEWIDTH, ImPageWidth);
    TIFFGetField(TiffImage, TIFFTAG_IMAGELENGTH, ImPageHeight);
{$ELSE}
    TIFFGetField(TiffImage, TIFFTAG_IMAGEWIDTH, @ImPageWidth);
    TIFFGetField(TiffImage, TIFFTAG_IMAGELENGTH, @ImPageHeight);
{$ENDIF}
    if PaintType = tptResizePage then
    begin
      CurrentPage.Width := ImPageWidth;
      CurrentPage.Height := ImPageHeight;
    end;
    PageBitmap.Width := ImPageWidth;
    if ImPageHeight < PrintTempHeight then
    begin
      PageBitmap.Height := ImPageHeight;
      TIFFReadRGBAImage(TiffImage, ImPageWidth, ImPageHeight, PageBitmap.Scanline[ImPageHeight - 1], 0);
      TIFFReadRGBAImageSwapRB(ImPageWidth, ImPageHeight, PageBitmap.Scanline[ImPageHeight - 1]);
      if (Compression > icJpeg) then
      begin
        PageBitmap.PixelFormat := pf4bit;
        PageBitmap.Monochrome := true;
        PageBitmap.PixelFormat := pf1bit;
      end
      else
        PageBitmap.PixelFormat := pf24bit;
      ImIndex := AddImageFromTIFF(PageBitmap, Compression);
      CurrentPage.ShowImage(ImIndex, 0, 0, CurrentPage.Width, CurrentPage.Height, 0);
    end
    else
    begin
      ImRow := 0;
      ImCont := 0;

{$IFDEF LegacyC}
      TIFFGetField(TiffImage, TIFFTAG_IMAGELENGTH, ImageLength);
      TIFFGetField(TiffImage, TIFFTAG_PLANARCONFIG, PLConfig);
{$ELSE}
      TIFFGetField(TiffImage, TIFFTAG_IMAGELENGTH, @ImageLength);
      TIFFGetField(TiffImage, TIFFTAG_PLANARCONFIG, @PLConfig);
{$ENDIF}
      ImBuff := (PCardn(_TIFFmalloc(ImPageWidth * ImPageHeight * sizeof(Cardinal))));
      try
        TIFFReadRGBAImage(TiffImage, ImPageWidth, ImPageHeight, ImBuff, 0);
        while (ImCont < (ImageLength - 1)) do
        begin
          if ((ImCont + PrintTempHeight) < ImageLength - 1) then
            NewBMLength := PrintTempHeight
          else
            NewBMLength := (ImageLength - ImCont);
          PageBitmap.PixelFormat := pf32bit;
          PageBitmap.Height := NewBMLength;
          ImPos := ImCont;
          while ((ImPos - ImCont) < NewBMLength) do
          begin
            ScanLen := (ImPageWidth * (ImPageHeight - (ImPos + 1)));
            SCLineStr := ImBuff;
            Inc(SCLineStr, ScanLen);
            MoveMemory(PageBitmap.Scanline[ImPos - ImCont], SCLineStr, ImPageWidth * sizeof(Cardinal));
            Inc(ImPos);
          end;
          if (Compression > icJpeg) then
          begin
            PageBitmap.PixelFormat := pf4bit;
            PageBitmap.Monochrome := true;
            PageBitmap.PixelFormat := pf1bit;
          end
          else
            PageBitmap.PixelFormat := pf24bit;
          ImIndex := AddImageFromTIFF(PageBitmap, Compression);
          if PaintType = tptResizePage then
          begin
            CurrentPage.ShowImage(ImIndex, 0, ImRow, ImPageWidth, NewBMLength, 0);
            ImRow := ImRow + (NewBMLength - 1);
          end
          else
          begin
            CurrentPage.ShowImage(ImIndex, 0, ImRow, CurrentPage.Width, NewBMLength * (CurrentPage.Width / ImPageWidth), 0);
            ImRow := ImRow + trunc((NewBMLength * (CurrentPage.Width / ImPageWidth)) - 1);
          end;
          ImCont := ImCont + (NewBMLength - 1);
        end;
      finally
        _TIFFfree(ImBuff);
      end;
    end;
  end;

begin
  TiffImage := TIFFOpen(AnsiString(FileName), 'r');
  try
    if (TiffImage = nil) then
      raise exception.Create('Invalid tiff file name');
    PageBitmap := TBitmap.Create;
    try
      PageBitmap.PixelFormat := pf32bit;
      ShowTiffImage;
      while (TIFFReadDirectory(TiffImage) > 0) do
      begin
        AddPage;
        PageBitmap.PixelFormat := pf32bit;
        ShowTiffImage;
      end;
    finally
      PageBitmap.Free;
    end;
  finally
    TIFFClose(TiffImage);
  end;
end;

procedure THotPDF.SetDocImagearray(Width, Height: Integer);
begin
  SetLength(FSizes, FCurrentImageIndex + 1);
  FSizes[FCurrentImageIndex].Width := Width;
  FSizes[FCurrentImageIndex].heigh := Height;
end;

procedure THotPDF.SetResolution(const Value: Integer);
begin
  DocScale := Value / 72; //72
  FResolution := Value;
end;

procedure THotPDF.SetViewerPreferences(const Value: THPDFViewerPreferences);
begin
  FVPChanged := true;
  FViewerPreference := Value;
end;

function THotPDF.AddImageFromTIFF(Image: TBitmap; Compression: THPDFImageCompressionType): Integer;
begin
  result := MapImage(Image, Compression, false, -1);
end;

function THotPDF.LoadDocHeader: AnsiString;
begin
  FInStream.Position := 0;
  result := LoadDocString;
end;

procedure THotPDF.PadTrunc(S: Pointer; SL: Integer; D: Pointer);
var
  I, J: Integer;
begin
  I := 0;
  while ((I < 32) and (I < SL)) do
  begin
    PByteArray(D)[I] := PByteArray(S)[I];
    Inc(I);
  end;
  J := 1;
  while (I < 32) do
  begin
    PByteArray(D)[I] := PadStr[J];
    Inc(I);
    Inc(J);
  end;
end;

procedure THotPDF.CreateKeys;
var
  RCData, RTCData: TRC4Data;
  ID1Supl, ID2Supl, ID3Supl: AnsiString;
  CryptoStr: AnsiString;
  CryptContext: MD5Context;
  TmpDigest, Digest: MD5Digest;
  TmpKeyLength: Integer;
  Pass: array[0..31] of byte;
  I, h, ol, ul: Integer;
begin
  ol := 0;
  ul := 0;
  if (FOwnerPassword <> '') then
    ol := Length(FOwnerPassword);
  if (FUserPassword <> '') then
    ul := Length(FUserPassword);
  if (ol = 0) then
    PadTrunc(@FUserPassword[1], ul, @Pass[0])
  else
    PadTrunc(@FOwnerPassword[1], ol, @Pass[0]);
  Digest := MD5String(@Pass[0], 32);
  if (FRevision = 3) then
  begin
    for I := 1 to 50 do
      Digest := MD5String(@Digest, 16);
    TmpKeyLength := 16;
  end
  else
  begin
    TmpKeyLength := 5;
  end;
  RC4Init(RCData, @Digest, TmpKeyLength);
  PadTrunc(@FUserPassword[1], ul, @Pass[0]);
  SetLength(OwUsPass, 32);
  RC4Crypt(RCData, @Pass[0], @OwUsPass[1], 32);
  if (FRevision = 3) then
    for I := 1 to 19 do
    begin
      for h := 0 to 15 do
        TmpDigest[h] := Digest[h] xor I;
      RC4Init(RCData, @TmpDigest, TmpKeyLength);
      RC4Crypt(RCData, @OwUsPass[1], @OwUsPass[1], 32);
    end;
  ID2Supl := OwUsPass;
  SetLength(ID1Supl, 32);
  PadTrunc(@FUserPassword[1], ul, @ID1Supl[1]);
  MD5Init(CryptContext);
  MD5Update(CryptContext, @ID1Supl[1], 32);
  MD5Update(CryptContext, @ID2Supl[1], 32);
  MD5Update(CryptContext, PAnsiChar(@ProtectFlags), 4);
  ID3Supl := '';
  for I := 1 to 16 do
    ID3Supl := ID3Supl + AnsiChar(chr(StrToInt('$' + String(DocID[i shl 1 - 1]) + String(DocID[i shl 1]))));
  MD5Update(CryptContext, @ID3Supl[1], 16);
  MD5Final(CryptContext, Digest);
  if (FRevision = 3) then
  begin
    for I := 1 to 50 do
    begin
      Digest := MD5String(@Digest, 16);
    end;
  end;
  Move(Digest, FCurrentKey, TmpKeyLength);
  if (FRevision = 2) then
  begin
    RC4Init(RTCData, @FCurrentKey, 5);
    RC4Crypt(RTCData, @PadStr, @Pass, 32);
    UsPassStr := '';
    for I := 1 to 32 do
      UsPassStr := UsPassStr + AnsiChar(chr(Pass[I - 1]));
  end
  else
  begin
    MD5Init(CryptContext);
    MD5Update(CryptContext, @PadStr, 32);
    CryptoStr := '';
    for I := 1 to 16 do
      CryptoStr := CryptoStr + AnsiChar(chr(StrToInt('$' + String(DocID[i shl 1 - 1]) + String(DocID[i shl 1]))));
    MD5Update(CryptContext, @CryptoStr[1], 16);
    MD5Final(CryptContext, Digest);
    RC4Init(RTCData, @FCurrentKey[0], 16);
    RC4Crypt(RTCData, @Digest, @Digest, 16);
    for I := 1 to 19 do
    begin
      for h := 0 to 15 do
        TmpDigest[h] := FCurrentKey[h] xor I;
      RC4Init(RTCData, @TmpDigest, 16);
      RC4Crypt(RTCData, @Digest, @Digest, 16);
    end;
    SetLength(UsPassStr, 32);
    Move(Digest, UsPassStr[1], 16);
    for I := 17 to 32 do
      UsPassStr[I] := ' ';
  end;
end;

procedure THotPDF.EnableEncrypt;
var
  CFObj: THPDFDictionaryObject;
  StdCFObj: THPDFDictionaryObject;
  EncryptObj: THPDFDictionaryObject;
begin
  EncryptObj := CreateIndirectDictionary;
  FEncryprtIndex := FMaxObjNum - 1;
  EncryptObj.AddNameValue('Filter', 'Standard');
  if FCryptKeyType = k40 then
  begin
    ProtectFlags := -64;
    EncryptObj.AddNumericValue('V', 1);
    EncryptObj.AddNumericValue('R', 2);
    if prPrint in FProtectOption then
      ProtectFlags := ProtectFlags or 4;
    if prModifyStructure in FProtectOption then
      ProtectFlags := ProtectFlags or 8;
    if prInformationCopy in FProtectOption then
      ProtectFlags := ProtectFlags or 16;
    if prEditAnnotations in FProtectOption then
      ProtectFlags := ProtectFlags or 32;
  end
  else
  begin
    ProtectFlags := -3904;
    if FCryptKeyType = k128 then
    begin
      EncryptObj.AddNumericValue('V', 2);
      EncryptObj.AddNumericValue('R', 3);
    end
    else
    begin
      EncryptObj.AddNameValue('StmF', 'StdCF');
      EncryptObj.AddNumericValue('V', 4);
      EncryptObj.AddNumericValue('R', 4);
      StdCFObj := THPDFDictionaryObject.Create(nil);
      StdCFObj.AddNumericValue('Length', 16);
      StdCFObj.AddNameValue('CFM', 'AESV2');
      StdCFObj.AddNameValue('AuthEvent', 'DocOpen');
      CFObj := THPDFDictionaryObject.Create(nil);
      CFObj.AddValue('StdCF', StdCFObj);
      EncryptObj.AddValue('CF', CFObj);
      EncryptObj.AddNameValue('StrF', 'StdCF');
    end;
    EncryptObj.AddNumericValue('Length', 128);
    if prPrint in FProtectOption then
      ProtectFlags := ProtectFlags or 4;
    if prModifyStructure in FProtectOption then
      ProtectFlags := ProtectFlags or 8;
    if prInformationCopy in FProtectOption then
      ProtectFlags := ProtectFlags or 16;
    if prEditAnnotations in FProtectOption then
      ProtectFlags := ProtectFlags or 32;
    if prFillAnnotations in FProtectOption then
      ProtectFlags := ProtectFlags or 256;
    if prExtractContent in FProtectOption then
      ProtectFlags := ProtectFlags or 512;
    if prAssemble in FProtectOption then
      ProtectFlags := ProtectFlags or 1024;
    if prPrint12bit in FProtectOption then
      ProtectFlags := ProtectFlags or 2048;
  end;
  CreateKeys;
  EncryptObj.AddNumericValue('P', ProtectFlags);
  EncryptObj.AddStringValue('O', OwUsPass);
  EncryptObj.AddStringValue('U', UsPassStr);
end;

function THotPDF.CrptStr(Data: AnsiString; Password: MD5Digest; PassLength: THPDFKeyType; ObjID: Integer): AnsiString;
var
  I: Integer;
  RC4Len, MD5Len: Integer;
  CrStr: AnsiString;
  RCKey: TRC4Data;
  Digest: MD5Digest;
  CCover: THPDFContCover;
  FCRLen: Integer;
  CryptKey: array[0..20] of Byte;
  VECT: array[0..15] of Byte;
  EncData: AnsiString;
begin
  if Data = '' then
  begin
    result := Data;
    Exit;
  end;
  CrStr := Data;
  FillChar(CryptKey, 21, 0);
  if PassLength = k40 then
  begin
    Move(Password, CryptKey, 5);
    Move(ObjID, CryptKey[5], 3);
    MD5Len := 10;
    RC4Len := 10;
  end
  else
  begin
    Move(Password, CryptKey, 16);
    Move(ObjID, CryptKey[16], 3);
    MD5Len := 21;
    RC4Len := 16;
  end;
  Digest := MD5String(@CryptKey[0], MD5Len);
  if PassLength = aes128 then
  begin
    Exit;
    CCover := THPDFContCover.Create;
    try
      FCRLen := (Length(Data) div 16) * 16;
      //CCover.InitAlg(@Digest[0], 16);
      for I := 0 to 15 do
      begin
        VECT[I] := byte(Random(220) + 33);
      end;
      EncData := '';
      for I := 0 to FCRLen - 1 do
      begin
        EncData := EncData + ' ';
      end;
      // CCover.DelimitAlg(@DataBlock[0], FCRLen, @EncData[1], @VECT[0]);
    finally
      CCover.Free;
    end;
    CrStr := '';
    for I := 0 to 15 do
    begin
      CrStr := CrStr + AnsiChar(chr(VECT[I]));
    end;
    result := CrStr + EncData;
  end
  else
  begin
    RC4Init(RCKey, @Digest, RC4Len);
    RC4Crypt(RCKey, @CrStr[1], @CrStr[1], Length(CrStr));
    result := CrStr;
  end;
end;

procedure THotPDF.CrptStrm(Data: TMemoryStream; Password: MD5Digest;
  PassLength: THPDFKeyType; ObjID: Integer);
var
  RCKey: TRC4Data;
  CryptKey: array[0..20] of Byte;
  Digest: MD5Digest;
  RC4Len, MD5Len: Integer;
begin
  if Data.Size = 0 then
    Exit;
  FillChar(CryptKey, 21, 0);
  if PassLength = k40 then
  begin
    Move(Password, CryptKey, 5);
    Move(ObjID, CryptKey[5], 3);
    MD5Len := 10;
    RC4Len := 10;
  end
  else
  begin
    Move(Password, CryptKey, 16);
    Move(ObjID, CryptKey[16], 3);
    MD5Len := 21;
    RC4Len := 16;
  end;
  Digest := MD5String(@CryptKey[0], MD5Len);
  RC4Init(RCKey, @Digest, RC4Len);
  RC4Crypt(RCKey, Data.Memory, Data.Memory, Data.Size);
end;

function THotPDF.LoadIsLinearized: boolean;
var
  LinearVal: AnsiString;
  DocStr: AnsiString;
  DocPos: Integer;
  LinearPos: Integer;
begin
  result := false;
  DocPos := FInStream.Position;
  DocStr := LoadDocString;
  while Pos('obj', LowerCase(String(DocStr))) = 0 do
    DocStr := LoadDocString;
  while Pos('endobj', LowerCase(String(DocStr))) = 0 do
  begin
    DocStr := LoadDocString;
    LinearPos := Pos('linearized', LowerCase(String(DocStr)));
    if LinearPos > 0 then
    begin
      LinearPos := LinearPos + 10;
      while DocStr[LinearPos] = ' ' do
        Inc(LinearPos);
      LinearVal := '';
      while (((DocStr[LinearPos] >= '-') and (DocStr[LinearPos] <= '9'))) do
      begin
        LinearVal := LinearVal + DocStr[LinearPos];
        Inc(LinearPos);
      end;
      if (LinearVal = '1') then
        result := true;
      break;
    end;
  end;
  FInStream.Position := DocPos;
end;

procedure THotPDF.LoadXrefArray;
var
  PSkLine: Pointer;
  SkLine: AnsiString;
  DocLine: AnsiString;
  PrevPos: Integer;

  function GetFirstNunmer(NumStr: AnsiString; Index: Integer): AnsiString;
  var
    StrLen: Integer;
    StrIndex: Integer;
  begin
    StrLen := Length(NumStr);
    StrIndex := Index;
    result := '';
    while (StrIndex <= StrLen) do
    begin
      if ((NumStr[StrIndex] >= '0') and (NumStr[StrIndex] <= '9'))
        then
        result := result + NumStr[StrIndex]
      else
        break;
      Inc(StrIndex);
    end;
  end;

begin
  PrevPos := FInStream.Position;
  SkLine := '                    ';
  PSkLine := @SkLine[1];
  while (Pos('trailer', LowerCase(String(SkLine))) = 0) do
  begin
    PrevPos := FInStream.Position;
    DocLine := LoadDocString;
    while ((Pos('n', String(DocLine)) = 0) and (Pos('trailer', LowerCase(String(DocLine))) = 0)) do
    begin
      PrevPos := FInStream.Position;
      DocLine := LoadDocString;
    end;
    FInStream.Position := PrevPos;
    FInStream.Read(PSkLine^, 20);
    if SkLine[18] = 'n' then
    begin
      Inc(FXrefLen);
      SetLength(FXref, FXrefLen);
      FXref[FXrefLen - 1].Offset := StrToInt(Copy(String(SkLine), 1, 10));
    end;
  end;
  FInStream.Position := PrevPos;
end;

function THotPDF.GetOutlineRoot: THPDFDocOutlineObject;
begin
  if (FOutlineRoot = nil) then
  begin
    FOutlineRoot := THPDFDocOutlineObject.Create;
    FOutlineRoot.Init(Self);
  end;
  result := FOutlineRoot;
end;

function THotPDF.GetObjectNumber: THPDFObjectNumber;
var
  ObjPos: Integer;
  DelimiterPos: Integer;
  PLineStr: Pointer;
  LineStr: AnsiString;
begin
  ObjPos := FInStream.Position;
  LineStr := '            ';
  PLineStr := @LineStr[1];
  FInStream.Read(PLineStr^, 10);
  LineStr := AnsiString(Trim(String(LineStr)));
  DelimiterPos := Pos(' ', String(LineStr));
  result.ObjectNumber := StrToInt(Copy(String(LineStr), 1, DelimiterPos - 1));
  LineStr := AnsiString(Trim(Copy(String(LineStr), DelimiterPos + 1, 12 - DelimiterPos)));
  DelimiterPos := Pos(' ', String(LineStr));
  Result.GenerationNumber := StrToInt(Copy(String(LineStr), 1, DelimiterPos - 1));
  FInStream.Position := ObjPos;
end;

function THotPDF.LoadBooleanObject(ObjStream: TStream): THPDFBooleanObject;
var
  BoolObjPos: Integer;
  DocChar: AnsiChar;
  DocStr: AnsiString;
begin
  result := THPDFBooleanObject.Create(nil);
  ObjStream.Read(DocChar, 1);
  while (DocChar < 'A') do
    ObjStream.Read(DocChar, 1);
  DocStr := '';
  BoolObjPos := ObjStream.Position;
  while ((DocChar > #32) and (not (DocChar in EscapeChars))) do
  begin
    DocStr := DocStr + DocChar;
    ObjStream.Read(DocChar, 1);
  end;
  if (LowerCase(String(DocStr)) = 'true') then
  begin
    result.Value := true;
    ObjStream.Position := BoolObjPos + 3;
  end
  else
  begin
    result.Value := false;
    ObjStream.Position := BoolObjPos + 4;
  end;
end;

procedure THotPDF.CryptStream(Data: TMemoryStream; ID: Integer);
begin
  if FProtection then
    CrptStrm(Data, FCurrentKey, FCryptKeyType, ID);
end;

function THotPDF.CryptString(Data: AnsiString; ID: Integer): AnsiString;
begin
  if FProtection then
    result := CrptStr(Data, FCurrentKey, FCryptKeyType, ID)
  else
    result := Data;
end;

procedure THotPDF.SetOutputStream(const Value: TStream);
begin
  if FDocStarted then
    raise Exception.Create('Cannot set OutputStream value - document in progress.');
  FOutputStream := Value;
  FMemStream := true;
end;

function THotPDF.LoadNumericObject(ObjStream: TStream): THPDFNumericObject;
var
  WholePart: AnsiString;
  DecLen: Integer;
  DecPart: AnsiString;
  DecPointPos: Integer;
  DecLenPart: Integer;
  DocChar: AnsiChar;
  Decpt: Real;
  DocStr: AnsiString;

begin
  result := THPDFNumericObject.Create(nil);
  ObjStream.Read(DocChar, 1);
  while (DocChar < '-') do
    ObjStream.Read(DocChar, 1);
  DocStr := '';
  while ((DocChar > #32) and (not (DocChar in EscapeChars))) do
  begin
    DocStr := DocStr + DocChar;
    ObjStream.Read(DocChar, 1);
  end;
  if (DocChar in EscapeChars) then
    ObjStream.Position := ObjStream.Position - 1;
  DecPointPos := Pos('.', String(DocStr));
  if (DecPointPos <> 0) then
  begin
    WholePart := Copy(DocStr, 1, DecPointPos - 1);
    DecLenPart := Length(DocStr) - DecPointPos;
    if (DecLenPart <= 8) then
      DecPart := Copy(DocStr, DecPointPos + 1, DecLenPart)
    else
      DecPart := Copy(DocStr, DecPointPos + 1, 8);
    DecLen := Length(DecPart);
    Decpt := StrToInt(String(DecPart)) / Power(10, DecLen);
    result.Value := ABS(StrToInt(String(WholePart))) + Decpt;
    if (DocStr[1] = '-') then
      Result.Value := -Result.Value;
  end
  else
    Result.Value := StrToInt(String(DocStr));
end;

function THotPDF.MapImage(Image: TGraphic; Compression: THPDFImageCompressionType; IsMask: boolean; MaskIndex: Integer): Integer;
var
  I: Integer;
  XMLen: Integer;
  XOLink: THPDFLink;
  NewObjName: AnsiString;
  XObject: THPDFObject;
  CurrVImage: THPDFImage;
begin
  XOLink := nil;
  XMLen := Length(XImages);
  if MaskIndex > -1 then
  begin
    if MaskIndex > FCurrentImageIndex then
    begin
      raise exception.Create('Incorrect mask index.');
      Exit;
    end;
    for I := 0 to XMLen - 1 do
    begin
      if (MaskIndex = XImages[I].Index) then
      begin
        XObject := XImages[I].ImageObject;
        XOLink := THPDFLink.Create;
        XOLink.Value.ObjectNumber := XObject.ID.ObjectNumber;
        XOLink.Value.GenerationNumber := XObject.ID.GenerationNumber;
        break;
      end;
    end;
  end;
  SetDocImagearray(Image.Width, Image.Height);
  if not ((Image is TJPEGImage) or (Image is TBitmap)) then
    raise exception.Create('Unsupported image format.');
  NewObjName := 'Im' + AnsiString(IntToStr(FCurrentImageIndex));
  while (FCurrentPage.CompareResName(0, NewObjName) > -1) do
  begin
    Inc(FCurrentImageIndex);
    NewObjName := 'Im' + AnsiString(IntToStr(FCurrentImageIndex));
  end;
  case Compression of
    icJpeg: begin
        CurrVImage := THPDFImage.Create(Image, 0, IsMask, XOLink, NewObjName, FJpegQuality);
      end;
    icCCITT31: begin
        CurrVImage := THPDFImage.Create(Image, 2, IsMask, XOLink, NewObjName, FJpegQuality);
      end;
    icCCITT32: begin
        CurrVImage := THPDFImage.Create(Image, 3, IsMask, XOLink, NewObjName, FJpegQuality);
      end;
    icCCITT42: begin
        CurrVImage := THPDFImage.Create(Image, 4, IsMask, XOLink, NewObjName, FJpegQuality);
      end;
  else
    begin
      CurrVImage := THPDFImage.Create(Image, 1, IsMask, XOLink, NewObjName, FJpegQuality);
    end;
  end;
  Inc(FMaxObjNum);
  CurrVImage.IsIndirect := true;
  CurrVImage.ID.ObjectNumber := FMaxObjNum;
  Inc(XMLen);
  SetLength(XImages, XMLen);
  XImages[XMLen - 1].Index := FCurrentImageIndex;
  XImages[XMLen - 1].Name := NewObjName;
  XImages[XMLen - 1].ImageObject := CurrVImage;
  XImages[XMLen - 1].Width := Image.Width;
  XImages[XMLen - 1].Height := Image.Height;
  IndirectObjects.Add(CurrVImage);
  result := (XMLen - 1);
end;


function THotPDF.AddImage(Image: TGraphic; Compression: THPDFImageCompressionType; IsMask: boolean = False; MaskIndex: Integer = -1): Integer;

begin
  result := MapImage(Image, Compression, IsMask, MaskIndex);
end;

{$IFDEF BCB}
function THotPDF.AddImageFromFile(FileName: TFileName; Compression: THPDFImageCompressionType; IsMask: boolean = False; MaskIndex: Integer = -1): Integer;
{$ELSE}

function THotPDF.AddImage(FileName: TFileName; Compression: THPDFImageCompressionType; IsMask: boolean = False; MaskIndex: Integer = -1): Integer;
{$ENDIF}
var
  Ext: AnsiString;
  JPF: TJPEGImage;
  BMI: TBitmap;
begin
  Ext := AnsiString(ExtractFileExt(String(FileName)));
  case Compression of
    icJpeg:
      begin
        if Ext = '.bmp' then
        begin
          BMI := TBitmap.Create;
          try
            BMI.LoadFromFile(FileName);
            result := MapImage(BMI, Compression, IsMask, MaskIndex);
          finally
            BMI.Free;
          end;
        end
        else
        begin
          JPF := TJPEGImage.Create;
          try
            JPF.LoadFromFile(FileName);
            result := MapImage(JPF, Compression, IsMask, MaskIndex);
          finally
            JPF.Free;
          end;
        end;
      end;
    icCCITT31:
      begin
        BMI := TBitmap.Create;
        try
          BMI.LoadFromFile(FileName);
          result := MapImage(BMI, Compression, IsMask, MaskIndex);
        finally
          BMI.Free;
        end;
      end;
    icCCITT32:
      begin
        BMI := TBitmap.Create;
        try
          BMI.LoadFromFile(FileName);
          result := MapImage(BMI, Compression, IsMask, MaskIndex);
        finally
          BMI.Free;
        end;
      end;
    icCCITT42:
      begin
        BMI := TBitmap.Create;
        try
          BMI.LoadFromFile(FileName);
          result := MapImage(BMI, Compression, IsMask, MaskIndex);
        finally
          BMI.Free;
        end;
      end;
  else
    begin
      BMI := TBitmap.Create;
      try
        if Ext = '.jpg' then
        begin
          JPF := TJPEGImage.Create;
          try
            JPF.LoadFromFile(FileName);
            BMI.Assign(JPF);
          finally
            JPF.Free;
          end;
        end
        else
          BMI.LoadFromFile(FileName);
        result := MapImage(BMI, Compression, IsMask, MaskIndex);
      finally
        BMI.Free;
      end;
    end;
  end;
end;

function THotPDF.LoadStringObject(ObjStream: TStream): THPDFStringObject;
var
  Sequence: AnsiString;
  SeqLen: Integer;
  CloseBracket: AnsiChar;
  DocChar, PrevChar: AnsiChar;
  DocStr: AnsiString;
begin
  CloseBracket := ' ';
  result := THPDFStringObject.Create(nil);
  ObjStream.Read(DocChar, 1);
  while ((DocChar <> '<') and (DocChar <> '(')) do
    ObjStream.Read(DocChar, 1);
  if (DocChar = '<') then
  begin
    CloseBracket := '>';
    result.IsHexadecimal := true;
  end
  else if (DocChar = '(') then
  begin
    CloseBracket := ')';
    result.IsHexadecimal := false;
  end;
  DocStr := '';
  PrevChar := DocChar;
  ObjStream.Read(DocChar, 1);
  while ((DocChar <> CloseBracket) or (PrevChar = '\')) do
  begin
    DocStr := DocStr + DocChar;
    SeqLen := Length(DocStr);
    if (SeqLen > 2) then
    begin
      Sequence := Copy(DocStr, SeqLen - 2, 3);
      if (Sequence = '\\)') then
      begin
        DocStr := Copy(DocStr, 1, SeqLen - 1);
        break;
      end;
    end;
    PrevChar := DocChar;
    ObjStream.Read(DocChar, 1);
  end;
  result.Value := _UnEscapeText(DocStr);
end;

function THotPDF.LoadArrayObject(ObjStream: TStream): THPDFArrayObject;
var
  DocChar: AnsiChar;
  NextObj: THPDFObject;
  DocItemType: Integer;
  LinkObj: THPDFLink;
  NullObj: THPDFNullObject;
  BooleanObj: THPDFBooleanObject;
  NumericObj: THPDFNumericObject;
  StringObj: THPDFStringObject;
  NameObj: THPDFNameObject;
  StreamObj: THPDFStreamObject;
  ArrayObj: THPDFArrayObject;
  DictionaryObj: THPDFDictionaryObject;

  function GetArrayItemType: Integer;
  begin
    ObjStream.Read(DocChar, 1);
    while ((DocChar < #33)) do
      ObjStream.Read(DocChar, 1);
    if (DocChar = ']') then
      result := 99
    else
    begin
      ObjStream.Position := ObjStream.Position - 1;
      result := Integer(GetObjectType(ObjStream, false));
    end;
  end;

  function ReadNextObject(ItemType: Integer): THPDFObject;
  begin
    result := nil;
    case ItemType of
      0:
        begin
          BooleanObj := LoadBooleanObject(ObjStream);
          result := THPDFObject(BooleanObj);
        end;
      1:
        begin
          NumericObj := LoadNumericObject(ObjStream);
          result := THPDFObject(NumericObj);
        end;
      2:
        begin
          StringObj := LoadStringObject(ObjStream);
          result := THPDFObject(StringObj);
        end;
      3:
        begin
          NameObj := LoadNameObject(ObjStream);
          result := THPDFObject(NameObj);
        end;
      4:
        begin
          ArrayObj := LoadArrayObject(ObjStream);
          result := THPDFObject(ArrayObj);
        end;
      5:
        begin
          DictionaryObj := LoadDictionaryObject(ObjStream);
          result := THPDFObject(DictionaryObj);
        end;
      6:
        begin
          StreamObj := LoadStreamObject(ObjStream);
          result := THPDFObject(StreamObj);
        end;
      7:
        begin
          NullObj := THPDFNullObject.Create;
          ObjStream.Position := ObjStream.Position + 4;
          result := THPDFObject(NullObj);
        end;
      8:
        begin
          LinkObj := LoadLinkObject(ObjStream);
          result := THPDFObject(LinkObj);
        end;
    end;
  end;

begin
  result := THPDFArrayObject.Create(nil);
  ObjStream.Read(DocChar, 1);
  while (DocChar <> '[') do
    ObjStream.Read(DocChar, 1);
  DocItemType := GetArrayItemType;
  while (DocItemType <> 99) do
  begin
    NextObj := ReadNextObject(DocItemType);
    NextObj.FParent := result;
    result.AddObject(NextObj);
    DocItemType := GetArrayItemType;
  end;
end;

function THotPDF.LoadStreamObject(ObjStream: TStream): THPDFStreamObject;
var
  I: Integer;
  DocLine: AnsiString;
  StopPos, BodyLen: Integer;
  FilterItem, LengthItem: Integer;
  XRefObj: THPDFObject;
  LenObj: THPDFNumericObject;
  SkipLet: AnsiChar;
begin
  result := THPDFStreamObject.Create(nil);
  result.Dictionary.Free;
  result.Dictionary := LoadDictionaryObject(FInStream);
  LengthItem := result.Dictionary.FindValue('Length');
  if (LengthItem >= 0) then
  begin
    XRefObj := result.Dictionary.GetIndexedItem(LengthItem);
    if XRefObj.ObjectType = otLink then
    begin
      StopPos := FInStream.Position;
      I := 0;
      while (CompareObjectID(FXref[I].ObjNumber, THPDFLink(XRefObj).Value) <> 0)
        do Inc(I);
      FXref[I].Loaded := true;
      FInStream.Position := FXref[I].Offset;
      GetObjectType(FInStream, true);
      LenObj := LoadNumericObject(FInStream);
      BodyLen := Round(LenObj.Value);
      LenObj.IsIndirect := false;
      result.Dictionary.ReplaceValue('Length', LenObj);
      result.Length := BodyLen;
      FInStream.Position := StopPos;
    end
    else
    begin
      BodyLen := Round(THPDFNumericObject(XRefObj).Value);
      result.Length := BodyLen;
    end;
  end
  else
    raise exception.Create('Cannot load stream length');
  FilterItem := result.Dictionary.FindValue('Filter');
  DocLine := LoadDocString;
  while not (Pos('stream', LowerCase(String(DocLine))) > 0) do
    DocLine := LoadDocString;
  FInStream.Read(SkipLet, 1);
  if ((SkipLet = #13) or (SkipLet = #10)) then
    FInStream.Read(SkipLet, 1);
  if ((SkipLet <> #13) and (SkipLet <> #10)) then
    FInStream.Position := FInStream.Position - 1;
  if (FilterItem < 0) then
  begin
    TMemoryStream(result.Stream).SetSize(BodyLen);
    FInStream.Read(TMemoryStream(result.Stream).Memory^, BodyLen);
  end
  else
  begin
    TMemoryStream(result.Stream).SetSize(BodyLen);
    FInStream.Read(TMemoryStream(result.Stream).Memory^, BodyLen);
  end;
end;

function THotPDF.LoadDictionaryObject(ObjStream: TStream): THPDFDictionaryObject;
var
  DocChar: AnsiChar;
  Key: AnsiString;

  function IsEof: boolean;
  var
    DocPos: Integer;
  begin
    result := false;
    DocPos := ObjStream.Position;
    ObjStream.Read(DocChar, 1);
    while (DocChar <= #32) do
      ObjStream.Read(DocChar, 1);
    if DocChar = '>' then
    begin
      ObjStream.Read(DocChar, 1);
      if DocChar = '>' then
        result := true;
    end;
    ObjStream.Position := DocPos;
  end;

var
  ValObject: THPDFObject;
  ValKey: THPDFNameObject;
  ObjType: THPDFObjectType;
  DocPos: Integer;

begin
  result := THPDFDictionaryObject.Create(nil);
  ObjStream.Read(DocChar, 1);
  while (DocChar <> '<') do
    ObjStream.Read(DocChar, 1);
  ObjStream.Read(DocChar, 1);
  DocPos := ObjStream.Position;
  ObjStream.Read(DocChar, 1);
  while (DocChar <= #32) do
    ObjStream.Read(DocChar, 1);
  if (DocChar = '>') then
  begin
    ObjStream.Read(DocChar, 1);
    Exit;
  end
  else
    ObjStream.Position := DocPos;
  while (not (IsEof)) do
  begin
    ValKey := LoadNameObject(ObjStream);
    try
      Key := ValKey.Value;
    finally
      ValKey.Free;
    end;
    if (LowerCase(String(Key)) = 'linearized') then
      IsLinearized := true;
    ObjType := GetObjectType(FInStream, false);
    ValObject := AddTypeObject(ObjType, false);
    ValObject.FParent := result;
    ValObject.IsIndirect := false;
    result.AddValue(Key, ValObject);
  end;
  ObjStream.Read(DocChar, 1);
  while (DocChar <> '>') do
    ObjStream.Read(DocChar, 1);
  ObjStream.Read(DocChar, 1);
end;

function THotPDF.LoadNameObject(ObjStream: TStream): THPDFNameObject;
var
  DocChar: AnsiChar;
  DocStr: AnsiString;
begin
  result := THPDFNameObject.Create(nil);
  ObjStream.Read(DocChar, 1);
  while (DocChar <> '/') do
    ObjStream.Read(DocChar, 1);
  DocStr := '';
  ObjStream.Read(DocChar, 1);
  while ((not (DocChar in EscapeChars))) do
  begin
    DocStr := DocStr + DocChar;
    ObjStream.Read(DocChar, 1);
  end;
  if (DocChar <> ' ') then
    ObjStream.Position := ObjStream.Position - 1;
  result.Value := DocStr;
end;

function THotPDF.LoadLinkObject(ObjStream: TStream): THPDFLink;
var
  DocChar: AnsiChar;
  DocStr: AnsiString;
begin
  result := THPDFLink.Create;
  ObjStream.Read(DocChar, 1);
  while (DocChar <= #32) do
    ObjStream.Read(DocChar, 1);
  DocStr := '';
  while (DocChar > #32) do
  begin
    DocStr := DocStr + DocChar;
    ObjStream.Read(DocChar, 1);
  end;
  result.Value.ObjectNumber := StrToInt(Trim(String(DocStr)));
  while (DocChar <= #32) do
    ObjStream.Read(DocChar, 1);
  DocStr := '';
  while (DocChar > #32) do
  begin
    DocStr := DocStr + DocChar;
    ObjStream.Read(DocChar, 1);
  end;
  result.Value.GenerationNumber := StrToInt(Trim(String(DocStr)));
  while (LowerCase(Char(DocChar)) <> 'r') do
  begin
    ObjStream.Read(DocChar, 1);
  end;
end;

procedure THPDFPage.SetFont(FontName: AnsiString; FontStyle: TFontStyles; ASize: Single; FontCharset: TFontCharset = ANSI_CHARSET; IsVertical: boolean = false);
var
  HDescript: HDC;
  LogFont: TLogFont;
  IsInstalled: boolean;

  function InspectFont(const LogFonts: ENUMLOGFONTEX; const NewTextMetrics: TNEWTEXTMETRICEXA; FontType: DWORD; var IsInstalled: Boolean): Integer; stdcall;
  begin
    IsInstalled := (FontType = TRUETYPE_FONTTYPE);
    result := 0;
  end;
{$IFDEF UNICODE}

var
  LFName: string;
{$ENDIF}
begin
  if ((LowerCase(String(FontName)) = 'tahoma') and (fsItalic in FontStyle)) then
  begin
    FontName := 'Tahoma Italic';
  end;
  IsInstalled := false;
  FillChar(LogFont, sizeof(LogFont), 0);
  LogFont.lfCharSet := DEFAULT_CHARSET;
{$IFDEF UNICODE}
  LFName := String(FontName);
  Move(LFName[1], LogFont.lfFaceName, Length(FontName) * 2);
{$ELSE}
  Move(FontName[1], LogFont.lfFaceName, Length(FontName));
{$ENDIF}
  if FontName <> 'ZapfDingbats' then
  begin
    HDescript := GetDC(0);
    try
      EnumFontFamiliesEx(HDescript, LogFont, @InspectFont, Integer(@IsInstalled), 0);
    finally
      ReleaseDC(0, HDescript);
    end;
  end
  else
    IsInstalled := true;
  if not (IsInstalled) then
  begin
    FontName := 'Arial Unicode MS';
    FontCharset := DEFAULT_CHARSET;
  end;
  if ((CurrentFontObj.OldName <> FontName) or
    (CurrentFontObj.FontStyle <> FontStyle) or
    (CurrentFontObj.FontCharset <> FontCharset) or
    (CurrentFontObj.IsVertical <> IsVertical)) then
  begin
    if (CurrentFontObj.IsUsed) then
      StoreCurrentFont;
    CurrentFontObj.Name := FontName;
    CurrentFontObj.FontStyle := FontStyle;
    CurrentFontObj.FontCharset := FontCharset;
    CurrentFontObj.IsVertical := IsVertical;
    CurrentFontObj.IsUsed := false;
    CurrentFontObj.IsUnicode := false;
    CurrentFontObj.ClearTables;
    CurrentFontObj.ParseFontName;
    CurrentFontObj.GetFontFeatures;
  end;
  if (FSize = psUserDefined) then
    CurrentFontObj.Size := ASize
  else
  begin
    if (FParent.DocScale = 0) then
      FParent.DocScale := 1;
    CurrentFontObj.Size := (ASize / FParent.DocScale) * DPI;
  end;
end;

function THPDFPage.TextHeight(Text: AnsiString): Real;
begin
  result := CurrentFontObj.Size * CurrentFontObj.Ascent / 1000;
end;

function THotPDF.AddTypeObject(ObjType: THPDFObjectType; IsIndirect: boolean): THPDFObject;
var
  NullObj: THPDFNullObject;
  BooleanObj: THPDFBooleanObject;
  NumericObj: THPDFNumericObject;
  StringObj: THPDFStringObject;
  NameObj: THPDFNameObject;
  StreamObj: THPDFStreamObject;
  ArrayObj: THPDFArrayObject;
  DictionaryObj: THPDFDictionaryObject;
  LinkObj: THPDFLink;
begin
  result := nil;
  case ObjType of
    otNull:
      begin
        NullObj := THPDFNullObject.Create;
        if (IsIndirect) then
          IndirectObjects.Add(NullObj);
        result := THPDFObject(NullObj);
      end;
    otBoolean:
      begin
        BooleanObj := LoadBooleanObject(FInStream);
        if (IsIndirect) then
          IndirectObjects.Add(BooleanObj);
        result := THPDFObject(BooleanObj);
      end;
    otNumeric:
      begin
        NumericObj := LoadNumericObject(FInStream);
        if (IsIndirect) then
          IndirectObjects.Add(NumericObj);
        result := THPDFObject(NumericObj);
      end;
    otString:
      begin
        StringObj := LoadStringObject(FInStream);
        if (IsIndirect) then
          IndirectObjects.Add(StringObj);
        result := THPDFObject(StringObj);
      end;
    otName:
      begin
        NameObj := LoadNameObject(FInStream);
        if (IsIndirect) then
          IndirectObjects.Add(NameObj);
        result := THPDFObject(NameObj);
      end;
    otArray:
      begin
        ArrayObj := LoadArrayObject(FInStream);
        if (IsIndirect) then
          IndirectObjects.Add(ArrayObj);
        result := THPDFObject(ArrayObj);
      end;
    otDictionary:
      begin
        DictionaryObj := LoadDictionaryObject(FInStream);
        if (IsLinearized) then
        begin
          IsLinearized := false;
          DictionaryObj.Free;
        end
        else
        begin
          if (IsIndirect) then
            IndirectObjects.Add(DictionaryObj);
          result := THPDFObject(DictionaryObj);
        end;
      end;
    otStream:
      begin
        StreamObj := LoadStreamObject(FInStream);
        if (IsIndirect) then
          IndirectObjects.Add(StreamObj);
        result := THPDFObject(StreamObj);
      end;
    otLink:
      begin
        LinkObj := LoadLinkObject(FInStream);
        if (IsIndirect) then
          IndirectObjects.Add(LinkObj);
        result := THPDFObject(LinkObj);
      end;
  end;
  if (result <> nil) then
    result.IsDeleted := false;
end;

function THotPDF.GetObjectType(ObjStream: TStream; UseIndirect: boolean): THPDFObjectType;
var
  StreamBegin: Integer;
  IsStream: boolean;
  IsNum: boolean;
  StackLos: AnsiString;
  StackLen: Integer;
  DocStr: AnsiString;
  DocPos: Integer;
  DocChar: AnsiChar;
  LinePos: Integer;

  procedure PushChar(StChar: AnsiChar);
  begin
    if StackLen = 3 then
    begin
      StackLos[1] := AnsiChar(StackLos[2]);
      StackLos[2] := AnsiChar(StackLos[3]);
      StackLos[3] := AnsiChar(StChar);
    end
    else
    begin
      Inc(StackLen);
      StackLos := StackLos + StChar;
    end;
  end;

begin
  if UseIndirect then
  begin
    DocPos := ObjStream.Position;
    LinePos := ObjStream.Position;
    DocStr := LoadDocString;
    while (Pos('endobj', LowerCase(String(DocStr))) = 0) do
    begin
      StreamBegin := Pos('stream', LowerCase(String(DocStr)));
      if (StreamBegin > 0) then
      begin
        IsStream := false;
        ObjStream.Position := LinePos + StreamBegin + 5;
        ObjStream.Read(DocChar, 1);
        if (DocChar = #10) then
          IsStream := true;
        ObjStream.Read(DocChar, 1);
        if (DocChar = #10) then
          IsStream := true;
        if (IsStream) then
        begin
          result := otStream;
          ObjStream.Position := DocPos;
          Exit;
        end;
      end;
      LinePos := ObjStream.Position;
      DocStr := LoadDocString;
    end;
    ObjStream.Position := DocPos;
    StackLen := 0;
    StackLos := '';
    while (LowerCase(String(StackLos)) <> 'obj') do
    begin
      ObjStream.Read(DocChar, 1);
      PushChar(DocChar);
    end;
  end;
  IsNum := false;
  ObjStream.Read(DocChar, 1);
  while (DocChar <= #32) do
    ObjStream.Read(DocChar, 1);
  DocPos := ObjStream.Position - 1;
  while (DocChar > #32) do
  begin
    if not ((DocChar >= '-') and (DocChar <= '9')) then
    begin
      IsNum := false;
      break;
    end
    else
      IsNum := true;
    ObjStream.Read(DocChar, 1);
  end;
  if (IsNum) then
  begin
    while (DocChar <= #32) do
      ObjStream.Read(DocChar, 1);
    while (DocChar > #32) do
    begin
      if not ((DocChar >= '0') and (DocChar <= '9')) then
      begin
        IsNum := false;
        break;
      end;
      ObjStream.Read(DocChar, 1);
    end;
    if (IsNum) then
    begin
      while (DocChar <= #32) do
        ObjStream.Read(DocChar, 1);
      if not (LowerCase(Char(DocChar)) = 'r') then
        IsNum := false;
    end;
  end;
  ObjStream.Position := DocPos;
  if (IsNum) then
  begin
    result := otLink;
    Exit;
  end;
  ObjStream.Read(DocChar, 1);
  while (DocChar <= #32) do
    ObjStream.Read(DocChar, 1);
  if DocChar = '/' then
    result := otName
  else if DocChar = '(' then
    result := otString
  else if DocChar = '[' then
    result := otArray
  else if DocChar = '<' then
  begin
    ObjStream.Read(DocChar, 1);
    if DocChar = '<' then
      result := otDictionary
    else
      result := otString;
  end
  else
  begin
    DocStr := '';
    while ((DocChar > #32) and (not (DocChar in EscapeChars))) do
    begin
      DocStr := DocStr + DocChar;
      ObjStream.Read(DocChar, 1);
    end;
    if LowerCase(String(DocStr)) = 'null' then
      result := otNull
    else if ((LowerCase(String(DocStr)) = 'true') or (LowerCase(String(DocStr)) = 'false')) then
      result := otBoolean
    else
      result := otNumeric;
  end;
  ObjStream.Position := DocPos;
end;

function THotPDF.LoadDocString: AnsiString;
var
  LsSym: AnsiChar;
  LsString: AnsiString;
  DocSize: Integer;

begin
  LsString := '';
  DocSize := FInStream.Size;
  if (FInStream.Position >= (DocSize - 1)) then
    result := ''
  else
  begin
    FInStream.Read(LsSym, 1);
    while ((not((LsSym = #13) or (LsSym = #10))) and
      (FInStream.Position < DocSize)) do
    begin
      LsString := LsString + LsSym;
      FInStream.Read(LsSym, 1);
    end;
    if (FInStream.Position = DocSize) then
      LsString := LsString + LsSym;
    result := LsString;
  end;
end;

{$IFDEF BCB}

procedure THPDFPage.ShowMetafileEx(MetaFile: TMetafile);
{$ELSE}

procedure THPDFPage.ShowMetafile(MetaFile: TMetafile);
{$ENDIF}
var
  OldRes: Integer;
begin
  GStateSave;
  OldRes := FResolution;
  SetResolution(72); //72
  ShowMetafile(MetaFile, 0, 0, 1, 1);
  SetResolution(OldRes);
  GStateRestore;
end;

procedure THPDFPage.ShowMetafile(MetaFile: TMetafile; X, Y, HorScale, VertScale: Extended);
var
  PDFMeta: THPDFWmf;
  NMF: TMetafile;
  MFC: TMetafileCanvas;
  AX: Extended;
  XS, YS: Integer;
  w, h: Integer;
begin
  AX := 1;
  NMF := TMetafile.Create;
  try
    XS := GetDeviceCaps(FParent.FCHandle, HORZRES);
    YS := GetDeviceCaps(FParent.FCHandle, VERTRES);
    if (MetaFile.Height > YS) or (MetaFile.Width > XS) then
    begin
      AX := Min(YS / MetaFile.Height, XS / MetaFile.Width);
      NMF.Height := Round(MetaFile.Height * AX);
      NMF.Width := Round(MetaFile.Width * AX);
    end
    else
    begin
      NMF.Height := MetaFile.Height;
      NMF.Width := MetaFile.Width;
    end;
    w := NMF.Width;
    h := NMF.Height;
    MFC := TMetafileCanvas.Create(NMF, FParent.FCHandle);
    try
      if AX = 1 then
        MFC.Draw(0, 0, MetaFile)
      else
        MFC.StretchDraw(Rect(0, 0, w - 1, h - 1), MetaFile);
    finally
      MFC.Free;
    end;
    PDFMeta := THPDFWmf.Create(Self);
    try
      NMF.Enhanced := true;
      PDFMeta.Analyse(NMF);
    finally
      PDFMeta.Free;
    end;
  finally
    NMF.Free;
  end;
end;

function THotPDF.LoadBackDocString: AnsiString;
var
  DocChar: AnsiChar;
  DocPos, StPos: Integer;
begin
  result := '';
  FInStream.Read(DocChar, 1);
  while ((DocChar <> #13) and (DocChar <> #10) and (FInStream.Position > 1)) do
  begin
    FInStream.Position := FInStream.Position - 2;
    FInStream.Read(DocChar, 1);
  end;
  StPos := FInStream.Position;
  while (DocChar < #32) do
  begin
    FInStream.Position := FInStream.Position - 2;
    FInStream.Read(DocChar, 1);
  end;
  DocPos := FInStream.Position;
  FInStream.Position := StPos;
  result := LoadDocString;
  FInStream.Position := DocPos - 1;
end;

function THotPDF.LoadFromFile(FileName: TFileName): Integer;
var
  FSDoc: TFileStream;
begin
  FSDoc := TFileStream.Create(FileName, fmOpenRead);
  try
    result := LoadFromStream(FSDoc);
  finally
    FSDoc.Free;
  end;
end;

function THotPDF.LoadFromStream(DocStream: TStream): Integer;
var
  I, h, LocLen, BlockInd: Integer;
  PageCountObj: THPDFNumericObject;
  LinkObj: THPDFLink;
  XRefObj: THPDFObject;
  DocLine: AnsiString;
  DocHeaderStr: AnsiString;
  IsHavePrev: boolean;
  FPagesNewObj: THPDFDictionaryObject;
  FPNArray: THPDFArrayObject;
  ONCompare: THPDFObjectNumber;
  ObjType: THPDFObjectType;
  LObjTable: array of TGenByteArray;

  procedure SkipToWord(SubString: AnsiString; GoBack: boolean);
  var
    RollPos: Integer;
    LineStr: AnsiString;
  begin
    RollPos := FInStream.Position;
    if GoBack then
      LineStr := AnsiString(LowerCase(String(LoadBackDocString)))
    else
      LineStr := AnsiString(LowerCase(String(LoadDocString)));
    while (Pos(SubString, LineStr) = 0) do
    begin
      RollPos := FInStream.Position;
      if GoBack then
        LineStr := AnsiString(LowerCase(String(LoadBackDocString)))
      else
        LineStr := AnsiString(LowerCase(String(LoadDocString)));
    end;
    if (not (GoBack)) then
      FInStream.Position := RollPos + Pos(SubString, LineStr) + Length(SubString);
  end;

  function TrimChars(InStr: AnsiString): AnsiString;
  var
    StrLen: Integer;
    Index: Integer;
  begin
    Index := 1;
    StrLen := Length(InStr);
    result := '';
    while (not ((InStr[Index] >= '0') and (InStr[Index] <= '9'))) do
    begin
      Inc(Index);
      if (Index > StrLen) then
        break;
    end;
    if (not (Index > StrLen)) then
    begin
      while ((InStr[Index] >= '0') and (InStr[Index] <= '9')) do
      begin
        result := result + InStr[Index];
        Inc(Index);
        if (Index > StrLen) then
          break;
      end;
    end;
  end;

  function GetVersionNum(DocHeader: AnsiString): TPDFVersType;
  var
    TmpStr: AnsiString;
    Delpos: Integer;
  begin
    TmpStr := AnsiString(Trim(String(DocHeader)));
    Delpos := Pos('.', String(DocHeader));
    TmpStr := Copy(TmpStr, Delpos + 1, Length(TmpStr) - Delpos);
    result := TPDFVersType(StrToInt(String(TmpStr)));
  end;

var
  CCContentObj: THPDFObject;
  ParentMBObject: THPDFArrayObject;
  RootDictObj: THPDFDictionaryObject;
begin
  FMaxObjNum := 0;
  FIsLoaded := true;
  result := 0;
  if (FDocStarted) then
    raise exception.Create('Please load the document before using BeginDoc.');
  FInStream := DocStream;
  DocHeaderStr := LoadDocHeader;
  FVersion := GetVersionNum(DocHeaderStr);
  try
    FInStream.Position := FInStream.Size - 1;
    SkipToWord('%%eof', true);
    DocLine := LoadBackDocString;
    DocLine := TrimChars(DocLine);
    FInStream.Position := StrToInt(Trim(String(DocLine)));
    IsHavePrev := true;
    while (IsHavePrev) do
    begin
      LoadDocString;
      LoadXrefArray;
      SkipToWord('trailer', false);
      TrailerObj := LoadDictionaryObject(FInStream);
      try
        I := TrailerObj.FindValue('Prev');
        if (I >= 0) then
        begin
          IsHavePrev := true;
          FInStream.Position := Round(THPDFNumericObject(TrailerObj.GetIndexedItem(I)).Value);
        end
        else
          IsHavePrev := false;
        I := TrailerObj.FindValue('Root');
        if (I >= 0) then
        begin
          LinkObj := THPDFLink(TrailerObj.GetIndexedItem(I));
          FRootLink := LinkObj.Value;
        end;
        I := TrailerObj.FindValue('Info');
        if (I >= 0) then
        begin
          LinkObj := THPDFLink(TrailerObj.GetIndexedItem(I));
          FInfoLink := LinkObj.Value;
        end;
        I := TrailerObj.FindValue('Encrypt');
        if (I >= 0) then
        begin
          FIsEncrypted := true;
          Exit;
          raise exception.Create('File structure is incorrect');
          LinkObj := THPDFLink(TrailerObj.GetIndexedItem(I));
          FEncryprtLink := LinkObj.Value;
          FIsEncrypted := true;
        end
        else
          FIsEncrypted := false;
      finally
        TrailerObj.Free;
      end;
    end;
    SetLength(LObjTable, FXrefLen);
    LocLen := FXrefLen;
    for I := 0 to FXrefLen - 1 do
    begin
      if (0 < FXref[I].Offset) then
      begin
        FInStream.Position := FXref[I].Offset;
        ONCompare := GetObjectNumber;
        if (ONCompare.ObjectNumber > FMaxObjNum) then
        begin
          FMaxObjNum := ONCompare.ObjectNumber;
        end;
        if (LocLen < ONCompare.ObjectNumber) then
        begin
          SetLength(LObjTable, ONCompare.ObjectNumber);
          LocLen := ONCompare.ObjectNumber;
        end;
        if (LObjTable[ONCompare.ObjectNumber - 1][ONCompare.GenerationNumber] = 5) then
        begin
          FXref[I].ObjNumber.ObjectNumber := 0;
          FXref[I].ObjNumber.GenerationNumber := 0;
        end
        else
        begin
          LObjTable[ONCompare.ObjectNumber - 1][ONCompare.GenerationNumber] := 5;
          FXref[I].ObjNumber := ONCompare;
          FXref[I].Loaded := false;
        end;
      end;
    end;
    LObjTable := nil;
    for I := 0 to FXrefLen - 1 do
    begin
      if (not (FXref[I].Loaded)) then
      begin
        FInStream.Position := FXref[I].Offset;
        ObjType := GetObjectType(FInStream, true);
        XRefObj := AddTypeObject(ObjType, true);
        if (XRefObj <> nil) then
        begin
          XRefObj.IsIndirect := true;
          if (0 < FXref[I].ObjNumber.ObjectNumber) then
          begin
            XRefObj.ID := FXref[I].ObjNumber;
            if (CompareObjectID(XRefObj.ID, FRootLink) = 0) then
            begin
              FRootIndex := IndirectObjects.Count - 1;
              BlockInd := THPDFDictionaryObject(XRefObj).FindValue('Pages');
              if (BlockInd >= 0) then
              begin
                Linkobj := THPDFLink(THPDFDictionaryObject(XRefObj).GetIndexedItem(BlockInd));
                FPageslink := LinkObj.Value;
              end;
            end;
            if (CompareObjectID(XRefObj.ID, FInfoLink) = 0) then
              FInfoIndex := IndirectObjects.Count - 1;
            if (CompareObjectID(XRefObj.ID, FEncryprtLink) = 0) then
              FEncryprtIndex := IndirectObjects.Count - 1;
          end
          else
            XRefObj.IsDeleted := true;
        end;
      end;
    end;
    for I := 0 to (IndirectObjects.Count - 1) do
    begin
      if (CompareObjectID(THPDFObject(IndirectObjects.Items[I]).ID, FPagesLink) = 0) then
      begin
        FPagesIndex := I;
        BlockInd := THPDFDictionaryObject(IndirectObjects.Items[I]).FindValue('MediaBox');
        if (BlockInd >= 0) then
        begin
          CCContentObj := (THPDFDictionaryObject(IndirectObjects.Items[I]).GetIndexedItem(BlockInd));
          if (CCContentObj.ObjectType = otLink) then
          begin
            ParentMBObject := THPDFArrayObject(GetObjectByLink(THPDFLink(CCContentObj)));
          end
          else
            ParentMBObject := THPDFArrayObject(THPDFDictionaryObject(IndirectObjects.Items[I]).GetIndexedItem(BlockInd));
          	FParentMB[0] := THPDFNumericObject(ParentMBObject.GetIndexedItem(0)).Value;
          	FParentMB[1] := THPDFNumericObject(ParentMBObject.GetIndexedItem(1)).Value;
          	FParentMB[2] := THPDFNumericObject(ParentMBObject.GetIndexedItem(2)).Value;
          	FParentMB[3] := THPDFNumericObject(ParentMBObject.GetIndexedItem(3)).Value;
          FIsParented := true;
        end;
        BlockInd := THPDFDictionaryObject(IndirectObjects.Items[I]).FindValue('Count');
        if (BlockInd >= 0) then
        begin
          PageCountObj := THPDFNumericObject(THPDFDictionaryObject(IndirectObjects.Items[I]).GetIndexedItem(BlockInd));
          result := trunc(PageCountObj.Value);
          FPagesCount := result;
          PageArrPosition := 0;
          SetLength(PageArr, result);
          ListExtDictionary(THPDFDictionaryObject(IndirectObjects.Items[I]), FPagesLink);
          break;
        end;
      end;
    end;
    FPagesNewObj := CreateIndirectDictionary;
    FPagesNewObj.AddNameValue('Type', 'Pages');
    FPagesNewObj.AddNumericValue('Count', result);
    FPNArray := THPDFArrayObject.Create(nil);
    FPagesNewObj.AddValue('Kids', FPNArray);
    FPagesIndex := IndirectObjects.Count - 1;
    for h := 0 to (result - 1) do
    begin
      PageArr[h].PageObj.ReplaceValue('Parent', FPagesNewObj);
    end;
    RootDictObj := THPDFDictionaryObject(IndirectObjects.Items[FRootIndex]);
    RootDictObj.ReplaceValue('Pages', FPagesNewObj);
  except
    raise exception.Create('The file is damaged.');
  end;
end;

function THPDFPage.TextWidth(Text: AnsiString): Single;
var
  I: Integer;
  ch: AnsiChar;
  tmpWidth: Single;
  chv: Integer;
begin
  result := 0;
  for I := 1 to Length(Text) do
  begin
    ch := Text[I];
    chv := CurrentFontObj.GetCharWidth(Text, I);
    tmpWidth := chv * CurrentFontObj.Size / 1000;
    if FHorizontalScaling <> 100 then
      tmpWidth := tmpWidth * FHorizontalScaling / 100;
    if tmpWidth > 0 then
      tmpWidth := tmpWidth + FCharSpace
    else
      tmpWidth := 0;
    if (ch = ' ') and (FWordSpace > 0) and (I <> Length(Text)) then
      tmpWidth := tmpWidth + FWordSpace;
    result := result + tmpWidth;
  end;
  result := (result / DocScale);
end;

{ THPDFFontObj }

procedure THPDFFontObj.ClearTables;
var
  I: Integer;
begin
  ArrIndex := -1;
  SymbolTable[32] := true;
  SymbolTable[33] := true;
  for I := 34 to 255 do
    SymbolTable[I] := false;
end;

function THPDFFontObj.GetCharWidth(AText: AnsiString; APos: Integer): Integer;
var
  ChCode: byte;
begin
  ChCode := Ord(AText[APos]);
  if not IsMonospaced then
    Result := ABCArray[ChCode].abcA + Integer(ABCArray[ChCode].abcB) + ABCArray[ChCode].abcC
  else
    result := ABCArray[0].abcA + Integer(ABCArray[0].abcB) + ABCArray[0].abcC;
end;

procedure THPDFFontObj.CopyFontFetures(InFnt: THPDFFontObj);
var
  I: Integer;
  UnicTblLen: Integer;
begin
  Move(InFnt.OutTextM, OutTextM, sizeof(OUTLINETEXTMETRIC));
  for I := 0 to 255 do
  begin
    ABCArray[I].abcA := InFnt.ABCArray[I].abcA;
    ABCArray[I].abcB := InFnt.ABCArray[I].abcB;
    ABCArray[I].abcC := InFnt.ABCArray[I].abcC;
  end;
  for I := 32 to 255 do
  begin
    SymbolTable[I] := InFnt.SymbolTable[I];
  end;
  if (IsUnicode) then
  begin
    UnicTblLen := Length(InFnt.UnicodeTable);
    SetLength(UnicodeTable, UnicTblLen);
    for I := 0 to (UnicTblLen - 1) do
    begin
      UnicodeTable[I].CharCode := InFnt.UnicodeTable[I].CharCode;
      UnicodeTable[I].Index := InFnt.UnicodeTable[I].Index;
    end;
    UnicTblLen := Length(InFnt.Symbols);
    SetLength(Symbols, UnicTblLen);
    for I := 0 to (UnicTblLen - 1) do
    begin
      Symbols[I].Index := InFnt.Symbols[I].Index;
      Symbols[I].Width := InFnt.Symbols[I].Width;
    end;
  end
  else
  begin
    UnicodeTable := nil;
    Symbols := nil;
  end;
end;

procedure THPDFFontObj.GetFontFeatures;
var
  HDescript: HDC;
  LFont: TLogFont;
  FntObj: THandle;
  TextM: TTextMetric;
begin
  HDescript := CreateCompatibleDC(0);
  try
    FillChar(LFont, sizeof(LFont), 0);
    with LFont do
    begin
      lfHeight := -1000;
      lfWidth := 0;
      lfEscapement := 0;
      lfOrientation := 0;
      if fsBold in FontStyle then
        lfWeight := FW_BOLD
      else
        lfWeight := FW_NORMAL;
      lfItalic := byte(fsItalic in FontStyle);
      lfCharSet := FontCharset;
      StrPCopy(lfFaceName, String(OldName));
      lfQuality := DEFAULT_QUALITY;
      lfOutPrecision := OUT_DEFAULT_PRECIS;
      lfClipPrecision := CLIP_DEFAULT_PRECIS;
      lfPitchAndFamily := DEFAULT_PITCH;
    end;
    FntObj := CreateFontIndirect(LFont);
    try
      SelectObject(HDescript, FntObj);
      GetTextMetrics(HDescript, TextM);
      Ascent := TextM.tmAscent;
      IsMonospaced := (TextM.tmPitchAndFamily and TMPF_FIXED_PITCH) = 0;
      FontLen := TextM.tmHeight - TextM.tmInternalLeading;
      FillChar(OutTextM, sizeof(OutTextM), 0);
      OutTextM.otmSize := sizeof(OutTextM);
      GetCharABCWidths(HDescript, 0, 255, ABCArray);
      GetOutlineTextMetrics(HDescript, sizeof(OutTextM), @OutTextM);
    finally
      DeleteObject(FntObj);
    end;
  finally
    DeleteDC(HDescript);
  end;
end;

procedure THPDFFontObj.ParseFontName;
var
  FontName, NewName: AnsiString;
  IsStdFont: boolean;
begin
  IsStdFont := false;
  OldName := Name;
  FontName := AnsiString(LowerCase(String(Name)));
  NewName := Name;
  if (FontName = 'arial') or (FontName = 'helvetica') then
  begin
    if (fsBold in FontStyle) and (fsItalic in FontStyle) then
      NewName := 'Helvetica-BoldOblique'
    else if (fsItalic in FontStyle) then
      NewName := 'Helvetica-Oblique'
    else if (fsBold in FontStyle) then
      NewName := 'Helvetica-Bold'
    else
      NewName := 'Helvetica';
    IsStdFont := true;
  end
  else if FontName = 'times new roman' then
  begin
    if (fsBold in FontStyle) and (fsItalic in FontStyle) then
      NewName := 'Times-BoldItalic'
    else if (fsItalic in FontStyle) then
      NewName := 'Times-Italic'
    else if (fsBold in FontStyle) then
      NewName := 'Times-Bold'
    else
      NewName := 'Times-Roman';
    IsStdFont := true;
  end
  else if FontName = 'courier new' then
  begin
    if (fsBold in FontStyle) and (fsItalic in FontStyle) then
      NewName := 'Courier-BoldOblique'
    else if (fsItalic in FontStyle) then
      NewName := 'Courier-Oblique'
    else if (fsBold in FontStyle) then
      NewName := 'Courier-Bold'
    else
      NewName := 'Courier';
    IsStdFont := true;
  end
  else if FontName = 'symbol' then
  begin
    NewName := 'Symbol';
    IsStdFont := true;
  end
  else if FontName = 'zapfdingbats' then
  begin
    NewName := 'ZapfDingbats';
    IsStdFont := true;
  end;
  if not(IsStdFont) then
  begin
    if (fsBold in FontStyle) and (fsItalic in FontStyle) then
      NewName := NewName + '-ItalicBold'
    else if (fsItalic in FontStyle) then
      NewName := NewName + '-Italic'
    else if (fsBold in FontStyle) then
      NewName := NewName + '-Bold';
  end;
  Name := NewName;
  IsStandard := IsStdFont;
end;

procedure THotPDF.SetActionScript(ActionType: THPDFActionScriptType; ActionText: AnsiString);
begin
  case ActionType of
    astOpen:
      begin
        FastOpen := ActionText;
      end;
    astClose:
      begin
        FastClose := ActionText;
      end;
    astWillSave:
      begin
        FastWillSave := ActionText;
      end;
    astDidSave:
      begin
        FastDidSave := ActionText;
      end;
    astWillPrint:
      begin
        FastWillPrint := ActionText;
      end;
    astDidPrint:
      begin
        FastDidPrint := ActionText;
      end;
  end;
end;

procedure THotPDF.ListExtDictionary(PageObject: THPDFDictionaryObject; PageLink: THPDFObjectNumber);
var
  I: Integer;
  MediaBoxObj: THPDFArrayObject;
  LinkObj: THPDFLink;
  KidsArr: THPDFArrayObject;
  NextDict: THPDFDictionaryObject;
  TypeObj: THPDFNameObject;
  PageDictLen: Integer;
begin
  PageDictLen := PageObject.FindValue('Type');
  if (PageDictLen >= 0) then
  begin
    TypeObj := THPDFNameObject(PageObject.GetIndexedItem(PageDictLen));
    if (Pos('pages', LowerCase(String(TypeObj.Value))) > 0) then
    begin
      PageObject.IsDeleted := true;
      PageDictLen := PageObject.FindValue('Kids');
      KidsArr := THPDFArrayObject(PageObject.GetIndexedItem(PageDictLen));
      for I := 0 to KidsArr.Items.Count - 1 do
      begin
        LinkObj := THPDFLink(KidsArr.GetIndexedItem(I));
        NextDict := THPDFDictionaryObject(GetObjectByLink(LinkObj));
        ListExtDictionary(NextDict, LinkObj.Value);
      end;
    end
    else
    begin
      PageArr[PageArrPosition].PageObj := PageObject;
      PageArr[PageArrPosition].PageLink := PageLink;
      if ((FIsParented) and (PageObject.FindValue('MediaBox') < 0)) then
      begin
        MediaBoxObj := THPDFArrayObject.Create(nil);
        MediaBoxObj.AddNumericValue(FParentMB[0]);
        MediaBoxObj.AddNumericValue(FParentMB[1]);
        MediaBoxObj.AddNumericValue(FParentMB[2]);
        MediaBoxObj.AddNumericValue(FParentMB[3]);
        PageObject.AddValue('MediaBox', MediaBoxObj);
      end;
      Inc(PageArrPosition);
    end;
  end
  else raise Exception.Create('Cannot load the pages.');
end;

function THotPDF.GetObjectByLink(LinkObj: THPDFLink): THPDFObject;
var
  I: Integer;
  NextObj: THPDFObject;
begin
  result := nil;
  for I := 0 to (IndirectObjects.Count - 1) do
  begin
    NextObj := THPDFObject(IndirectObjects.Items[I]);
    if (CompareObjectID(LinkObj.Value, NextObj.ID) = 0) then
    begin
      result := NextObj;
      Exit;
    end;
  end;
end;

procedure THotPDF.CloseIndirectObjects;
var
  I: Integer;
  DictionaryObj: THPDFDictionaryObject;
begin
  for I := 0 to IndirectObjects.Count - 1 do
  begin
    if THPDFObject(IndirectObjects.Items[I]).ObjectType = otDictionary then
    begin
      DictionaryObj := THPDFDictionaryObject(IndirectObjects.Items[I]);
      DictionaryObj.Free;
    end
    else
      if THPDFObject(IndirectObjects.Items[I]).ObjectType = otArray
        then THPDFArrayObject(IndirectObjects.Items[I]).Free
      else
        if THPDFObject(IndirectObjects.Items[I]).ObjectType = otStream
          then THPDFStreamObject(IndirectObjects.Items[I]).Free
    else
      TObject(IndirectObjects.Items[I]).Free;
    IndirectObjects.Items[I] := nil;
  end;
  IndirectObjects.Free;
end;

procedure THotPDF.SaveToFile(FileName: TFileName);
var
  FSDoc: TFileStream;
begin
  FSDoc := TFileStream.Create(FileName, fmCreate);
  try
    SaveToStream(TStream(FSDoc));
  finally
    FSDoc.Free;
  end;
  if (FAutoLaunch) then
{$IFDEF UNICODE}
    ShellExecute(0, 'open', PWideChar(FileName), nil, nil, 0);
{$ELSE}
    ShellExecute(0, 'open', PChar(FileName), nil, nil, 0);
{$ENDIF}
end;

procedure THotPDF.StreamSaveString(DocStream: TStream; DocSting: AnsiString);
var
  FirBu: PAnsiChar;
begin
  FirBu := @DocSting[1];
  DocStream.Write(FirBu^, Length(DocSting));
end;

procedure THotPDF.SaveToStream(DocStream: TStream);
var
  ShortLink: AnsiString;
  NXrefLen: Integer;
  I, h, XrefPos: Integer;
  CurrentObj: THPDFObject;
  XREFList: array of Integer;

  function ConvertObjectNumber(ValObject: THPDFObject): AnsiString;
  begin
    result := AnsiString(IntToStr(ValObject.ID.ObjectNumber)) + ' ';
    result := result + AnsiString(IntToStr(ValObject.ID.GenerationNumber));
  end;

begin
  StreamSaveString(DocStream, AnsiString('%PDF-1.4' + #10));
  StreamSaveString(DocStream, AnsiString('%vPlC' + #10));
  NXrefLen := 0;
  LinksLen := IndirectObjects.Count;
  SetLength(LinkTable, LinksLen);
  for I := 0 to IndirectObjects.Count - 1 do
  begin
    CurrentObj := THPDFObject(IndirectObjects.Items[I]);
    if (not (CurrentObj.IsDeleted)) then
    begin
      Inc(NXrefLen);
      if (CurrentObj.ID.ObjectNumber > LinksLen) then
      begin
        LinksLen := CurrentObj.ID.ObjectNumber;
        SetLength(LinkTable, LinksLen);
      end;
      LinkTable[CurrentObj.ID.ObjectNumber - 1][CurrentObj.ID.GenerationNumber] := NXrefLen;
      CurrentObj.ID.ObjectNumber := NXrefLen;
      CurrentObj.ID.GenerationNumber := 0;
    end;
  end;
  NXrefLen := 0;
  for I := 0 to IndirectObjects.Count - 1 do
  begin
    CurrentObj := THPDFObject(IndirectObjects.Items[I]);
    if (not (CurrentObj.IsDeleted)) then
    begin
      Inc(NXrefLen);
      SetLength(XREFList, NXrefLen);
      XREFList[NXrefLen - 1] := DocStream.Position;
      SaveTypeObject(CurrentObj, DocStream, false);
    end;
  end;
  XrefPos := DocStream.Position;
  StreamSaveString(DocStream, 'xref' + #13 + #10);
  StreamSaveString(DocStream, '0 ' + AnsiString(IntToStr(NXrefLen + 1)) + #13 + #10);
  StreamSaveString(DocStream, '0000000000 65535 f' + #13 + #10);
  for I := 0 to NXrefLen - 1 do
  begin
    ShortLink := AnsiString(IntToStr(XREFList[I]));
    h := Length(ShortLink);
    while h < 10 do
    begin
      ShortLink := '0' + ShortLink;
      Inc(h)
    end;
    StreamSaveString(DocStream, AnsiString(ShortLink) + ' 00000 n' + #13 + #10);
  end;
  StreamSaveString(DocStream, 'trailer' + #13 + #10);
  StreamSaveString(DocStream, '<<' + #13 + #10 + '/Size ' + AnsiString(IntToStr(NXrefLen + 1)) + #13 + #10);
  StreamSaveString(DocStream, '/Root ' + AnsiString(ConvertObjectNumber(THPDFObject(IndirectObjects.Items[FRootIndex])) + ' R' + #13 + #10));
  StreamSaveString(DocStream, '/Info ' + AnsiString(ConvertObjectNumber(THPDFObject(IndirectObjects.Items[FInfoIndex])) + ' R' + #13 + #10));
  if ((FIsEncrypted) or (FProtection)) then
  begin
    StreamSaveString(DocStream, '/Encrypt ' + AnsiString(ConvertObjectNumber(THPDFObject(IndirectObjects.Items[FEncryprtIndex])) + ' R' + #13 + #10));
    StreamSaveString(DocStream, '/ID ' + '[<' + AnsiString(DocID) + '> <' + AnsiString(DocID) + '>]' + #13 + #10);
  end;
  StreamSaveString(DocStream, '>>' + #13 + #10);
  StreamSaveString(DocStream, 'startxref' + #13 + #10);
  StreamSaveString(DocStream, AnsiString(IntToStr(XrefPos)) + #13 + #10);
  StreamSaveString(DocStream, '%%EOF' + #13 + #10);
end;

function THotPDF.SaveTypeObject(ValObject: THPDFObject; ObjStream: TStream; IsArrayItem: boolean): Integer;
var
  ObjType: THPDFObjectType;
  NullObj: THPDFNullObject;
  BooleanObj: THPDFBooleanObject;
  NumericObj: THPDFNumericObject;
  StringObj: THPDFStringObject;
  NameObj: THPDFNameObject;
  StreamObj: THPDFStreamObject;
  ArrayObj: THPDFArrayObject;
  DictionaryObj: THPDFDictionaryObject;
  LinkObj: THPDFLink;
begin
  result := 0;
  ObjType := ValObject.ObjectType;
  case ObjType of
    otNull:
      begin
        NullObj := THPDFNullObject(ValObject);
        result := SaveNullObject(NullObj, ObjStream);
      end;
    otBoolean:
      begin
        BooleanObj := THPDFBooleanObject(ValObject);
        result := SaveBooleanObject(BooleanObj, ObjStream);
      end;
    otNumeric:
      begin
        NumericObj := THPDFNumericObject(ValObject);
        result := SaveNumericObject(NumericObj, ObjStream);
      end;
    otString:
      begin
        StringObj := THPDFStringObject(ValObject);
        result := SaveStringObject(StringObj, ObjStream);
      end;
    otName:
      begin
        NameObj := THPDFNameObject(ValObject);
        result := SaveNameObject(NameObj, ObjStream);
      end;
    otArray:
      begin
        ArrayObj := THPDFArrayObject(ValObject);
        result := SaveArrayObject(ArrayObj, ObjStream);
      end;
    otDictionary:
      begin
        DictionaryObj := THPDFDictionaryObject(ValObject);
        result := SaveDictionaryObject(DictionaryObj, ObjStream, IsArrayItem);
      end;
    otStream:
      begin
        StreamObj := THPDFStreamObject(ValObject);
        result := SaveStreamObject(StreamObj, ObjStream);
      end;
    otLink:
      begin
        LinkObj := THPDFLink(ValObject);
        result := SaveLinkObject(LinkObj, ObjStream);
      end;
  end;
end;

function THotPDF.SaveArrayObject(ValObject: THPDFArrayObject;
  ObjStream: TStream): Integer;
var
  I, SavLen, ArrayStep: Integer;
  FullLenS: Integer;
  PrevStr, PostStr: AnsiString;
  ItemObj: THPDFObject;
begin
  FullLenS := 0;
  if (ValObject.IsIndirect) then
  begin
    PrevStr := AnsiString(IntToStr(ValObject.ID.ObjectNumber)) + ' ' + AnsiString(IntToStr(ValObject.ID.GenerationNumber)) + ' obj' + #13 + #10;
    StreamSaveString(ObjStream, AnsiString(PrevStr));
    FullLenS := FullLenS + Length(PrevStr);
  end;
  FullLenS := FullLenS + 1;
  StreamSaveString(ObjStream, '[ ');
  ArrayStep := 0;
  for I := 0 to ValObject.Items.Count - 1 do
  begin
    if (ArrayStep >= 100) then
    begin
      StreamSaveString(ObjStream, #13 + #10);
      FullLenS := FullLenS + 2;
      ArrayStep := 0;
    end;
    ItemObj := THPDFObject(ValObject.Items.Items[I]);
    if (not (ItemObj.IsDeleted)) then
    begin
      SavLen := SaveTypeObject(ItemObj, ObjStream, true);
      if (I < (ValObject.Items.Count - 1)) then
        StreamSaveString(ObjStream, ' ');
      FullLenS := FullLenS + SavLen;
      ArrayStep := ArrayStep + SavLen;
    end;
  end;
  StreamSaveString(ObjStream, ' ]');
  FullLenS := FullLenS + 1;
  if (ValObject.IsIndirect) then
  begin
    PostStr := #13 + #10 + 'endobj' + #13 + #10;
    FullLenS := FullLenS + Length(PostStr);
    StreamSaveString(ObjStream, AnsiString(PostStr));
  end;
  result := FullLenS;
end;

function THotPDF.SaveBooleanObject(ValObject: THPDFBooleanObject;
  ObjStream: TStream): Integer;
begin
  if ValObject.Value then
    result := SaveObjectValue(THPDFObject(ValObject), 'true', ObjStream)
  else
    result := SaveObjectValue(THPDFObject(ValObject), 'false', ObjStream);
end;

function THotPDF.SaveDictionaryObject(ValObject: THPDFDictionaryObject; ObjStream: TStream; IsArrayItem: boolean): Integer;
var
  I: Integer;
  FullLenS: Integer;
  ItemObj: THPDFObject;
  PItem: PHPDFDictionaryItem;
  PrevStr, PostStr: AnsiString;
begin
  FullLenS := 0;
  if (ValObject.IsIndirect) then
  begin
    PrevStr := AnsiString(IntToStr(ValObject.ID.ObjectNumber)) + ' ' + AnsiString(IntToStr(ValObject.ID.GenerationNumber)) + ' obj' + #10;
    StreamSaveString(ObjStream, AnsiString(PrevStr));
    FullLenS := Length(PrevStr);
  end;
  StreamSaveString(ObjStream, '<< ');
  if (not (IsArrayItem)) then StreamSaveString(ObjStream, #10);
  FullLenS := FullLenS + 4;
  for I := 0 to ValObject.Items.Count - 1 do
  begin
    PItem := PHPDFDictionaryItem(ValObject.Items.Items[I]);
    if (not (PItem^.Value.IsDeleted)) then
    begin
      FullLenS := FullLenS + Length(PItem^.Key) + 2;
      StreamSaveString(ObjStream, '/' + AnsiString(PItem^.Key) + ' ');
      ItemObj := PItem^.Value;
      FullLenS := FullLenS + SaveTypeObject(ItemObj, ObjStream, true) + 1;
      if ((not(PItem^.Value.ObjectType = otString)) and
        (not(PItem^.Value.ObjectType = otDictionary))) then
        StreamSaveString(ObjStream, ' ');
      if (not(IsArrayItem)) then
        StreamSaveString(ObjStream, #10);
    end;
  end;
  StreamSaveString(ObjStream, '>> ');
  if (not(IsArrayItem)) then
    if (ValObject.IsIndirect) then
      StreamSaveString(ObjStream, #10);
  if (ValObject.IsIndirect) then
  begin
    PostStr := 'endobj' + #10;
    FullLenS := FullLenS + Length(PostStr);
    StreamSaveString(ObjStream, AnsiString(PostStr));
  end;
  result := FullLenS;
end;

function THotPDF.SaveLinkObject(ValObject: THPDFLink; ObjStream: TStream): Integer;
var
  ONVal: Integer;
begin
  if (ValObject.Value.ObjectNumber = 0) then
    result := SaveObjectValue(THPDFObject(ValObject), '0 0 R', ObjStream)
  else
  begin
    ONVal := LinkTable[ValObject.Value.ObjectNumber - 1][ValObject.Value.GenerationNumber];
    if (ONVal = 0) then
      ONVal := ValObject.Value.ObjectNumber;
    result := SaveObjectValue(THPDFObject(ValObject), AnsiString(IntToStr(ONVal)) + ' 0 R', ObjStream);
  end;
end;

function THotPDF.SaveNameObject(ValObject: THPDFNameObject;
  ObjStream: TStream): Integer;
var
  SVPos: Integer;
  SVName: AnsiString;
begin
  SVName := AnsiString(Trim(String(ValObject.Value)));
  SVPos := Pos(' ', String(SVName));
  while (SVPos > 0) do
  begin
    SVName := Copy(SVName, 1, SVPos - 1) + '#20' + Copy(SVName, SVPos + 1, Length(SVName) - (SVPos + 1));
    SVPos := Pos(' ', String(SVName));
  end;
  result := SaveObjectValue(THPDFObject(ValObject), '/' + SVName, ObjStream);
end;

function THotPDF.SaveNullObject(ValObject: THPDFNullObject;
  ObjStream: TStream): Integer;
begin
  result := SaveObjectValue(THPDFObject(ValObject), 'null', ObjStream);
end;

function THotPDF.SaveNumericObject(ValObject: THPDFNumericObject;
  ObjStream: TStream): Integer;
var
  FracPartVal, ValMask: Single;
  ZCount: Integer;
  IntPartVal: Integer;
  IntStr, FracStr: AnsiString;
begin
  IntPartVal := Round(Int(ValObject.Value));
  IntStr := AnsiString(IntToStr(IntPartVal));
  if (IntPartVal <> ValObject.Value) then
  begin
    ZCount := 1;
    ValMask := (ValObject.Value - IntPartVal) * 10;
    FracPartVal := Round(ValMask);
    while (ValMask <> FracPartVal) do
    begin
      Inc(ZCount);
      ValMask := ValMask * 10;
      FracPartVal := Round(ValMask);
    end;
    FracStr := AnsiString(IntToStr(ABS(Round(ValMask))));
    if (Length(FracStr) <> ZCount) then
    begin
      ZCount := ZCount - Length(FracStr);
      while (ZCount > 0) do
      begin
        FracStr := '0' + FracStr;
        Dec(ZCount);
      end;
    end;
    Result := SaveObjectValue(THPDFObject(ValObject), IntStr + '.' + FracStr, ObjStream);
  end
  else
    result := SaveObjectValue(THPDFObject(ValObject), IntStr, ObjStream);
end;

function THotPDF.SaveStreamObject(ValObject: THPDFStreamObject; ObjStream: TStream): Integer;
var
  CurrLenObject: THPDFNumericObject;
  StrLenIndex: Integer;
  FullLenS: Integer;
  PrevStr, PostStr: AnsiString;
  TmpStream: TMemoryStream;
begin
  FullLenS := 0;
  StrLenIndex := ValObject.Dictionary.FindValue('Length');
  if (StrLenIndex > -1) then
  begin
    CurrLenObject := THPDFNumericObject(ValObject.Dictionary.GetIndexedItem(StrLenIndex));
  end
  else
  begin
    CurrLenObject := THPDFNumericObject.Create(nil);
    ValObject.Dictionary.AddValue('Length', CurrLenObject);
  end;
  CurrLenObject.Value := ValObject.Stream.Size;
  if (ValObject.IsIndirect) then
  begin
    PrevStr := AnsiString(IntToStr(ValObject.ID.ObjectNumber)) + ' ' + AnsiString(IntToStr(ValObject.ID.GenerationNumber)) + ' obj' + #10;
    StreamSaveString(ObjStream, AnsiString(PrevStr));
    FullLenS := Length(PrevStr);
  end;
  FullLenS := FullLenS + SaveDictionaryObject(ValObject.Dictionary, ObjStream, true) + 18;
  StreamSaveString(ObjStream, #10 + 'stream' + #10);
  ValObject.Stream.Position := 0;
  if (ValObject.Stream.Size > 0) then
  begin
    if (FProtection) then
    begin
      TmpStream := TMemoryStream.Create;
      try
        TmpStream.CopyFrom(ValObject.Stream, 0);
        TmpStream.Position := 0;
        CryptStream(TMemoryStream(TmpStream), ValObject.ID.ObjectNumber);
        TmpStream.Position := 0;
        ObjStream.CopyFrom(TmpStream, 0);
        FullLenS := FullLenS + TmpStream.Size;
      finally
        TmpStream.Free;
      end;
    end
    else
    begin
      ObjStream.CopyFrom(ValObject.Stream, ValObject.Stream.Size);
      FullLenS := FullLenS + ValObject.Stream.Size;
    end;
  end;
  StreamSaveString(ObjStream, #10 + 'endstream' + #10);
  if (ValObject.IsIndirect) then
  begin
    PostStr := 'endobj' + #10;
    FullLenS := FullLenS + Length(PostStr);
    StreamSaveString(ObjStream, AnsiString(PostStr));
  end;
  result := FullLenS;
end;

function THotPDF.SaveStringObject(ValObject: THPDFStringObject;
  ObjStream: TStream): Integer;
var
  TypeStObj: THPDFObject;
  TypeStVal: AnsiString;

  function IsEncryptObj(EncrObj: THPDFObject): boolean;
  var
    EncrDict: THPDFDictionaryObject;
  begin
    result := false;
    if (EncrObj <> nil) then
    begin
      if (EncrObj.ObjectType = otDictionary) then
      begin
        EncrDict := THPDFDictionaryObject(EncrObj);
        result := ((EncrDict.FindValue('V') >= 0) and (EncrDict.FindValue('O') >= 0) and (EncrDict.FindValue('U') >= 0));
      end;
    end;
  end;

var
  IEOt: boolean;

begin
  if (ValObject.IsHexadecimal) then
    result := SaveObjectValue(THPDFObject(ValObject), '<' + ValObject.Value + '>', ObjStream)
  else
  begin
    TypeStVal := ValObject.Value;
    IEOt := IsEncryptObj(ValObject.FParent);
    if ((FProtection) and (Length(TypeStVal) > 0)) then
    begin
      TypeStObj := ValObject;
      if (not (IEOt)) then
      begin
        while ((not (TypeStObj.IsIndirect)) and (not (TypeStObj.FParent = nil))) do
        begin
          TypeStObj := ValObject.FParent;
        end;
        if (TypeStObj.IsIndirect) then TypeStVal := CryptString(ValObject.Value, TypeStObj.ID.ObjectNumber);
      end;
    end;
    Result := SaveObjectValue(THPDFObject(ValObject), '(' + _EscapeText(TypeStVal) + ')', ObjStream);
  end;
end;

function THotPDF.SaveObjectValue(ValObject: THPDFObject; Value: AnsiString;
  ObjStream: TStream): Integer;
var
  PrevStr, PostStr: AnsiString;
begin
  result := Length(Value);
  if (ValObject.IsIndirect) then
  begin
    PrevStr := AnsiString(IntToStr(ValObject.ID.ObjectNumber)) + ' ' +
      AnsiString(IntToStr(ValObject.ID.GenerationNumber)) + ' obj' + #13 + #10;
    PostStr := 'endobj' + #13 + #10;
    StreamSaveString(ObjStream, AnsiString(PrevStr));
    StreamSaveString(ObjStream, AnsiString(Value + #13 + #10));
    StreamSaveString(ObjStream, AnsiString(PostStr));
    result := result + Length(AnsiString(PrevStr)) + Length(PostStr);
  end
  else
  begin
    StreamSaveString(ObjStream, AnsiString(Value));
  end;
end;

procedure THotPDF.SetAutoLaunch(const Value: boolean);
begin
  if FProgress then
    raise Exception.Create('Cannot set AutoLaunch value - document in progress.');
  FAutoLaunch := Value;
end;

procedure THotPDF.SetCompressionMethod(const Value: THPDFCompressionMethod);
begin
  if FProgress then
    raise Exception.Create('Cannot set Compression value - document in progress.');
  FCompressionMethod := Value;
end;

procedure THotPDF.SetStandardFontEmulation(const Value: boolean);
begin
  if FProgress then
    raise Exception.Create('Cannot set StandardFontEmulation value - document in progress.');
  FStandardFontEmulation := Value;
end;

procedure THotPDF.SetCryptKeyType(const Value: THPDFKeyType);
begin
  FCryptKeyType := Value;
  if (FCryptKeyType = k40) then
  begin
    FRevision := 2;
  end
  else
  begin
    FRevision := 3;
  end;
end;

procedure THotPDF.SetNEmbeddedFont(const Value: TStringList);
begin
  FNEmbeddedFonts := Value;
end;

procedure THotPDF.SetPageLayout(const Value: THPDFPageLayout);
begin
  FPageLayout := Value;
end;

procedure THotPDF.EndDoc;
var
  I: Integer;
  PLInd: Integer;
  KidLen: Integer;
  BlockInd: Integer;
  AAArrayObj: THPDFDictionaryObject;
  PItem: PHPDFDictionaryItem;
  AAOpen: THPDFDictionaryObject;
  AAClose: THPDFDictionaryObject;
  AAWillSave: THPDFDictionaryObject;
  AADidSave: THPDFDictionaryObject;
  AAWillPrint: THPDFDictionaryObject;
  AADidPrint: THPDFDictionaryObject;
  VPObj: THPDFDictionaryObject;
  InfoObj: THPDFDictionaryObject;
  KidsObj: THPDFArrayObject;
  PagesObj: THPDFDictionaryObject;
  RootDictObj: THPDFDictionaryObject;
  OutlineDictObj: THPDFDictionaryObject;
begin
  RootDictObj := THPDFDictionaryObject(IndirectObjects.Items[FRootIndex]);
  PLInd := RootDictObj.FindValue('PageLayout');
  if (PLInd >= 0) then
  begin
    THPDFObject(RootDictObj.GetIndexedItem(PLInd)).Free;
    PItem := RootDictObj.Items.Items[PLInd];
    Dispose(PItem);
    RootDictObj.Items.Delete(PLInd);
  end;
  case FPageLayout of
    plSinglePage: RootDictObj.AddNameValue('PageLayout', 'SinglePage');
    plOneColumn: RootDictObj.AddNameValue('PageLayout', 'OneColumn');
    plTwoColumnLeft: RootDictObj.AddNameValue('PageLayout', 'TwoColumnLeft');
    plTwoColumnRight: RootDictObj.AddNameValue('PageLayout', 'TwoColumnRight');
  end;

  PLInd := RootDictObj.FindValue('PageMode');
  if (PLInd >= 0) then
  begin
    THPDFObject(RootDictObj.GetIndexedItem(PLInd)).Free;
    PItem := RootDictObj.Items.Items[PLInd];
    Dispose(PItem);
    RootDictObj.Items.Delete(PLInd);
  end;
  case FPageMode of
    pmUseNone: RootDictObj.AddNameValue('PageMode', 'UseNone');
    pmUseOutlines: RootDictObj.AddNameValue('PageMode', 'UseOutlines');
    pmUseThumbs: RootDictObj.AddNameValue('PageMode', 'UseThumbs');
    pmFullScreen: RootDictObj.AddNameValue('PageMode', 'FullScreen');
    pmUseAttachments: RootDictObj.AddNameValue('PageMode', 'UseAttachments');
  end;

  if (FVPChanged) then
  begin
    PLInd := RootDictObj.FindValue('ViewerPreferences');
    if (PLInd >= 0) then
    begin
      THPDFObject(RootDictObj.GetIndexedItem(PLInd)).Free;
      PItem := RootDictObj.Items.Items[PLInd];
      Dispose(PItem);
      RootDictObj.Items.Delete(PLInd);
    end;
    VPObj := THPDFDictionaryObject.Create(nil);
    if (vpHideToolbar in FViewerPreference) then
      VPObj.AddBooleanValue('HideToolbar', true);
    if (vpHideMenubar in FViewerPreference) then
      VPObj.AddBooleanValue('HideMenubar', true);
    if (vpHideWindowUI in FViewerPreference) then
      VPObj.AddBooleanValue('HideWindowUI', true);
    if (vpFitWindow in FViewerPreference) then
      VPObj.AddBooleanValue('FitWindow', true);
    if (vpCenterWindow in FViewerPreference) then
      VPObj.AddBooleanValue('CenterWindow', true);
    RootDictObj.AddValue('ViewerPreferences', THPDFObject(VPObj));
  end;
  BlockInd := RootDictObj.FindValue('AA');
  if (BlockInd >= 0) then
  begin
    AAArrayObj := THPDFDictionaryObject(RootDictObj.GetIndexedItem(BlockInd));
  end
  else
  begin
    AAArrayObj := THPDFDictionaryObject.Create(nil);
    RootDictObj.AddValue('AA', AAArrayObj);
  end;
  BlockInd := RootDictObj.FindValue('OpenAction');
  if (BlockInd >= 0) then
  begin
    if (FastOpen <> '') then
    begin
      AAOpen := CreateIndirectDictionary;
      AAOpen.AddNameValue('S', 'JavaScript');
      AAOpen.AddStringValue('JS', FastOpen);
      RootDictObj.ReplaceValue('OpenAction', AAOpen);
    end;
  end
  else
  begin
    if (FastOpen <> '') then
    begin
      AAOpen := CreateIndirectDictionary;
      AAOpen.AddNameValue('S', 'JavaScript');
      AAOpen.AddStringValue('JS', FastOpen);
      RootDictObj.AddValue('OpenAction', AAOpen);
    end;
  end;
  if (FastClose <> '') then
  begin
    AAClose := CreateIndirectDictionary;
    AAClose.AddNameValue('S', 'JavaScript');
    AAClose.AddStringValue('JS', FastClose);
    AAArrayObj.AddValue('WC', AAClose);
  end;
  if (FastWillSave <> '') then
  begin
    AAWillSave := CreateIndirectDictionary;
    AAWillSave.AddNameValue('S', 'JavaScript');
    AAWillSave.AddStringValue('JS', FastWillSave);
    AAArrayObj.AddValue('WS', AAWillSave);
  end;
  if (FastDidSave <> '') then
  begin
    AADidSave := CreateIndirectDictionary;
    AADidSave.AddNameValue('S', 'JavaScript');
    AADidSave.AddStringValue('JS', FastDidSave);
    AAArrayObj.AddValue('DS', AADidSave);
  end;
  if (FastWillPrint <> '') then
  begin
    AAWillPrint := CreateIndirectDictionary;
    AAWillPrint.AddNameValue('S', 'JavaScript');
    AAWillPrint.AddStringValue('JS', FastWillPrint);
    AAArrayObj.AddValue('WP', AAWillPrint);
  end;
  if (FastDidPrint <> '') then
  begin
    AADidPrint := CreateIndirectDictionary;
    AADidPrint.AddNameValue('S', 'JavaScript');
    AADidPrint.AddStringValue('JS', FastDidPrint);
    AAArrayObj.AddValue('DP', AADidPrint);
  end;
  if (FOutlineRoot <> nil) then
  begin
    if (FOutlineRoot.FCount > 0) then
    begin
      OutlineDictObj := CreateIndirectDictionary;
      OutlineDictObj.AddNameValue('Type', 'Outlines');
      OutlineDictObj.AddNumericValue('Count', FOutlineRoot.FCount);
      OutlineDictObj.AddValue('First', FOutlineRoot.FFirst.LinkedObj);
      OutlineDictObj.AddValue('Last', FOutlineRoot.FLast.LinkedObj);
      RootDictObj.AddValue('Outlines', OutlineDictObj);
      RootDictObj.AddNameValue('PageMode', 'UseOutlines');
      for I := 0 to OutlineEnsLen - 1 do
      begin
        OutlineEnsemble[I].Free;
      end;
    end;
    FOutlineRoot.LinkedObj.Free;
    FOutlineRoot.Free;
  end;
  if (FCurrentPage <> nil) then
  begin
    FCurrentPage.ClosePage;
    FCurrentPage.Free;
    FCurrentPage := nil;
  end;
  PagesObj := THPDFDictionaryObject(IndirectObjects.Items[FPagesIndex]);
  BlockInd := PagesObj.FindValue('Kids');
  if (BlockInd >= 0) then
  begin
    KidsObj := THPDFArrayObject(PagesObj.GetIndexedItem(BlockInd));
    for I := 0 to KidsObj.Items.Count - 1 do
    begin
      THPDFObject(KidsObj.Items.Items[I]).Free;
    end;
    KidsObj.Items.Clear;
    KidLen := Length(PageArr);
    for I := 0 to KidLen - 1 do
    begin
      if (not (PageArr[I].PageObj.IsDeleted)) then
        KidsObj.AddObject(PageArr[I].PageObj);
    end;
  end
  else
    raise exception.Create('Invalid pages object.');
  BlockInd := PagesObj.FindValue('Count');
  if (BlockInd >= 0) then
  begin
    THPDFNumericObject(PagesObj.GetIndexedItem(BlockInd)).Value := KidLen;
  end
  else
    raise exception.Create('Invalid pages object.');
  InfoObj := THPDFDictionaryObject(IndirectObjects.Items[FInfoIndex]);
  InfoObj.AddStringValue('CreationDate', _DateTimeToPdfDate(FCreationDate));
  InfoObj.AddStringValue('Creator', 'losLab HotPDF');
  InfoObj.AddStringValue('Author', FAuthor);
  InfoObj.AddStringValue('Title', FTitle);
  InfoObj.AddStringValue('Keywords', FKeywords);
  InfoObj.AddStringValue('Subject', FSubject);
  InfoObj.AddStringValue('ModDate', _DateTimeToPdfDate(FCreationDate));
  if FMemStream then
  begin
    SaveToStream(FOutputStream);
    FOutputStream.Position := 0;
  end
  else
  begin
    if (FFileName = '') then
      raise exception.Create('Invalid filename.');
    SaveToFile(FFileName);
  end;
end;

function THotPDF.CompareObjectID(IDL, IDR: THPDFObjectNumber): Integer;
begin
  if (IDL.ObjectNumber = IDR.ObjectNumber) then
    result := IDL.GenerationNumber - IDR.GenerationNumber
  else
    result := IDL.ObjectNumber - IDR.ObjectNumber;
end;

function THotPDF.GetCanvas: TCanvas;
begin
  if (FCurrentPage = nil) then
    result := nil
  else
    result := FCurrentPage.Canvas;
end;

procedure THotPDF.LoadUnFlateLZW(RegionStream, StrumStream: TStream; NameObj: AnsiString);
var
  LNameObj: AnsiString;
  OutbSize: Integer;
  InBuff: PAnsiChar;
  OutBuf: Pointer;
begin
  LNameObj := AnsiString(LowerCase(String(NameObj)));
  if (Pos('flatedecode', String(LNameObj)) > 0) then
  begin
    GetMem(InBuff, RegionStream.Size);
    try
      RegionStream.Position := 0;
      RegionStream.Read(InBuff^, RegionStream.Size);
      OutBuf := nil;
      OutbSize := 0;
      ZDecompress(InBuff, RegionStream.Size, OutBuf, OutbSize, 0);
      try
        StrumStream.Write(OutBuf^, OutbSize);
      finally
        FreeMem(OutBuf);
      end;
    finally
      FreeMem(InBuff);
    end;
  end
  else
  begin
    if (Pos('lzwdecode', String(LNameObj)) > 0) then
      ZDecompressStream(RegionStream, StrumStream)
    else
    begin
      if (Pos('asciihexdecode', String(LNameObj)) > 0) then
        FilterStream(RegionStream, StrumStream)
      else
      begin
        if (Pos('ascii85decode', String(LNameObj)) > 0) then
          DetachValue(RegionStream, StrumStream)
        else
        begin
          if (Pos('runlengthdecode', String(LNameObj)) > 0) then
            IncFlushByte(RegionStream, StrumStream);
        end;
      end;
    end;
  end;
end;

function THotPDF.GetIsHaveSimpleText: boolean;
var
  HI: Integer;
  SmStrLin: TStringStream;
  StrFlush: TMemoryStream;
  CCPageObj: THPDFDictionaryObject;
  CCContentObj: THPDFObject;
  CCPObjStream: THPDFStreamObject;
  FilterObj: THPDFObject;

  function GetContentcount: Integer;
  var
    ContentsIndex: Integer;
  begin
    result := 0;
    ContentsIndex := CCPageObj.FindValue('Contents');
    if (ContentsIndex >= 0) then
    begin
      CCContentObj := CCPageObj.GetIndexedItem(ContentsIndex);
      if (CCContentObj.ObjectType = otArray) then
      begin
        result := THPDFArrayObject(CCContentObj).Items.Count;
      end
      else
      begin
        result := 1;
      end;
    end;
  end;

  procedure GetPStreamFromArray(ArrObject: THPDFObject);
  var
    LinkObj: THPDFLink;
  begin
    LinkObj := THPDFLink(THPDFArrayObject(ArrObject).GetIndexedItem(HI));
    CCPObjStream := THPDFStreamObject(GetObjectByLink(LinkObj));
  end;

  procedure GetObjContent;
  var
    CCInternPOStream: THPDFObject;
  begin
    if (CCContentObj.ObjectType = otArray) then
    begin
      GetPStreamFromArray(CCContentObj);
    end
    else
    begin
      if (CCContentObj.ObjectType = otLink) then
      begin
        CCInternPOStream := GetObjectByLink(THPDFLink(CCContentObj));
        if (CCInternPOStream.ObjectType = otArray) then
        begin
          GetPStreamFromArray(CCInternPOStream);
        end
        else
          CCPObjStream := THPDFStreamObject(CCInternPOStream);
      end
      else if (CCContentObj.ObjectType = otStream) then
        CCPObjStream := THPDFStreamObject(CCContentObj);
    end;
  end;

  function TrimDocContentStr(ContentStr: AnsiString): boolean;
  var
    IK: Integer;
  begin
    result := false;
    ContentStr := AnsiString(Trim(String(ContentStr)));
    if (ContentStr[1] = '<') then
      result := true
    else
    begin
      IK := 2;
      while (ContentStr[IK] <> ')') do
      begin
        if (ContentStr[IK] <> ' ') then
        begin
          result := true;
          break;
        end;
        Inc(IK);
      end;
    end;

  end;

  procedure LoadUnCompressed;
  begin
    CCPObjStream.Stream.Position := 0;
    StrFlush.CopyFrom(CCPObjStream.Stream, 0);
  end;

var
  I, FI, UF: Integer;
  ContCount: Integer;
  FilterIndex: Integer;
  TXStrLs: TStringList;
  SaperT: AnsiString;
begin
  I := 0;
  result := false;
  while (I < FPagesCount) do
  begin
    CCPageObj := PageArr[I].PageObj;
    ContCount := GetContentcount;
    for HI := 0 to ContCount - 1 do
    begin
      GetObjContent;
      StrFlush := TMemoryStream.Create;
      try
        FilterIndex := CCPObjStream.Dictionary.FindValue('Filter');
        if (FilterIndex < 0) then
        begin
          LoadUnCompressed;
        end
        else
        begin
          FilterObj := CCPObjStream.Dictionary.GetIndexedItem(FilterIndex);
          if (FilterObj.ObjectType = otLink) then
          begin
            FilterObj := GetObjectByLink(THPDFLink(FilterObj));
          end;
          if (FilterObj.ObjectType = otArray) then
          begin
            if (THPDFArrayObject(FilterObj).Items.Count = 0) then
            begin
              LoadUnCompressed;
            end
            else
            begin
              CCPObjStream.Stream.Position := 0;
              for FI := 0 to THPDFArrayObject(FilterObj).Items.Count - 1 do
              begin
                LoadUnFlateLZW(CCPObjStream.Stream, StrFlush, THPDFNameObject(THPDFArrayObject(FilterObj).GetIndexedItem(FI)).Value);
              end;
            end;
          end
          else
          begin
            if (FilterObj.ObjectType = otName) then
            begin
              CCPObjStream.Stream.Position := 0;
              LoadUnFlateLZW(CCPObjStream.Stream, StrFlush, THPDFNameObject(FilterObj).Value);
            end;
          end;
        end;
        TXStrLs := TStringList.Create;
        try
          SmStrLin := TStringStream.Create('');
          SmStrLin.Position := 0;
          SmStrLin.CopyFrom(StrFlush, 0);
          SmStrLin.Position := 0;
          TXStrLs.LoadFromStream(SmStrLin);
          SmStrLin.Free;
          for UF := 0 to TXStrLs.Count - 1 do
          begin
            SaperT := AnsiString(TXStrLs.Strings[UF]);
            if Pos('Tj', String(SaperT)) > 0 then
            begin
              if (TrimDocContentStr(SaperT)) then
              begin
                result := true;
                break;
              end;
            end;
          end;
          if (result) then
            break;
        finally
          TXStrLs.Free;
        end;
      finally
        StrFlush.Free;
      end;
    end;
    if result then
      break;
    Inc(I);
  end;
end;

function THotPDF.CreateIndirectDictionary: THPDFDictionaryObject;
begin
  Inc(FMaxObjNum);
  result := THPDFDictionaryObject.Create(nil);
  result.IsIndirect := true;
  result.ID.ObjectNumber := FMaxObjNum;
  IndirectObjects.Add(result);
end;

function THotPDF.CreateParagraph(Indention: Single; Justification: THPDFJustificationType; LeftMargin, RightMargin, TopMargin, BottomMargin: Single): Integer;
begin
  Inc(FParaLen);
  SetLength(FParas, FParaLen);
  FParas[FParaLen - 1].Justification := Justification;
  FParas[FParaLen - 1].Indention := Indention;
  FParas[FParaLen - 1].LeftMargin := LeftMargin;
  FParas[FParaLen - 1].RightMargin := RightMargin;
  FParas[FParaLen - 1].TopMargin := TopMargin;
  FParas[FParaLen - 1].BottomMargin := BottomMargin;
  result := FParaLen;
end;

procedure THotPDF.BeginDoc(Initial: boolean);
var
  CatalObj: THPDFDictionaryObject;
  InfoObj: THPDFDictionaryObject;
  PagesObj: THPDFDictionaryObject;
  KidsBoxObj: THPDFArrayObject;
begin
  FDocStarted := true;
  DocId := AnsiString(FFileName) + AnsiString(FormatDateTime('ddd dd-mm-yyyy hh:nn:ss.zzz', Now));
  DocID := MD5CalcString(DocID);
  if (FIsLoaded) then
  begin
    SetCurrentPageNum(0);
    if (FProtection) then
    begin
      EnableEncrypt;
    end;
    InfoObj := THPDFDictionaryObject(IndirectObjects.Items[FInfoIndex]);
    InfoObj.AddStringValue('CreationDate', _DateTimeToPdfDate(FCreationDate));
    InfoObj.AddStringValue('Creator', 'losLab HotPDF');
    InfoObj.AddStringValue('Author', FAuthor);
    InfoObj.AddStringValue('Title', FTitle);
    InfoObj.AddStringValue('Keywords', FKeywords);
    InfoObj.AddStringValue('Subject', FSubject);
    InfoObj.AddStringValue('ModDate', _DateTimeToPdfDate(FCreationDate));
  end
  else
  begin
    FXrefLen := 0;
    FPagesCount := 0;
    FOutlineRoot := THPDFDocOutlineObject.Create;
    FOutlineRoot.Init(Self);
    OutlineEnsemble := nil;
    OutlineEnsLen := 0;
    CatalObj := CreateIndirectDictionary;
    FRootIndex := FMaxObjNum - 1;
    FRootLink.ObjectNumber := CatalObj.ID.ObjectNumber;
    FRootLink.GenerationNumber := CatalObj.ID.GenerationNumber;
    CatalObj.AddNameValue('Type', 'Catalog');
    if (FProtection) then
    begin
      EnableEncrypt;
    end;
    InfoObj := CreateIndirectDictionary;
    FInfoIndex := FMaxObjNum - 1;
    FInfoLink.ObjectNumber := InfoObj.ID.ObjectNumber;
    FInfoLink.GenerationNumber := InfoObj.ID.GenerationNumber;
    InfoObj.AddStringValue('CreationDate', _DateTimeToPdfDate(FCreationDate));
    InfoObj.AddStringValue('Creator', 'HotPDF VCL library');
    InfoObj.AddStringValue('Author', FAuthor);
    InfoObj.AddStringValue('Title', FTitle);
    InfoObj.AddStringValue('Keywords', FKeywords);
    InfoObj.AddStringValue('Subject', FSubject);
    InfoObj.AddStringValue('ModDate', _DateTimeToPdfDate(FCreationDate));
    PagesObj := CreateIndirectDictionary;
    FPagesIndex := IndirectObjects.Count - 1;
    CatalObj.AddValue('Pages', PagesObj);
    FPageLayout := plOneColumn;
    PagesObj.AddNameValue('Type', 'Pages');
    KidsBoxObj := THPDFArrayObject.Create(nil);
    PagesObj.AddValue('Kids', KidsBoxObj);
    PagesObj.AddNumericValue('Count', 0);
    AddPage;
  end;
  FCurrentFontIndex := 0;
  FCurrentImageIndex := 0;
end;

procedure THotPDF.SetCurrentPageNum(const Value: Integer);
begin
  if (not (FIsEncrypted)) then
  begin
    if (Value <> FCurrentPageNum) then
    begin
      if (Value >= FPagesCount) then
        raise exception.Create('Invalid page number.');
      if (FCurrentPage <> nil) then
      begin
        FCurrentPage.ClosePage;
        FCurrentPage.Free;
      end;
      FCurrentPage := THPDFPage.Create;
      FCurrentPage.FParent := Self;
      FCurrentPage.PageObj := PageArr[Value].PageObj;
      FCurrentPage.ConvertPageObject;
      FCurrentPageNum := Value;
      CurrentPage.SetRGBColor(0);
      CurrentPage.SetFont('Arial', [], 16);
      CurrentPage.SetTextRenderingMode(trFill);
    end;
  end;
end;

procedure THotPDF.DeletePage(PageIndex: Integer);
var
  I: Integer;
  OldPn: Integer;
  ItLen: Integer;
  PALen: Integer;
  BlockInd: Integer;
  AnnotsArr: THPDFObject;
  XObjectDict: THPDFDictionaryObject;
  ResourDict: THPDFDictionaryObject;
begin
  if (PageIndex >= FPagesCount) then
    raise exception.Create('Invalid page number')
  else
  begin
    if (FCurrentPage <> nil) then
    begin
      FCurrentPage.ClosePage;
      FCurrentPage.Free;
      FCurrentPage := nil;
    end;
    FPagesCount := FPagesCount - 1;
    DeleteObj(PageArr[PageIndex].PageObj, false);
    BlockInd := PageArr[PageIndex].PageObj.FindValue('Resources');
    if (BlockInd >= 0) then
    begin
      ResourDict := THPDFDictionaryObject(PageArr[PageIndex].PageObj.GetIndexedItem(BlockInd));
      BlockInd := ResourDict.FindValue('XObject');
      if (BlockInd >= 0) then
      begin
        XObjectDict := THPDFDictionaryObject(ResourDict.GetIndexedItem(BlockInd));
        ItLen := XObjectDict.Items.Count;
        for I := 0 to ItLen - 1 do
        begin
          DeleteObj(XObjectDict.GetIndexedItem(I), false);
        end;
      end;
    end;
    BlockInd := PageArr[PageIndex].PageObj.FindValue('Contents');
    if (BlockInd >= 0) then
    begin
      DeleteObj(PageArr[PageIndex].PageObj.GetIndexedItem(BlockInd), false);
    end;
    BlockInd := PageArr[PageIndex].PageObj.FindValue('Annots');
    if (BlockInd >= 0) then
    begin
      AnnotsArr := THPDFObject(PageArr[PageIndex].PageObj.GetIndexedItem(BlockInd));
      if (AnnotsArr.ObjectType = otLink) then
      begin
        AnnotsArr := GetObjectByLink(THPDFLink(AnnotsArr));
      end;
      for I := 0 to THPDFArrayObject(AnnotsArr).Items.Count - 1 do
      begin
        DeleteObj(THPDFArrayObject(AnnotsArr).GetIndexedItem(I), true);
      end;
    end;
    PALen := Length(PageArr);
    for I := PageIndex to PALen - 2 do
    begin
      PageArr[I] := PageArr[I + 1];
    end;
    SetLength(PageArr, PALen - 1);
    PageArrPosition := FPagesCount - 1;
    if (FCurrentPageNum = PageIndex) then
    begin
      Dec(FCurrentPageNum);
      if FCurrentPageNum < 0 then
        FCurrentPageNum := 0;
    end;
    OldPn := FCurrentPageNum;
    FCurrentPageNum := -1;
    SetCurrentPageNumber(OldPn);
  end;
end;

procedure THotPDF.DeleteObj(InObj: THPDFObject; Recursive: boolean);
var
  I: Integer;
  ItLen: Integer;
  SimpO: THPDFObject;
  DicO: THPDFDictionaryObject;
begin
  case InObj.ObjectType of
    otBoolean: InObj.IsDeleted := true;
    otNumeric: InObj.IsDeleted := true;
    otString: InObj.IsDeleted := true;
    otName: InObj.IsDeleted := true;
    otNull: InObj.IsDeleted := true;
    otArray:
      begin
        InObj.IsDeleted := true;
      end;
    otDictionary:
      begin
        InObj.IsDeleted := true;
        if (Recursive) then
        begin
          DicO := THPDFDictionaryObject(InObj);
          ItLen := DicO.Items.Count;
          for I := 0 to ItLen - 1 do
          begin
            SimpO := DicO.GetIndexedItem(I);
            if (((SimpO.IsIndirect) or (SimpO.ObjectType = otLink)) and (DicO.GetIndexedKey(I) <> 'Parent')) then
              DeleteObj(SimpO, true);
          end;
        end;
      end;
    otStream:
      begin
        InObj.IsDeleted := true;
      end;
    otLink:
      begin
        SimpO := GetObjectByLink(THPDFLink(InObj));
        if (Recursive) then
          DeleteObj(SimpO, true)
        else
          SimpO.IsDeleted := true;
      end;
  end;
end;

procedure THotPDF.SetCurrentPageNumber(const Value: Integer);
begin
  SetCurrentPageNum(Value);
end;

{ THPDFPage }

procedure THPDFPage.CloseCanvas;
var
  ScreenRes: Extended;
  NewMeta: TMetafile;
  XCaps, YCaps: Integer;
  NewCanvas: TMetafileCanvas;
  WMeta: THPDFWmf;
begin
  FCanvas.Free;
  PageMeta.Enhanced := true;
  GStateSave;
  XCaps := GetDeviceCaps(FParent.FCHandle, HORZRES);
  YCaps := GetDeviceCaps(FParent.FCHandle, VERTRES);
  WMeta := THPDFWmf.Create(Self);
  try
    NewMeta := TMetafile.Create;
    try
      NewCanvas := TMetafileCanvas.Create(NewMeta, 0);
      try
        if (PageMeta.Width > XCaps) or (PageMeta.Height > YCaps) then
        begin
          ScreenRes := Min(XCaps / PageMeta.Width, YCaps / PageMeta.Height);
          NewMeta.Width := Round(PageMeta.Width * ScreenRes);
          NewMeta.Height := Round(PageMeta.Height * ScreenRes);
        end
        else
        begin
          NewMeta.Width := PageMeta.Width;
          NewMeta.Height := PageMeta.Height;
        end;
        NewCanvas.StretchDraw(Rect(0, 0, NewMeta.Width, NewMeta.Height), PageMeta);
      finally
        NewCanvas.Free;
      end;
      WMeta.Analyse(NewMeta);
    finally
      NewMeta.Free;
    end;
  finally
    WMeta.Free;
  end;
  GStateRestore;
  SetRGBColor($FF);
  SetFont('Arial', [fsBold], Width / 12);
  SetTextRenderingMode(trStroke);
  SetLineWidth(2);
  SetRGBColor(0);
  SetFont('Arial', [], 10);
  PageMeta.Free;
  FCanvas := nil;
end;

procedure THPDFPage.ClosePage;
var
  TmpStream: TStream;
  FilterIndex: Integer;
  FilterObj: THPDFObject;
  FiltObj: THPDFArrayObject;
  FilterNameObj: THPDFNameObject;
  TempArcad, StrFlush: TMemoryStream;

  procedure LoadUnCompressed;
  begin
    PageObjStream.Stream.Position := 0;
    FDescStream.CopyFrom(PageObjStream.Stream, 0);
  end;

var
  FI: Integer;

begin
  CloseCanvas;
  if (CurrentFontObj.IsUsed) then
    StoreCurrentFont;
  if (PageContent.Count > 0) then
  begin
    FDescStream := TMemoryStream.Create;
    try
      TempArcad := TMemoryStream.Create;
      try
        StrFlush := TMemoryStream.Create;
        try
          FilterIndex := PageObjStream.Dictionary.FindValue('Filter');
          if (FilterIndex < 0) then
          begin
            LoadUnCompressed;
          end
          else
          begin
            FilterObj := PageObjStream.Dictionary.GetIndexedItem(FilterIndex);
            if (FilterObj.ObjectType = otLink) then
            begin
              FilterObj := FParent.GetObjectByLink(THPDFLink(FilterObj));
            end;
            if (FilterObj.ObjectType = otArray) then
            begin
              if (THPDFArrayObject(FilterObj).Items.Count = 0) then
              begin
                LoadUnCompressed;
              end
              else
              begin
                TempArcad.CopyFrom(PageObjStream.Stream, 0);
                TempArcad.Position := 0;
                for FI := 0 to THPDFArrayObject(FilterObj).Items.Count - 1 do
                begin
                  FParent.LoadUnFlateLZW(TempArcad, StrFlush, THPDFNameObject(THPDFArrayObject(FilterObj).GetIndexedItem(FI)).Value);
                  TempArcad.Clear;
                  TempArcad.CopyFrom(StrFlush, 0);
                  TempArcad.Position := 0;
                end;
                FDescStream.CopyFrom(TempArcad, 0);
              end;
            end
            else
            begin
              if (FilterObj.ObjectType = otName) then
              begin
                TempArcad.CopyFrom(PageObjStream.Stream, 0);
                TempArcad.Position := 0;
                FParent.LoadUnFlateLZW(TempArcad, StrFlush, THPDFNameObject(FilterObj).Value);
                StrFlush.Position := 0;
                FDescStream.CopyFrom(StrFlush, 0);
              end;
            end;
          end;
          if (FParent.Compression = cmNone) then
          begin
            PageObjStream.Dictionary.DeleteValue('Filter');
          end
          else
          begin
            FiltObj := THPDFArrayObject.Create(nil);
            FilterNameObj := THPDFNameObject.Create(nil);
            FiltObj.AddObject(FilterNameObj);
            FilterNameObj.Value := 'FlateDecode';
            PageObjStream.Dictionary.ReplaceValue('Filter', FiltObj);
          end;
        finally
          StrFlush.Free;
        end;
      finally
        TempArcad.Free;
      end;
      FDescStream.Position := FDescStream.Size;
      SaveToPageStream(#13 + #10);
      PageContent.SaveToStream(FDescStream);
      if (FParent.Compression = cmFlateDecode) then
      begin
        TmpStream := TMemoryStream.Create;
        try
//{$IFDEF D16}
          //with TZCompressionStream.Create(TmpStream, zcMax, -15) do
//{$ELSE}
          with TZCompressionStream.Create(TmpStream, zcMax) do
//{$ENDIF}
          begin
            FDescStream.Position := 0;
            CopyFrom(FDescStream, FDescStream.Size);
            Free;
          end;
          TmpStream.Position := 0;
          PageObjStream.Stream.Free;
          PageObjStream.Stream := TMemoryStream.Create;
          PageObjStream.Stream.CopyFrom(TmpStream, TmpStream.Size);
          PageObjStream.Length := TmpStream.Size;
        finally
          TmpStream.Free;
        end;
      end
      else
      begin
        FDescStream.Position := 0;
        PageObjStream.Stream.Free;
        PageObjStream.Stream := TMemoryStream.Create;
        PageObjStream.Stream.CopyFrom(FDescStream, 0);
        PageObjStream.Length := FDescStream.Size;
      end;
    finally
      FDescStream.Free;
    end;
  end;
  PageContent.Free;
end;

procedure THPDFPage.ConvertPageObject;
var
  I: Integer;
  XLen: Integer;
  InternPOStream: THPDFObject;
  ResouceIndex, XObjectIndex, FontIndex: Integer;
  ResouceObj: THPDFObject;
  ContentsIndex: Integer;
  ContentObj: THPDFObject;
  ResTmpLink: THPDFDictionaryObject;

  procedure GetPStreamFromArray(ArrObject: THPDFObject);
  var
    LinkObj: THPDFLink;
    ArrLen: Integer;
  begin
    ArrLen := THPDFArrayObject(ArrObject).Items.Count;
    LinkObj := THPDFLink(THPDFArrayObject(ArrObject).GetIndexedItem(ArrLen - 1));
    PageObjStream := THPDFStreamObject(FParent.GetObjectByLink(LinkObj));
  end;

var
  MBXMin: THPDFNumericObject;
  MBYMin: THPDFNumericObject;
  MBXMax: THPDFNumericObject;
  MBYMax: THPDFNumericObject;

begin
  ContentsIndex := PageObj.FindValue('Contents');
  if (ContentsIndex >= 0) then
  begin
    ContentObj := PageObj.GetIndexedItem(ContentsIndex);
    if (ContentObj.ObjectType = otArray) then
    begin
      GetPStreamFromArray(ContentObj);
    end
    else
    begin
      if (ContentObj.ObjectType = otLink) then
      begin
        InternPOStream := FParent.GetObjectByLink(THPDFLink(ContentObj));
        if (InternPOStream.ObjectType = otArray) then
        begin
          GetPStreamFromArray(InternPOStream);
        end
        else PageObjStream := THPDFStreamObject(InternPOStream);
      end
      else
       if (ContentObj.ObjectType = otStream) then
        PageObjStream := THPDFStreamObject(ContentObj);
    end;
  end;
  PageContent := TStringList.Create;
  SaveToPageStream(#13 + #10);
  ContentsIndex := PageObj.FindValue('MediaBox');
  if (ContentsIndex >= 0) then
  begin
    FMediaBoxArray := THPDFArrayObject(PageObj.GetIndexedItem(ContentsIndex));
    FMinXVal := THPDFNumericObject(FMediaBoxArray.Items.Items[0]).Value;
    FMinYVal := THPDFNumericObject(FMediaBoxArray.Items.Items[1]).Value;
    FMaxXVal := THPDFNumericObject(FMediaBoxArray.Items.Items[2]).Value;
    FMaxYVal := THPDFNumericObject(FMediaBoxArray.Items.Items[3]).Value;
    FWidth := FMaxXVal - FMinXVal;
    FHeight := FMaxYVal - FMinYVal;
    mWidth := FWidth;
    mHeight := FHeight;
  end
  else
  begin
    FMinXVal := FParent.FParentMB[0];
    FMinYVal := FParent.FParentMB[1];
    FMaxXVal := FParent.FParentMB[2];
    FMaxYVal := FParent.FParentMB[3];
    MBXMin := THPDFNumericObject.Create(nil);
    MBXMin.Value := FMinXVal;
    MBYMin := THPDFNumericObject.Create(nil);
    MBYMin.Value := FMinYVal;
    MBXMax := THPDFNumericObject.Create(nil);
    MBXMax.Value := FMaxXVal;
    MBYMax := THPDFNumericObject.Create(nil);
    MBYMax.Value := FMaxYVal;
    FMediaBoxArray := THPDFArrayObject.Create(nil);
    FMediaBoxArray.AddObject(THPDFObject(MBXMin));
    FMediaBoxArray.AddObject(THPDFObject(MBYMin));
    FMediaBoxArray.AddObject(THPDFObject(MBXMax));
    FMediaBoxArray.AddObject(THPDFObject(MBYMax));
    PageObj.AddValue('MediaBox', FMediaBoxArray);
    FWidth := FMaxXVal - FMinXVal;
    FHeight := FMaxYVal - FMinYVal;
    mWidth := FWidth;
    mHeight := FHeight;
  end;
  ResouceIndex := PageObj.FindValue('Resources');
  if (ResouceIndex >= 0) then
  begin
    ResouceObj := PageObj.GetIndexedItem(ResouceIndex);
    if (ResouceObj.ObjectType = otLink) then
    begin
      ResTmpLink := THPDFDictionaryObject(FParent.GetObjectByLink(THPDFLink(ResouceObj)));
      ResouceObj := THPDFObject(ResTmpLink);
    end;
    XObjectIndex := THPDFDictionaryObject(ResouceObj).FindValue('XObject');
    if (XObjectIndex >= 0) then
    begin
      XObjectObj := THPDFDictionaryObject(ResouceObj).GetIndexedItem(XObjectIndex);
      if (XObjectObj.ObjectType = otLink) then
      begin
        XObjectObj := THPDFObject(FParent.GetObjectByLink(THPDFLink(XObjectObj)));
      end;
      XLen := THPDFDictionaryObject(XObjectObj).Items.Count;
      SetLength(XObjectNames, XLen);
      if (XLen > 0) then
      begin
        for I := 0 to XLen - 1 do
        begin
          XObjectNames[I] := THPDFDictionaryObject(XObjectObj).GetIndexedKey(I);
        end;
      end;
    end
    else
      SetLength(XObjectNames, 0);
    FontIndex := THPDFDictionaryObject(ResouceObj).FindValue('Font');
    if (FontIndex >= 0) then
    begin
      FontObjectObj := THPDFDictionaryObject(ResouceObj).GetIndexedItem(FontIndex);
      if (FontObjectObj.ObjectType = otLink) then
      begin
        ResTmpLink := THPDFDictionaryObject(FParent.GetObjectByLink(THPDFLink(FontObjectObj)));
        FontObjectObj := THPDFObject(ResTmpLink);
      end;
      XLen := THPDFDictionaryObject(FontObjectObj).Items.Count;
      SetLength(FParent.FontNames, XLen);
      if (XLen > 0) then
      begin
        for I := 0 to XLen - 1 do
        begin
          FParent.FontNames[I] := THPDFDictionaryObject(FontObjectObj).GetIndexedKey(I);
        end;
      end;
    end
    else
    begin
      FontObjectObj := THPDFDictionaryObject.Create(nil);
      THPDFDictionaryObject(ResouceObj).AddValue('Font', FontObjectObj);
      SetLength(XObjectNames, 0);
    end;
  end
  else
    raise exception.Create('Invalid page resources.');
end;

function THPDFPage.CompareResName(Index: Integer; Name: AnsiString): Integer;
var
  I: Integer;
  LowCName: AnsiString;
  ArrVal: AnsiString;
  ArrLen: Integer;
begin
  result := -1;
  if (Index = 0) then
  begin
    LowCName := AnsiString(LowerCase(String(Name)));
    ArrLen := Length(FParent.XImages);
    if ArrLen > 0 then
    begin
      for I := 0 to ArrLen - 1 do
      begin
        ArrVal := FParent.XImages[I].Name;
        if (LowCName = AnsiString(LowerCase(String(ArrVal)))) then
        begin
          result := I;
          Exit;
        end;
      end;
    end;
    ArrLen := Length(XObjectNames);
    if ArrLen > 0 then
    begin
      for I := 0 to ArrLen - 1 do
      begin
        ArrVal := XObjectNames[I];
        if (LowCName = AnsiString(LowerCase(String(ArrVal)))) then
        begin
          result := I;
          Exit;
        end;
      end;
    end;
  end
  else
  begin
    LowCName := AnsiString(LowerCase(String(Name)));
    ArrLen := Length(FParent.FontNames);
    if ArrLen > 0 then
    begin
      for I := 0 to ArrLen - 1 do
      begin
        ArrVal := FParent.FontNames[I];
        if (LowCName = AnsiString(LowerCase(String(ArrVal)))) then
        begin
          result := I;
          Exit;
        end;
      end;
    end;
  end;
end;

procedure THPDFPage.CalculateFormat;
begin
  FUpdateSize := true;
  case FSize of
    psLetter:
      begin
        mHeight := 792;
        mWidth := 612;
      end;
    psA4:
      begin
        mHeight := 842;
        mWidth := 595;
      end;
    psA3:
      begin
        mHeight := 1190;
        mWidth := 842;
      end;
    psLegal:
      begin
        mHeight := 1008;
        mWidth := 612;
      end;
    psB5:
      begin
        mHeight := 728;
        mWidth := 516;
      end;
    psC5:
      begin
        mHeight := 649;
        mWidth := 459;
      end;
    ps8x11:
      begin
        mHeight := 792;
        mWidth := 576;
      end;
    psB4:
      begin
        mHeight := 1031;
        mWidth := 728;
      end;
    psA5:
      begin
        mHeight := 595;
        mWidth := 419;
      end;
    psFolio:
      begin
        mHeight := 936;
        mWidth := 612;
      end;
    psExecutive:
      begin
        mHeight := 756;
        mWidth := 522;
      end;
    psEnvB4:
      begin
        mHeight := 1031;
        mWidth := 728;
      end;
    psEnvB5:
      begin
        mHeight := 708;
        mWidth := 499;
      end;
    psEnvC6:
      begin
        mHeight := 459;
        mWidth := 323;
      end;
    psEnvDL:
      begin
        mHeight := 623;
        mWidth := 312;
      end;
    psEnvMonarch:
      begin
        mHeight := 540;
        mWidth := 279;
      end;
    psEnv9:
      begin
        mHeight := 639;
        mWidth := 279;
      end;
    psEnv10:
      begin
        mHeight := 684;
        mWidth := 297;
      end;
    psEnv11:
      begin
        mHeight := 747;
        mWidth := 324;
      end;
  else
    begin
      mHeight := Height;
      mWidth := Width;
    end;
  end;
  DPI := 72 / Resolution; //72
{$IFDEF BCB}
  if FOrientation = vpoPortrait then
{$ELSE}
  if FOrientation = poPortrait then
{$ENDIF}
  begin
    FHeight := mHeight;
    FWidth := mWidth;
  end
  else
  begin
    FHeight := mWidth;
    FWidth := mHeight;
    mHeight := FHeight;
    mWidth := FWidth;
  end;
  SetPageWidth(mWidth);
  SetPageHeight(mHeight);
  FHeight := trunc(mHeight / 72 * Resolution); //72
  FWidth := trunc(mWidth / 72 * Resolution); //72
  PageMeta.Height := Round(FHeight);
  PageMeta.Width := Round(FWidth);
  FUpdateSize := false;
end;

constructor THPDFPage.Create;
begin
  inherited Create;
  STextBegin := false;
  FDocScale := 1;
  FResolution := 72; //72
  DPI := 1;
  FHorizontalScaling := 100;
{$IFDEF BCB}
  FOrientation := vpoPortrait;
{$ELSE}
  FOrientation := poPortrait;
{$ENDIF}
  FKBegin := false;
  SKBegin := false;
  FAnnotsObj := nil;
  FHyperColor := clBlue;
  PageMeta := TMetafile.Create;
  PageMeta.Inch := FResolution;
  FCanvas := TMetafileCanvas.Create(PageMeta, 0);
  CurrentFontObj := THPDFFontObj.Create;
  FCanvas.Font.Name := 'Arial';
  FCanvas.Font.PixelsPerInch := 72; //72
end;

destructor THPDFPage.Destroy;
var
  I: Integer;
  FALen: Integer;
begin
  FALen := Length(FontArr);
  for I := 0 to FALen - 1 do
  begin
    StoreFont(FontArr[I]);
    FontArr[I].Free;
  end;
  CurrentFontObj.Free;
  FontArr := nil;
  inherited;
end;

procedure THPDFPage.StoreFont(FontObj: THPDFFontObj);
var
  I: Integer;
  CharsCount: WORD;
  LongLength: Integer;
  FBBox: THPDFArrayObject;
  TUStream: TStringStream;
  TmpStream, OutStream: TMemoryStream;
  FontBuffer: array of WORD;
  FWidths: THPDFArrayObject;
  FullName, ShortName: AnsiString;
  FDescent, FAbscent: Integer;
  AnsiTableStream: TTrueTypeTables;
  ToUnicode, FontStream2: THPDFStreamObject;
  FLOArr, FLArr, FDecf, WArray, KAASa: THPDFArrayObject;
  FLastChar, FFirstChar, CurWidth: Integer;
  DescendantFonts, FontDictObj, FontDescObj, CSInfo: THPDFDictionaryObject;

  function AddArrNameValue(FLArr: THPDFArrayObject; Val: AnsiString): THPDFNameObject;
  begin
    result := THPDFNameObject.Create(nil);
    result.Value := Val;
    FLArr.AddObject(result);
  end;

  function AddArrNumericValue(FLArr: THPDFArrayObject; Val: Integer): THPDFNumericObject;
  begin
    result := THPDFNumericObject.Create(nil);
    result.Value := Val;
    FLArr.AddObject(result);
  end;

begin
  Inc(FParent.FMaxObjNum);
  FontDictObj := THPDFDictionaryObject.Create(nil);
  FontDictObj.IsIndirect := true;
  FontDictObj.ID.ObjectNumber := FParent.FMaxObjNum;
  FParent.IndirectObjects.Add(FontDictObj);
  THPDFDictionaryObject(FontObjectObj).AddValue(FontObj.OrdinalName, FontDictObj);
  Inc(FParent.FMaxObjNum);
  FontDescObj := THPDFDictionaryObject.Create(nil);
  FontDescObj.IsIndirect := true;
  FontDescObj.ID.ObjectNumber := FParent.FMaxObjNum;
  FParent.IndirectObjects.Add(FontDescObj);
  ShortName := FontObj.Name;
  FullName := FontObj.Name;
  FFirstChar := 32;
  FLastChar := 255;
  FontDictObj.AddNameValue('Type', 'Font');
  FontDictObj.AddNameValue('BaseFont', FullName);
  if ((FontObj.IsUnicode) and (FParent.FontIsEmbedded(ShortName))) then
    FontDictObj.AddNameValue('Subtype', 'Type0')
  else
    FontDictObj.AddNameValue('Subtype', 'TrueType');
  if fsBold in FontObj.FontStyle then
    FullName := FullName + ' Bold';
  if fsItalic in FontObj.FontStyle then
    FullName := FullName + ' Italic';
  if FontObj.IsUnicode then
  begin
    if FontObj.IsVertical then
      FontDictObj.AddNameValue('Encoding', 'Identity-V')
    else
      FontDictObj.AddNameValue('Encoding', 'Identity-H');
  end
  else
  begin
    FontDictObj.AddNameValue('Encoding', 'WinAnsiEncoding');
    for I := 32 to 255 do
    begin
      if FontObj.SymbolTable[I] then
      begin
        FontDictObj.AddNumericValue('FirstChar', I);
        FFirstChar := I;
        break;
      end;
    end;
    if I <> 256 then
    begin
      for I := 255 downto 32 do
      begin
        if FontObj.SymbolTable[I] then
        begin
          FLastChar := I;
          FontDictObj.AddNumericValue('LastChar', I);
          break;
        end;
      end;
    end
    else
      FLastChar := 31;
  end;
  FontDictObj.AddNameValue('Name', FontObj.OrdinalName);
  FDescent := FontObj.OutTextM.otmDescent;
  FAbscent := FontObj.Ascent;
  if (not (FontObj.IsUnicode)) then
    FontDictObj.AddValue('FontDescriptor', FontDescObj);
  FontDescObj.AddNameValue('Type', 'FontDescriptor');
  FontDescObj.AddNumericValue('Ascent', FAbscent);
  FontDescObj.AddNumericValue('CapHeight', 666);
  FontDescObj.AddNumericValue('Descent', FDescent);
  if ((FontObj.IsStandard) and (FParent.FStandardFontEmulation) and
    (not(FontObj.IsUnicode)) and (FontObj.FontCharset = 0)) then
    FontDescObj.AddNumericValue('Flags', 32)
  else
    FontDescObj.AddNumericValue('Flags', 4);
  FBBox := THPDFArrayObject.Create(nil);
  AddArrNumericValue(FBBox, FontObj.OutTextM.otmrcFontBox.Left);
  AddArrNumericValue(FBBox, FontObj.OutTextM.otmrcFontBox.Bottom);
  AddArrNumericValue(FBBox, FontObj.OutTextM.otmrcFontBox.Right);
  AddArrNumericValue(FBBox, FontObj.OutTextM.otmrcFontBox.Top);
  FontDescObj.AddValue('FontBBox', FBBox);
  FontDescObj.AddNameValue('FontName', FontObj.Name);
  FontDescObj.AddNumericValue('ItalicAngle', FontObj.OutTextM.otmItalicAngle);
  FontDescObj.AddNumericValue('StemV', 87);
  if not((FontObj.IsStandard) and (FParent.FStandardFontEmulation) and
    (not(FontObj.IsUnicode)) and (FontObj.FontCharset = 0)) then
  begin
    if FParent.FontIsEmbedded(ShortName) or (FontObj.FontCharset <> 0) then
    begin
      FontStream2 := THPDFStreamObject.Create(nil);
      CharsCount := 0;
      LongLength := Length(FontObj.UnicodeTable);
      if FontObj.IsUnicode then
      begin
        for I := 0 to LongLength - 1 do
        begin
          if FontObj.UnicodeTable[I].CharCode then
          begin
            Inc(CharsCount);
            SetLength(FontBuffer, CharsCount);
            FontBuffer[CharsCount - 1] := I;
          end;
        end;
        if ((CharsCount = 1) and (FontBuffer[0] = 32)) then
        begin
          Inc(CharsCount);
          SetLength(FontBuffer, CharsCount);
          FontBuffer[CharsCount - 1] := 33;
        end;
      end
      else
      begin
        SetLength(FontBuffer, 224);
        for I := 32 to 255 do
        begin
          if FontObj.SymbolTable[I] then
          begin
            FontBuffer[CharsCount] := I;
            Inc(CharsCount);
          end;
        end;
      end;
      OutStream := TMemoryStream.Create;
      try
        AnsiTableStream := TTrueTypeTables.Create;
        try
          AnsiTableStream.CharSet := FontObj.FontCharset;
          if FontObj.IsUnicode then
            AnsiTableStream.GetFontLongEnsemble(String(FontObj.OldName),
              FontObj.FontStyle, OutStream, FontBuffer, CharsCount)
          else
            AnsiTableStream.GetFontEnsemble(String(FontObj.OldName),
              FontObj.FontStyle, OutStream, FontBuffer, CharsCount);
        finally
          AnsiTableStream.Free;
        end;
        I := OutStream.Size;
        OutStream.Position := 0;
        TmpStream := TMemoryStream.Create;
        try
//{$IFDEF D16}
          //with TZCompressionStream.Create(TmpStream, zcMax, -15) do
//{$ELSE}
          with TZCompressionStream.Create(TmpStream, zcMax) do
//{$ENDIF}
          begin
            CopyFrom(OutStream, 0);
            Free;
          end;
          TmpStream.Position := 0;
          FontStream2.Stream.CopyFrom(TmpStream, TmpStream.Size);
          FontStream2.Stream.Position := 0;
        finally
          TmpStream.Free;
        end;
      finally
        OutStream.Free;
      end;
      FLArr := THPDFArrayObject.Create(nil);
      AddArrNameValue(FLArr, 'FlateDecode');
      FontStream2.Dictionary.AddValue('Filter', FLArr);
      FontStream2.Dictionary.AddNumericValue('Length1', I);
      Inc(FParent.FMaxObjNum);
      FontStream2.ID.ObjectNumber := FParent.FMaxObjNum;
      FontStream2.IsIndirect := true;
      FParent.IndirectObjects.Add(FontStream2);
      FontDescObj.AddValue('FontFile2', FontStream2);
      if FontObj.IsUnicode then
      begin
        DescendantFonts := THPDFDictionaryObject.Create(nil);
        DescendantFonts.IsIndirect := true;
        Inc(FParent.FMaxObjNum);
        DescendantFonts.ID.ObjectNumber := FParent.FMaxObjNum;
        FParent.IndirectObjects.Add(DescendantFonts);
        DescendantFonts.AddNameValue('Type', 'Font');
        DescendantFonts.AddNameValue('BaseFont', FullName);
        DescendantFonts.AddNameValue('Subtype', 'CIDFontType2');
        DescendantFonts.AddValue('FontDescriptor', FontDescObj);
        FDecf := THPDFArrayObject.Create(nil);
        FDecf.AddObject(DescendantFonts);
        FontDictObj.AddValue('DescendantFonts', FDecf);
        WArray := THPDFArrayObject.Create(nil);
        for I := 0 to LongLength - 1 do
        begin
          if FontObj.UnicodeTable[I].CharCode then
          begin
            AddArrNumericValue(WArray, FontObj.Symbols[I].Index);
            KAASa := THPDFArrayObject.Create(nil);
            AddArrNumericValue(KAASa, FontObj.Symbols[I].Width);
            WArray.AddObject(KAASa);
          end;
        end;
        DescendantFonts.AddValue('W', WArray);
        CSInfo := THPDFDictionaryObject.Create(nil);
        CSInfo.AddStringValue('Registry', 'Adobe');
        CSInfo.AddStringValue('Ordering', 'Identity');
        CSInfo.AddNumericValue('Supplement', 0);
        DescendantFonts.AddValue('CIDSystemInfo', CSInfo);
        ToUnicode := THPDFStreamObject.Create(nil);
        ToUnicode.IsIndirect := true;
        TUStream := TStringStream.Create('');
        try
          TUStream.WriteString('/CIDInit /ProcSet findresource begin 12 dict begin begincmap /CIDSystemInfo <<' + #13 + #10);
          TUStream.WriteString('/Registry (' + String(FontObj.Name) + ') /Ordering (T42UV) /Supplement 0 >> def' + #13 + #10);
          TUStream.WriteString('/CMapName /' + String(FontObj.Name) + ' def' + #13 + #10 + '/CMapType 2 def' + #13 + #10);
          TUStream.WriteString('1 begincodespacerange <023B> <024F> endcodespacerange' + #13 + #10);
          TUStream.WriteString(IntToStr(FontObj.UniLen) + ' beginbfchar' + #13 + #10);
          for I := 0 to LongLength - 1 do
          begin
            if FontObj.UnicodeTable[I].CharCode then
              TUStream.WriteString('<' + IntToHex(FontObj.Symbols[I].Index, 4) + '> <' + IntToHex(I, 4) + '>' + #13 + #10);
          end;
          TUStream.WriteString('endbfchar' + #13 + #10);
          TUStream.WriteString('endcmap CMapName currentdict /CMap defineresource pop end end' + #13 + #10);
          I := TUStream.Size;
          TUStream.Position := 0;
          TmpStream := TMemoryStream.Create;
          try
//{$IFDEF D16}
            //with TZCompressionStream.Create(TmpStream, zcMax, -15) do
//{$ELSE}
            with TZCompressionStream.Create(TmpStream, zcMax) do
//{$ENDIF}
            begin
              CopyFrom(TUStream, 0);
              Free;
            end;
            TmpStream.Position := 0;
            ToUnicode.Stream.CopyFrom(TmpStream, TmpStream.Size);
            ToUnicode.Stream.Position := 0;
          finally
            TmpStream.Free;
          end;
        finally
          TUStream.Free;
        end;
        FLOArr := THPDFArrayObject.Create(nil);
        AddArrNameValue(FLOArr, 'FlateDecode');
        ToUnicode.Dictionary.AddValue('Filter', FLOArr);
        ToUnicode.Dictionary.AddNumericValue('Length1', I);
        Inc(FParent.FMaxObjNum);
        ToUnicode.ID.ObjectNumber := FParent.FMaxObjNum;
        FParent.IndirectObjects.Add(ToUnicode);
        FontDictObj.AddValue('ToUnicode', ToUnicode);

      end;
    end;
  end
  else
    FontDictObj.AddNameValue('Encoding', 'WinAnsiEncoding');
  if (not (FontObj.IsUnicode)) then
  begin
    FWidths := THPDFArrayObject.Create(nil);
    if not (FontObj.IsUnicode) then
    begin
      if FLastChar <> 31 then
      begin
        for I := FFirstChar to FLastChar do
        begin
          if FontObj.SymbolTable[I] then
          begin
            if not FontObj.IsMonospaced then
              CurWidth := FontObj.ABCArray[I].abcA +
                Integer(FontObj.ABCArray[I].abcB) + FontObj.ABCArray[I].abcC
            else
              CurWidth := FontObj.ABCArray[0].abcA +
                Integer(FontObj.ABCArray[0].abcB) + FontObj.ABCArray[0].abcC;
            AddArrNumericValue(FWidths, CurWidth);
          end
          else
            AddArrNumericValue(FWidths, 0);
        end;
      end
      else
        AddArrNumericValue(FWidths, 0);
    end
    else
    begin
      AddArrNumericValue(FWidths, 656);
      AddArrNumericValue(FWidths, 719);
      AddArrNumericValue(FWidths, 667);
    end;
    FontDictObj.AddValue('Widths', FWidths);
  end;
end;

procedure THPDFPage.SetResolution(const Value: Integer);
begin
  FResolution := Value;
  CalculateFormat;
end;

function THPDFPage.GetPageHeight: Single;
begin
  result := FHeight;
end;

function THPDFPage.GetPageWidth: Single;
begin
  result := FWidth;
end;

procedure THPDFPage.SetPageHeight(const Value: Single);
var
  I: Integer;
  SVal: Single;
begin
  SVal := Value + THPDFNumericObject(FMediaBoxArray.Items.Items[1]).Value;
  FMaxYVal := SVal;
  if not (FUpdateSize) then
  begin
    THPDFNumericObject(FMediaBoxArray.Items.Items[3]).Value := SVal;
    FSize := psUserDefined;
    FHeight := Value;
    CalculateHeightFormat;
  end
  else
    THPDFNumericObject(FMediaBoxArray.Items.Items[3]).Value := SVal;
  if PageMeta <> nil then
  begin
    begin
      I := GetDeviceCaps(FParent.FCHandle, LOGPIXELSY);
      PageMeta.Height := MulDiv(Round(FHeight), I, 72); //72
      FCanvas.Free;
      FCanvas := TMetafileCanvas.Create(PageMeta, FParent.FCHandle);
    end;
  end;
end;

procedure THPDFPage.SetPageWidth(const Value: Single);
var
  I: Integer;
  SVal: Single;
begin
  SVal := Value + THPDFNumericObject(FMediaBoxArray.Items.Items[0]).Value;
  FMaxXVal := SVal;
  if not (FUpdateSize) then
  begin
    THPDFNumericObject(FMediaBoxArray.Items.Items[2]).Value := SVal;
    FSize := psUserDefined;
    FWidth := Value;
    CalculateWidthFormat;
  end
  else
    THPDFNumericObject(FMediaBoxArray.Items.Items[2]).Value := SVal;
  if PageMeta <> nil then
  begin
    I := GetDeviceCaps(FParent.FCHandle, LOGPIXELSX);
    PageMeta.Width := MulDiv(Round(FWidth), I, 72); //72
    FCanvas.Free;
    FCanvas := TMetafileCanvas.Create(PageMeta, FParent.FCHandle);
  end;
end;

procedure THPDFPage.SetPageSize(const Value: THPDFPageSize);
begin
  FSize := Value;
  CalculateFormat;
end;

procedure THPDFPage.SetFontAndSize(FontName: AnsiString; Size: Single);
begin
  SaveToPageStream('/' + FontName + ' ' + _FloatToStrR(Size) + ' Tf');
end;

function THotPDF.CopyObject(SourceDoc: THotPDF; SourceObj: THPDFObject): THPDFObject;
var
  I: Integer;
  ObjInd: Integer;
  SelKey, TpName: AnsiString;
  CSLen: Integer;
  LgObj: THPDFObject;
  TpObj: THPDFObject;
  ArrItemObj: THPDFObject;
  DictItemObj: THPDFObject;
  NullObj: THPDFNullObject;
  BooleanObj: THPDFBooleanObject;
  NumericObj: THPDFNumericObject;
  StringObj: THPDFStringObject;
  NameObj: THPDFNameObject;
  StreamObj: THPDFStreamObject;
  ArrayObj: THPDFArrayObject;
  DictionaryObj: THPDFDictionaryObject;

  procedure LogNewObj;
  begin
    if (SourceObj.ID.ObjectNumber <> 0) then
    begin
      SetLength(CSArray, CSLen + 1);
      CSArray[CSLen].ObjectNumber.ObjectNumber := SourceObj.ID.ObjectNumber;
      CSArray[CSLen].ObjectNumber.GenerationNumber := SourceObj.ID.GenerationNumber;
      CSArray[CSLen].ObjectBody := LgObj;
    end;
  end;

begin
  CSLen := Length(CSArray);
  result := nil;
  if (SourceObj.IsIndirect) then
  begin
    for I := 0 to CSLen - 1 do
    begin
      if ((CSArray[I].ObjectNumber.ObjectNumber = SourceObj.ID.ObjectNumber) and (CSArray[I].ObjectNumber.GenerationNumber = SourceObj.ID.GenerationNumber)) then
      begin
        result := CSArray[I].ObjectBody;
        break;
      end;
    end;
  end;
  if (result = nil) then
  begin
    case SourceObj.ObjectType of
      otNull:
        begin
          NullObj := THPDFNullObject.Create;
          if (SourceObj.IsIndirect) then
          begin
            Inc(FMaxObjNum);
            NullObj.ID.ObjectNumber := FMaxObjNum;
            NullObj.IsIndirect := true;
            IndirectObjects.Add(NullObj);
            LgObj := THPDFObject(NullObj);
            LogNewObj;
          end;
          result := THPDFObject(NullObj);
        end;
      otBoolean:
        begin
          BooleanObj := THPDFBooleanObject.Create(nil);
          if (SourceObj.IsIndirect) then
          begin
            Inc(FMaxObjNum);
            BooleanObj.ID.ObjectNumber := FMaxObjNum;
            BooleanObj.IsIndirect := true;
            IndirectObjects.Add(BooleanObj);
            LgObj := THPDFObject(BooleanObj);
            LogNewObj;
          end;
          BooleanObj.Value := THPDFBooleanObject(SourceObj).Value;
          result := THPDFObject(BooleanObj);
        end;
      otNumeric:
        begin
          NumericObj := THPDFNumericObject.Create(nil);
          if (SourceObj.IsIndirect) then
          begin
            Inc(FMaxObjNum);
            NumericObj.ID.ObjectNumber := FMaxObjNum;
            NumericObj.IsIndirect := true;
            IndirectObjects.Add(NumericObj);
            LgObj := THPDFObject(NumericObj);
            LogNewObj;
          end;
          NumericObj.Value := THPDFNumericObject(SourceObj).Value;
          result := THPDFObject(NumericObj);
        end;
      otString:
        begin
          StringObj := THPDFStringObject.Create(nil);
          if (SourceObj.IsIndirect) then
          begin
            Inc(FMaxObjNum);
            StringObj.ID.ObjectNumber := FMaxObjNum;
            StringObj.IsIndirect := true;
            IndirectObjects.Add(StringObj);
            LgObj := THPDFObject(StringObj);
            LogNewObj;
          end;
          StringObj.Value := THPDFStringObject(SourceObj).Value;
          result := THPDFObject(StringObj);
        end;
      otName:
        begin
          NameObj := THPDFNameObject.Create(nil);
          if (SourceObj.IsIndirect) then
          begin
            Inc(FMaxObjNum);
            NameObj.ID.ObjectNumber := FMaxObjNum;
            NameObj.IsIndirect := true;
            IndirectObjects.Add(NameObj);
            LgObj := THPDFObject(NameObj);
            LogNewObj;
          end;
          NameObj.Value := THPDFNameObject(SourceObj).Value;
          result := THPDFObject(NameObj);
        end;
      otArray:
        begin
          ArrayObj := THPDFArrayObject.Create(nil);
          if (SourceObj.IsIndirect) then
          begin
            Inc(FMaxObjNum);
            ArrayObj.ID.ObjectNumber := FMaxObjNum;
            ArrayObj.IsIndirect := true;
            IndirectObjects.Add(ArrayObj);
            LgObj := THPDFObject(ArrayObj);
            LogNewObj;
          end;
          ObjInd := 0;
          while (ObjInd < THPDFArrayObject(SourceObj).Items.Count) do
          begin
            ArrItemObj := THPDFArrayObject(SourceObj).GetIndexedItem(ObjInd);
            ArrayObj.AddObject(CopyObject(SourceDoc, ArrItemObj));
            Inc(ObjInd);
          end;
          result := THPDFObject(ArrayObj);
        end;
      otDictionary:
        begin
          ObjInd := THPDFDictionaryObject(SourceObj).FindValue('Type');
          if (ObjInd >= 0) then
          begin
            TpObj := THPDFDictionaryObject(SourceObj).GetIndexedItem(ObjInd);
            if (TpObj.ObjectType = otName) then
            begin
              TpName := THPDFNameObject(TpObj).Value;
              if (LowerCase(String(TpName)) = 'page') then
              begin
                result := nil;
                Exit;
              end;
            end;
          end;
          DictionaryObj := THPDFDictionaryObject.Create(nil);
          if (SourceObj.IsIndirect) then
          begin
            Inc(FMaxObjNum);
            DictionaryObj.ID.ObjectNumber := FMaxObjNum;
            DictionaryObj.IsIndirect := true;
            IndirectObjects.Add(DictionaryObj);
            LgObj := THPDFObject(DictionaryObj);
            LogNewObj;
          end;
          ObjInd := 0;
          while (ObjInd < THPDFDictionaryObject(SourceObj).Items.Count) do
          begin
            DictItemObj := THPDFDictionaryObject(SourceObj).GetIndexedItem(ObjInd);
            SelKey := THPDFDictionaryObject(SourceObj).GetIndexedKey(ObjInd);
            if (SelKey <> 'B') then
              DictionaryObj.AddValue(SelKey, CopyObject(SourceDoc, DictItemObj));
            Inc(ObjInd);
          end;
          result := THPDFObject(DictionaryObj);
        end;
      otStream:
        begin
          StreamObj := THPDFStreamObject.Create(nil);
          if (SourceObj.IsIndirect) then
          begin
            Inc(FMaxObjNum);
            StreamObj.ID.ObjectNumber := FMaxObjNum;
            StreamObj.IsIndirect := true;
            IndirectObjects.Add(StreamObj);
            LgObj := THPDFObject(StreamObj);
            LogNewObj;
          end;
          ObjInd := 0;
          while (ObjInd < THPDFStreamObject(SourceObj).Dictionary.Items.Count) do
          begin
            SelKey := THPDFStreamObject(SourceObj).Dictionary.GetIndexedKey(ObjInd);
            StreamObj.Dictionary.AddValue(SelKey, CopyObject(SourceDoc, THPDFStreamObject(SourceObj).Dictionary.GetIndexedItem(ObjInd)));
            Inc(ObjInd);
          end;
          THPDFStreamObject(SourceObj).Stream.Position := 0;
          StreamObj.Stream.CopyFrom(THPDFStreamObject(SourceObj).Stream, 0);
          result := THPDFObject(StreamObj);
        end;
    else
      begin
        Result := CopyObject(SourceDoc, SourceDoc.GetObjectByLink(THPDFLink(SourceObj)));
        LgObj := result;
        LogNewObj;
      end;
    end;
  end;
end;

procedure THotPDF.CopyPageFromDocument(SourceDoc: THotPDF; SourceIndex: Integer; DestIndex: Integer);
var
  I, BockId, ParentInd, AcLen: Integer;
  SourcePageObj, DestPageObj, PagesObj: THPDFDictionaryObject;
  ItIndKey: AnsiString;
  KidsArr: THPDFArrayObject;
  CountP: THPDFNumericObject;
  TmpArrItm, TmpArrItm1: THPDFDictArrItem;
begin
  Inc(FPagesCount);
  if (SourceIndex > SourceDoc.FPagesCount - 1) then
  begin
    raise exception.Create('Invalid source page number');
  end;
  if (DestIndex > FPagesCount - 1) then
  begin
    raise exception.Create('Invalid destination page number');
  end;
  SourcePageObj := SourceDoc.PageArr[SourceIndex].PageObj;
  DestPageObj := CreateIndirectDictionary;
  ParentInd := SourcePageObj.FindValue('Parent');
  AcLen := Length(CSArray);
  SetLength(CSArray, AcLen + 1);
  CSArray[AcLen].ObjectNumber.ObjectNumber := SourcePageObj.ID.ObjectNumber;
  CSArray[AcLen].ObjectNumber.GenerationNumber := SourcePageObj.ID.GenerationNumber;
  CSArray[AcLen].ObjectBody := DestPageObj;
  for I := 0 to SourcePageObj.Items.Count - 1 do
  begin
    if (I <> ParentInd) then
    begin
      ItIndKey := SourcePageObj.GetIndexedKey(I);
      if (ItIndKey <> 'B') then
        DestPageObj.AddValue(ItIndKey, CopyObject(SourceDoc, SourcePageObj.GetIndexedItem(I)));
    end;
  end;
  CSArray := nil;
  PagesObj := THPDFDictionaryObject(IndirectObjects.Items[FPagesIndex]);
  DestPageObj.AddValue('Parent', PagesObj);
  BockId := PagesObj.FindValue('Kids');
  KidsArr := THPDFArrayObject(PagesObj.GetIndexedItem(BockId));
  KidsArr.SetIndexedItem(DestIndex, DestPageObj);
  BockId := PagesObj.FindValue('Count');
  CountP := THPDFNumericObject(PagesObj.GetIndexedItem(BockId));
  CountP.Value := CountP.Value + 1;
  SetLength(PageArr, FPagesCount);
  TmpArrItm.PageObj := PageArr[DestIndex].PageObj;
  TmpArrItm.PageLink.ObjectNumber := PageArr[DestIndex].PageLink.ObjectNumber;
  TmpArrItm.PageLink.GenerationNumber := PageArr[DestIndex].PageLink.GenerationNumber;
  PageArr[DestIndex].PageObj := DestPageObj;
  PageArr[DestIndex].PageLink.ObjectNumber := DestPageObj.ID.ObjectNumber;
  PageArr[DestIndex].PageLink.GenerationNumber := DestPageObj.ID.GenerationNumber;
  I := DestIndex + 1;
  while (I < FPagesCount) do
  begin
    TmpArrItm1.PageObj := PageArr[I].PageObj;
    TmpArrItm1.PageLink.ObjectNumber := PageArr[I].PageLink.ObjectNumber;
    TmpArrItm1.PageLink.GenerationNumber := PageArr[I].PageLink.GenerationNumber;
    PageArr[I].PageObj := TmpArrItm.PageObj;
    PageArr[I].PageLink.ObjectNumber := TmpArrItm.PageLink.ObjectNumber;
    PageArr[I].PageLink.GenerationNumber := TmpArrItm.PageLink.GenerationNumber;
    TmpArrItm.PageObj := TmpArrItm1.PageObj;
    TmpArrItm.PageLink.ObjectNumber := TmpArrItm1.PageLink.ObjectNumber;
    TmpArrItm.PageLink.GenerationNumber := TmpArrItm1.PageLink.GenerationNumber;
    Inc(I);
  end;
  FCurrentPageNum := -1;
  SetCurrentPageNum(DestIndex);
end;

procedure THotPDF.BeginParagraph(Index: Integer);
begin
  if ((Index > FParaLen) or (Index < 0)) then
  begin
    raise exception.Create('Incorrect paragraph index.');
  end
  else
  begin
    FActivePara := Index;
    FCurrentParaLine := FParas[Index - 1].TopMargin;
    FNewParaLine := true;
    FPrevParaOffset := 0;
    FCurrentParagraph := THPDFPara.Create(Self);
    FCurrentParagraph.FJustification := FParas[Index - 1].Justification;
    FCurrentParagraph.Indention := FParas[Index - 1].Indention;
    FCurrentParagraph.LeftMargin := FParas[Index - 1].LeftMargin;
    FCurrentParagraph.RightMargin := FParas[Index - 1].RightMargin;
    FCurrentParagraph.TopMargin := FParas[Index - 1].TopMargin;
    FCurrentParagraph.BottomMargin := FParas[Index - 1].BottomMargin;
  end;
end;

procedure THotPDF.EndParagraph;
begin
  FCurrentParagraph.Free;
  FActivePara := 0;
end;

procedure THPDFPage.BeginText;
begin
  if (STextBegin) then
    Exit;
  SaveToPageStream('BT');
  STextBegin := true;
end;

procedure THPDFPage.EndText;
begin
  if (not(STextBegin)) then
    Exit;
  SaveToPageStream('ET');
  STextBegin := false;
end;

procedure THPDFPage.StoreCurrentFont;
var
  FRLen: Integer;
begin
  if (not (CurrentFontObj.Saved)) then
  begin
    FRLen := Length(FontArr);
    Inc(FRLen);
    SetLength(FontArr, FRLen);
    FontArr[FRLen - 1] := THPDFFontObj.Create;
    FontArr[FRLen - 1].Size := CurrentFontObj.Size;
    FontArr[FRLen - 1].Ascent := CurrentFontObj.Ascent;
    FontArr[FRLen - 1].FontLen := CurrentFontObj.FontLen;
    FontArr[FRLen - 1].Name := CurrentFontObj.Name;
    FontArr[FRLen - 1].OldName := CurrentFontObj.OldName;
    FontArr[FRLen - 1].FontStyle := CurrentFontObj.FontStyle;
    FontArr[FRLen - 1].FontCharset := CurrentFontObj.FontCharset;
    FontArr[FRLen - 1].IsVertical := CurrentFontObj.IsVertical;
    FontArr[FRLen - 1].IsUnicode := CurrentFontObj.IsUnicode;
    FontArr[FRLen - 1].IsUsed := CurrentFontObj.IsUsed;
    FontArr[FRLen - 1].OrdinalName := CurrentFontObj.OrdinalName;
    FontArr[FRLen - 1].IsMonospaced := CurrentFontObj.IsMonospaced;
    FontArr[FRLen - 1].IsStandard := CurrentFontObj.IsStandard;
    FontArr[FRLen - 1].CopyFontFetures(CurrentFontObj);
    CurrentFontObj.Saved := true;
  end
  else
  begin
    FRLen := CurrentFontObj.ArrIndex;
    FontArr[FRLen].CopyFontFetures(CurrentFontObj);
  end;
end;

function THPDFPage.GetCanvas: TCanvas;
begin
  result := FCanvas;
end;

function THPDFPage.CompareCurrentFont: Integer;
var
  I: Integer;
  FRLen: Integer;

begin
  result := -1;
  FRLen := Length(FontArr);
  for I := 0 to FRLen - 1 do
  begin
    if ((FontArr[I].Name = CurrentFontObj.Name) and
      (FontArr[I].FontStyle = CurrentFontObj.FontStyle) and
      (FontArr[I].FontCharset = CurrentFontObj.FontCharset) and
      (FontArr[I].IsVertical = CurrentFontObj.IsVertical) and
      (FontArr[I].IsUnicode = CurrentFontObj.IsUnicode)) then
    begin
      result := I;
      break;
    end;
  end;
end;

procedure THPDFPage.ProcessFont(Unicoded: boolean);
var
  NewObjName: AnsiString;
  CFIndex: Integer;

  procedure ExtractFontDat;
  var
    I: Integer;
  begin
    NewObjName := FontArr[CFIndex].OrdinalName;
    CurrentFontObj.ArrIndex := CFIndex;
    for I := 32 to 255 do
      CurrentFontObj.SymbolTable[I] := FontArr[CFIndex].SymbolTable[I];
    CurrentFontObj.Saved := true;
  end;

  procedure UpdateFontNames;
  var
    ArrLen: Integer;
  begin
    ArrLen := Length(FParent.FontNames);
    SetLength(FParent.FontNames, ArrLen + 1);
    FParent.FontNames[ArrLen] := NewObjName;
  end;

begin
  if (Unicoded) then
  begin
    if (CurrentFontObj.IsUnicode) then
    begin
      if (not (CurrentFontObj.IsUsed)) then
      begin
        CFIndex := CompareCurrentFont;
        if (CFIndex >= 0) then
        begin
          ExtractFontDat;
        end
        else
        begin
          CurrentFontObj.Saved := false;
          NewObjName := 'F' + AnsiString(IntToStr(FParent.FCurrentFontIndex));
          while (CompareResName(1, NewObjName) > -1) do
          begin
            Inc(FParent.FCurrentFontIndex);
            NewObjName := 'F' + AnsiString(IntToStr(FParent.FCurrentFontIndex));
          end;
        end;
        CurrentFontObj.OrdinalName := NewObjName;
        UpdateFontNames;
        SetFontAndSize(NewObjName, CurrentFontObj.Size);
        CurrentFontObj.IsUsed := true;
      end;
    end
    else
    begin
      if (CurrentFontObj.IsUsed) then
        StoreCurrentFont;
      CurrentFontObj.IsUnicode := true;
      CFIndex := CompareCurrentFont;
      if (CFIndex >= 0) then
      begin
        ExtractFontDat;
      end
      else
      begin
        CurrentFontObj.Saved := false;
        NewObjName := 'F' + AnsiString(IntToStr(FParent.FCurrentFontIndex));
        while (CompareResName(1, NewObjName) > -1) do
        begin
          Inc(FParent.FCurrentFontIndex);
          NewObjName := 'F' + AnsiString(IntToStr(FParent.FCurrentFontIndex));
        end;
      end;
      CurrentFontObj.OrdinalName := NewObjName;
      UpdateFontNames;
      SetFontAndSize(NewObjName, CurrentFontObj.Size);
      CurrentFontObj.IsUsed := true;
      CurrentFontObj.FActive := false;
    end;
  end
  else
  begin
    if (CurrentFontObj.IsUnicode) then
    begin
      if (CurrentFontObj.IsUsed) then
        StoreCurrentFont;
      CFIndex := CompareCurrentFont;
      if (CFIndex >= 0) then
      begin
        ExtractFontDat;
      end
      else
      begin
        CurrentFontObj.Saved := false;
        CurrentFontObj.IsUnicode := false;
        NewObjName := 'F' + AnsiString(IntToStr(FParent.FCurrentFontIndex));
        while (CompareResName(1, NewObjName) > -1) do
        begin
          Inc(FParent.FCurrentFontIndex);
          NewObjName := 'F' + AnsiString(IntToStr(FParent.FCurrentFontIndex));
        end;
      end;
      CurrentFontObj.OrdinalName := NewObjName;
      UpdateFontNames;
      SetFontAndSize(NewObjName, CurrentFontObj.Size);
      CurrentFontObj.IsUsed := true;
    end
    else
    begin
      if (not (CurrentFontObj.IsUsed)) then
      begin
        CFIndex := CompareCurrentFont;
        if (CFIndex >= 0) then
        begin
          ExtractFontDat;
        end
        else
        begin
          CurrentFontObj.Saved := false;
          NewObjName := 'F' + AnsiString(IntToStr(FParent.FCurrentFontIndex));
          while (CompareResName(1, NewObjName) > -1) do
          begin
            Inc(FParent.FCurrentFontIndex);
            NewObjName := 'F' + AnsiString(IntToStr(FParent.FCurrentFontIndex));
          end;
        end;
        CurrentFontObj.OrdinalName := NewObjName;
        UpdateFontNames;
        SetFontAndSize(NewObjName, CurrentFontObj.Size);
        CurrentFontObj.IsUsed := true;
      end;
    end;
  end;
end;

{$IFDEF BCB}

procedure THPDFPage.OutText(X, Y: Single; angle: Extended; Text: AnsiString);
var
  I: Integer;
  MtxA, MtxB: Single;
  Abscissa: Single;
  StrHeight: Single;
  YProj: Single;
begin
  ProcessFont(false);
  for I := 1 to Length(Text) do
  begin
    if Ord(Text[I]) = 173 then
      Text[I] := chr(45);
    if (Ord(Text[I]) = 160) or (Ord(Text[I]) = 152) or (Ord(Text[I]) = 127) then
      Text[I] := chr(32);
    CurrentFontObj.SymbolTable[Ord(Text[I])] := true;
  end;
  BeginText;
  StrHeight := TextHeight(Text);
  if FTopTextPosition then
    YProj := 0
  else
    YProj := StrHeight;
  if angle <> 0 then
  begin
    Abscissa := angle * Pi / 180;
    X := XProjection(X) + StrHeight * sin(Abscissa);
    Y := (YProjection(Y)) - StrHeight * cos(Abscissa);
    MtxA := cos(Abscissa);
    MtxB := sin(Abscissa);
    SetTextMatrix(MtxA, MtxB, -MtxB, MtxA, X, Y);
  end
  else
    MoveTextPoint(X, Y + YProj);
  ShowText(Text);
  EndText;
  if angle <> 0 then
  begin
    X := XProjection(X);
    Y := YProjection(Y);
    YProj := 0;
  end;
  if (fsUnderline in CurrentFontObj.FontStyle) then
  begin
    RectangleRotate(X + 2 * sin((Pi / 180) * angle),
      Y + 2 * cos(angle * (Pi / 180)) + YProj, TextWidth(Text),
      CurrentFontObj.Size * 0.07, -1 * angle);
    Fill;
  end;
  if (fsStrikeOut in CurrentFontObj.FontStyle) then
  begin
    RectangleRotate(X - (StrHeight / 3) * sin((Pi / 180) * angle),
      Y - (StrHeight / 3) * cos(angle * (Pi / 180)) + YProj, TextWidth(Text),
      CurrentFontObj.Size * 0.07, -1 * angle);
    Fill;
  end;
  StoreCurrentFont;
  CurrentFontObj.IsUsed := false;
end;

procedure THPDFPage.PrintText(X, Y: Single; angle: Extended; Text: AnsiString);
{$ELSE}

procedure THPDFPage.TextOut(X, Y: Single; angle: Extended; Text: AnsiString);
{$ENDIF}
var
  I: Integer;
  MtxA, MtxB: Single;
  Abscissa: Single;
  StrHeight: Single;
  YProj: Single;
begin
  ProcessFont(false);
  for I := 1 to Length(Text) do
  begin
    if Ord(Text[I]) = 173 then
      Text[I] := chr(45);
    if (Ord(Text[I]) = 160) or (Ord(Text[I]) = 152) or (Ord(Text[I]) = 127) then
      Text[I] := chr(32);
    CurrentFontObj.SymbolTable[Ord(Text[I])] := true;
  end;
  BeginText;
  StrHeight := TextHeight(Text);
  if FTopTextPosition then
    YProj := 0
  else
    YProj := StrHeight;
  if angle <> 0 then
  begin
    Abscissa := angle * Pi / 180;
    X := XProjection(X) + StrHeight * sin(Abscissa);
    Y := (YProjection(Y)) - StrHeight * cos(Abscissa);
    MtxA := cos(Abscissa);
    MtxB := sin(Abscissa);
    SetTextMatrix(MtxA, MtxB, -MtxB, MtxA, X, Y);
  end
  else
    MoveTextPoint(X, Y + YProj);
  ShowText(Text);
  EndText;
  if angle <> 0 then
  begin
    X := XProjection(X);
    Y := YProjection(Y);
    YProj := 0;
  end;
  if (fsUnderline in CurrentFontObj.FontStyle) then
  begin
    RectangleRotate(X + 2 * sin((Pi / 180) * angle),
      Y + 2 * cos(angle * (Pi / 180)) + YProj, TextWidth(Text),
      CurrentFontObj.Size * 0.07, -1 * angle);
    Fill;
  end;
  if (fsStrikeOut in CurrentFontObj.FontStyle) then
  begin
    RectangleRotate(X - (StrHeight / 3) * sin((Pi / 180) * angle),
      Y - (StrHeight / 3) * cos(angle * (Pi / 180)) + YProj, TextWidth(Text),
      CurrentFontObj.Size * 0.07, -1 * angle);
    Fill;
  end;
  StoreCurrentFont;
  CurrentFontObj.IsUsed := false;
end;

procedure THPDFPage.ShowImage(ImageIndex: Integer; X, Y, w, h, angle: Extended);
var
  I: Integer;
  XODict: THPDFDictionaryObject;
  BlockInd: Integer;
  CurrIm: THPDFObject;
  XOLink: THPDFLink;
  CurrImName: AnsiString;
  ImWZoom: Extended;
  ImHZoom: Extended;
  ResourDict: THPDFDictionaryObject;
begin
  I := ImageIndex;
  CurrImName := FParent.XImages[I].Name;
  CurrIm := FParent.XImages[I].ImageObject;
  XOLink := THPDFLink.Create;
  XOLink.Value.ObjectNumber := CurrIm.ID.ObjectNumber;
  XOLink.Value.GenerationNumber := CurrIm.ID.GenerationNumber;
  if (XObjectObj = nil) then
  begin
    XODict := THPDFDictionaryObject.Create(nil);
    XObjectObj := THPDFObject(XODict);
    BlockInd := PageObj.FindValue('Resources');
    if (BlockInd >= 0) then
    begin
      ResourDict := THPDFDictionaryObject(PageObj.GetIndexedItem(BlockInd));
      ResourDict.AddValue('XObject', XODict);
    end;
  end;
  THPDFDictionaryObject(XObjectObj).AddValue(CurrImName, XOLink);
  if (FParent.FKeepImageAspectRatio) then
  begin
    ImWZoom := w / FParent.FSizes[ImageIndex].Width;
    ImHZoom := h / FParent.FSizes[ImageIndex].heigh;
    if ((ImWZoom < 1) and (ImHZoom < 1)) then
    begin
      if (ImWZoom < ImHZoom) then
        h := FParent.FSizes[ImageIndex].heigh * ImWZoom
      else
        w := FParent.FSizes[ImageIndex].Width * ImHZoom;
    end
    else
    begin
      if (ImWZoom > ImHZoom) then
        h := FParent.FSizes[ImageIndex].heigh * ImWZoom
      else
        w := FParent.FSizes[ImageIndex].Width * ImHZoom;
    end;
  end;
  DrawXObjectEx(x, YProjection(y),
    XProjection(FParent.FSizes[ImageIndex].Width),
    XProjection(FParent.FSizes[ImageIndex].heigh), XProjection(x),
    YProjection(y), w, h, CurrImName, angle, FDocScale);
end;

procedure THPDFPage.SetOrientation(const Value: THPDFPageOrientation);
begin
{$IFDEF BCB}
  if Value = vpoPortrait then
{$ELSE}
  if Value = poPortrait then
{$ENDIF}
    if FWidth > FHeight then
      TurnPage;
{$IFDEF BCB}
  if Value = vpoLandscape then
{$ELSE}
  if Value = poLandscape then
{$ENDIF}
    if FWidth < FHeight then
      TurnPage;
  FOrientation := Value;
end;

procedure THPDFPage.TurnPage;
var
  Tmp: Single;
begin
  Tmp := FHeight;
  FHeight := FWidth;
  FWidth := Tmp;
  Tmp := mHeight;
  mHeight := mWidth;
  mWidth := Tmp;
  SetPageWidth(mWidth);
  SetPageHeight(mHeight);
end;

procedure THPDFPage.CalculateHeightFormat;
begin
  FUpdateSize := true;
  mHeight := Height;
  DPI := 72 / Resolution; //72
{$IFDEF BCB}
  if FOrientation = vpoPortrait then
{$ELSE}
  if FOrientation = poPortrait then
{$ENDIF}
    FHeight := mHeight
  else
  begin
    FHeight := mWidth;
    FWidth := mHeight;
    mHeight := FHeight;
    mWidth := FWidth;
  end;
  SetPageHeight(mHeight);
  FHeight := trunc(mHeight / 72 * Resolution); //72
  PageMeta.Height := Round(FHeight);
  FUpdateSize := false;
end;

procedure THPDFPage.CalculateWidthFormat;
begin
  FUpdateSize := true;
  mWidth := Width;
  DPI := 72 / Resolution; //72
{$IFDEF BCB}
  if FOrientation = vpoPortrait then
{$ELSE}
  if FOrientation = poPortrait then
{$ENDIF}
    FWidth := mWidth
  else
  begin
    FHeight := mWidth;
    FWidth := mHeight;
    mHeight := FHeight;
    mWidth := FWidth;
  end;
  SetPageWidth(mWidth);
  FWidth := trunc(mWidth / 72 * Resolution); //72
  PageMeta.Width := Round(FWidth);
  FUpdateSize := false;
end;

procedure THPDFPage.Clip;
begin
  SaveToPageStream('W');
end;

procedure THPDFPage.ClosePath;
begin
  SaveToPageStream('h');
end;

procedure THPDFPage.ClosePathEoFillAndStroke;
begin
  SaveToPageStream('b*');
end;

procedure THPDFPage.ClosePathFillAndStroke;
begin
  SaveToPageStream('b');
end;

procedure THPDFPage.ClosePathStroke;
begin
  SaveToPageStream('s');
end;

procedure THPDFPage.EoClip;
begin
  SaveToPageStream('W*');
end;

procedure THPDFPage.EoFill;
begin
  SaveToPageStream('f*');
end;

procedure THPDFPage.EoFillAndStroke;
begin
  SaveToPageStream('B*');
end;

procedure THPDFPage.Fill;
begin
  SaveToPageStream('f');
end;

procedure THPDFPage.FillAndStroke;
begin
  SaveToPageStream('B');
end;

procedure THPDFPage.LineTo(X, Y: Single);
var
  S: AnsiString;
begin
  S := _FloatToStrR(XProjection(X)) + ' ' + _FloatToStrR(YProjection(Y)) + ' l';
  SaveToPageStream(S);
end;

procedure THPDFPage.MoveTo(X, Y: Single);
var
  S: AnsiString;
begin
  X := XProjection(X);
  Y := YProjection(Y);
  S := _FloatToStrR(X) + ' ' + _FloatToStrR(Y) + ' m';
  SaveToPageStream(S);
end;

procedure THPDFPage.NewPath;
begin
  SaveToPageStream('n');
end;

procedure THPDFPage.Stroke;
begin
  SaveToPageStream('S');
end;

procedure THPDFPage.SaveToPageStream(ValStr: AnsiString);
begin
  PageContent.Add(String(ValStr));
end;

function THPDFPage.XProjection(X: Single): Single;
begin
  if (FSize = psUserDefined) then
    result := X + FMinXVal
  else
  begin
    if (FParent.DocScale = 0) then
      FParent.DocScale := 1;
    result := (X / FParent.DocScale) * DPI + FMinXVal;
  end;
end;

function THPDFPage.YProjection(Y: Single): Single;
begin
  if (FSize = psUserDefined) then
    result := FMaxYVal - Y
  else
  begin
    if (FParent.DocScale = 0) then
      FParent.DocScale := 1;
    result := FMaxYVal - (Y / FParent.DocScale * DPI);
  end;
end;

procedure THPDFPage.SetRGBColor(Value: TColor);
begin
  SetRGBStrokeColor(Value);
  SetRGBFillColor(Value);
end;

procedure THPDFPage.RotateCoordinate(X, Y, Angle: Extended; var XR, YR: Extended);
var
  RCos, RSin: Single;
begin
  angle := angle * (Pi / 180);
  RCos := cos(angle);
  RSin := sin(angle);
  XR := ((RCos * X) - (RSin * Y));
  YR := ((RSin * X) + (RCos * Y));
end;

procedure THPDFPage.SetRGBFillColor(Value: TColor);
var
  S: AnsiString;
begin
  FFillColor := Value;
  S := _ColorToStrR(Value) + ' rg';
  SaveToPageStream(S);
  FKBegin := true;
end;

procedure THPDFPage.SetRGBStrokeColor(Value: TColor);
var
  S: AnsiString;
begin
  FStrokeColor := Value;
  S := _ColorToStrR(Value) + ' RG';
  SaveToPageStream(S);
  SKBegin := true;
end;

procedure THPDFPage.CurveToC(X1, Y1, X2, Y2, X3, Y3: Single);
var
  S: AnsiString;
begin
  X1 := XProjection(X1);
  Y1 := YProjection(Y1);
  X2 := XProjection(X2);
  Y2 := YProjection(Y2);
  X3 := XProjection(X3);
  Y3 := YProjection(Y3);
  S := _FloatToStrR(x1) + ' ' +
    _FloatToStrR(y1) + ' ' +
    _FloatToStrR(x2) + ' ' +
    _FloatToStrR(y2) + ' ' +
    _FloatToStrR(x3) + ' ' +
    _FloatToStrR(y3) + ' c';
  SaveToPageStream(S);
end;

procedure THPDFPage.CurveToV(X2, Y2, X3, Y3: Single);
var
  S: AnsiString;
begin
  X2 := XProjection(X2);
  Y2 := YProjection(Y2);
  X3 := XProjection(X3);
  Y3 := YProjection(Y3);
  S := _FloatToStrR(x2) + ' ' +
    _FloatToStrR(y2) + ' ' +
    _FloatToStrR(x3) + ' ' +
    _FloatToStrR(y3) + ' v';
  SaveToPageStream(S);
end;

procedure THPDFPage.CurveToY(X1, Y1, X3, Y3: Single);
var
  S: AnsiString;
begin
  X1 := XProjection(X1);
  Y1 := YProjection(Y1);
  X3 := XProjection(X3);
  Y3 := YProjection(Y3);
  S := _FloatToStrR(x1) + ' ' +
    _FloatToStrR(y1) + ' ' +
    _FloatToStrR(x3) + ' ' +
    _FloatToStrR(y3) + ' y';
  SaveToPageStream(S);
end;

function THPDFPage.DrawArc(X1, Y1, X2, Y2, X3, Y3, X4, Y4: Single): THPDFCurrPoint;
var
  YTmp, XTmp: Single;
  CenterX, CenterY: Extended;
  RadiusX, RadiusY: Extended;
  StartAngle, EndAngle, SweepRange: Extended;
  UseMoveTo: boolean;
begin
  Y1 := XProjection(Y1);
  Y2 := XProjection(Y2);
  YTmp := Y3;
  Y3 := XProjection(Y4);
  Y4 := XProjection(YTmp);
  X1 := XProjection(X1);
  X2 := XProjection(X2);
  XTmp := X3;
  X3 := XProjection(X4);
  X4 := XProjection(XTmp);
  CenterX := (X1 + X2) / 2;
  CenterY := (Y1 + Y2) / 2;
  RadiusX := (ABS(X1 - X2) - 1) / 2;
  RadiusY := (ABS(Y1 - Y2) - 1) / 2;
  if RadiusX < 0 then
    RadiusX := 0;
  if RadiusY < 0 then
    RadiusY := 0;
  StartAngle := ArcTan2(-(Y3 - CenterY) * RadiusX, (X3 - CenterX) * RadiusY);
  EndAngle := ArcTan2(-(Y4 - CenterY) * RadiusX, (X4 - CenterX) * RadiusY);
  result.X := CenterX + RadiusX * cos(StartAngle);
  result.Y := CenterY - RadiusY * sin(StartAngle);
  SweepRange := EndAngle - StartAngle;
  if SweepRange <= 0 then
    SweepRange := SweepRange + 2 * Pi;
  UseMoveTo := true;
  while SweepRange > Pi / 2 do
  begin
    CurveArc(CenterX, CenterY, RadiusX, RadiusY, StartAngle, Pi / 2, UseMoveTo);
    SweepRange := SweepRange - Pi / 2;
    StartAngle := StartAngle + Pi / 2;
    UseMoveTo := false;
  end;
  if SweepRange >= 0 then
    CurveArc(CenterX, CenterY, RadiusX, RadiusY, StartAngle, SweepRange,
      UseMoveTo);
end;

procedure THPDFPage.CurveArc(CenterX, CenterY, RadiusX, RadiusY, StartAngle, SweepRange: Extended; UseMoveTo: Boolean);
var
  Coord, C2: array[0..3] of THPDFCurrPoint;
  A, B, C, X, Y: Extended;
  ss, cc: Double;
  I: Integer;
begin
  if SweepRange = 0 then
  begin
    if UseMoveTo then
      RMoveTo(CenterX + RadiusX * cos(StartAngle),
        CenterY - RadiusY * sin(StartAngle));
    LineTo(CenterX + RadiusX * cos(StartAngle),
      CenterY - RadiusY * sin(StartAngle));
    Exit;
  end;
  B := sin(SweepRange / 2);
  C := cos(SweepRange / 2);
  A := 1 - C;
  X := A * 4 / 3;
  Y := B - X * C / B;
  ss := sin(StartAngle + SweepRange / 2);
  cc := cos(StartAngle + SweepRange / 2);
  Coord[0].X := C;
  Coord[0].Y := B;
  Coord[1].X := C + X;
  Coord[1].Y := Y;
  Coord[2].X := C + X;
  Coord[2].Y := -Y;
  Coord[3].X := C;
  Coord[3].Y := -B;
  for I := 0 to 3 do
  begin
    C2[I].X := CenterX + RadiusX * (Coord[I].X * cc + Coord[I].Y * ss) - 0.0001;
    C2[I].Y := CenterY + RadiusY * (-Coord[I].X * ss + Coord[I].Y * cc) - 0.0001;
  end;
  if UseMoveTo then
    RMoveTo(C2[0].X, Height - C2[0].Y);
  CurveToC(C2[1].X, C2[1].Y, C2[2].X, C2[2].Y, C2[3].X, C2[3].Y);
end;

procedure THPDFPage.DrawXObjectEx(X, Y, AWidth, AHeight: Single; ClipX, ClipY, ClipWidth, ClipHeight: Single; AXObjectName: AnsiString; angle, DocScale: Extended);
var
  Xsin, Xcos: Extended;
begin
  GStateSave;
  Concat(1, 0, 0, 1, XProjection(X), Y - XProjection(ClipHeight));
  if ABS(angle) > 0.001 then
  begin
    angle := angle * (Pi / 180);
    Xsin := sin(angle);
    Xcos := cos(angle);
    Concat(Xcos, Xsin, -Xsin, Xcos, 0, 0);
  end;
  Concat(XProjection(ClipWidth), 0, 0, XProjection(ClipHeight), 0, 0);
  ExecuteXObject(AXObjectName);
  GStateRestore;
end;

procedure THPDFPage.ExecuteXObject(XOName: AnsiString);
var
  S: AnsiString;
begin
  S := '/' + XOName + ' Do';
  SaveToPageStream(S);
end;

procedure THPDFPage.DrawPie(X1, Y1, X2, Y2, X3, Y3, X4, Y4: Extended);
var
  LineX, LineY: Extended;
  CurvePoint: THPDFCurrPoint;
begin
  CurvePoint := DrawArc(X1, Y1, X2, Y2, X3, Y3, X4, Y4);
  LineX := X1 + (X2 - X1) / 2;
  LineY := Y1 + (Y2 - Y1) / 2;
  LineTo(LineX, LineY);
  MoveTo(CurvePoint.X, CurvePoint.Y);
  LineTo(LineX, LineY);
end;

procedure THPDFPage.RMoveTo(X, Y: Single);
var
  S: AnsiString;
begin
  S := _FloatToStrR(X) + ' ' + _FloatToStrR(Y) + ' m';
  SaveToPageStream(S);
end;

function THPDFPage.GStateRestore: boolean;
begin
  if (ISStated > 0) then
  begin
    SaveToPageStream('Q');
    Dec(ISStated);
    result := true;
  end
  else
    result := false;
end;

procedure THPDFPage.GStateSave;
begin
  Inc(ISStated);
  SaveToPageStream('q');
end;

procedure THPDFPage.FEllipse(X, Y, Width, Height: Single);
begin
  RMoveTo(X * DPI, YProjection(Y + Height / 2));
  CurveToC(x, y + height / 2 - height / 2 * 11 / 20,
    x + width / 2 - width / 2 * 11 / 20, y, x + width / 2, y);
  CurveToC(x + width / 2 + width / 2 * 11 / 20,
    y, x + width, y + height / 2 - height / 2 * 11 / 20,
    x + width, y + height / 2);
  CurveToC(X + Width, Y + Height / 2 + Height / 2 * 11 / 20,
    x + width / 2 + width / 2 * 11 / 20,
    y + height, x + width / 2, y + height);
  CurveToC(x + width / 2 - width / 2 * 11 / 20,
    y + height, x, y + height / 2 + height / 2 * 11 / 20,
    x, y + height / 2);
end;

procedure THPDFPage.Concat(A, B, C, D, E, F: Single);
var
  S: AnsiString;
begin
  S := _FloatToStrR(a) + ' ' +
    _FloatToStrR(b) + ' ' +
    _FloatToStrR(c) + ' ' +
    _FloatToStrR(d) + ' ' +
    _FloatToStrR(e) + ' ' +
    _FloatToStrR(f) + ' cm';
  SaveToPageStream(S);
end;

{$IFDEF BCB}

procedure THPDFPage.UnicodeTextOut(X, Y: Single; Angle: Extended; Text: PWORD; TextLength: Integer);
{$ELSE}

procedure THPDFPage.UnicodeTextOut(X, Y: Single; Angle: Extended; Text: PWORD; TextLength: Integer);
{$ENDIF}
var
  I: Integer;
  ChCode: WORD;
  ChBuff: PWORD;
  DeltaH, DeltaW: Single;
  HorizontalLine: Single;
begin
  if CurrentFontObj.IsVertical then
  begin
    DeltaH := TextHeight('Zj');
    DeltaW := TextWidth('W');
    HorizontalLine := Y;
    ChBuff := @ChCode;
    for I := 0 to TextLength - 1 do
    begin
      ChCode := Text^;
      if (ChCode = $30FC) then
        ChCode := $7C;
      InternUnicodeTextOut(X + (DeltaW / 2), HorizontalLine - DeltaH, 0, ChBuff, 1);
      HorizontalLine := HorizontalLine + DeltaH;
      Inc(Text);
    end;
  end
  else
    InternUnicodeTextOut(X, Y, angle, Text, TextLength);
  StoreCurrentFont;
  CurrentFontObj.IsUsed := false;
end;

{$IFDEF BCB}

procedure THPDFPage.UnicodeTextOutStr(X, Y: Single; Angle: Extended; Text: WideString);
{$ELSE}

procedure THPDFPage.UnicodeTextOut(X, Y: Single; Angle: Extended; Text: WideString);
{$ENDIF}
var
  I: Integer;
  ChCode: WORD;
  ChBuff: PWORD;
  DeltaH, DeltaW: Single;
  HorizontalLine: Single;
  OutLen: Integer;
begin
  OutLen := Length(Text);
  if CurrentFontObj.IsVertical then
  begin
    DeltaH := TextHeight('Zj');
    DeltaW := TextWidth('W');
    HorizontalLine := Y;
    ChBuff := @ChCode;
    for I := 1 to OutLen do
    begin
      ChCode := Ord(Text[I]);
      if (ChCode = $30FC) then
        ChCode := $7C;
      InternUnicodeTextOut(X + (DeltaW / 2), HorizontalLine - DeltaH, 0, ChBuff, 1);
      HorizontalLine := HorizontalLine + DeltaH;
    end;
  end
  else
    InternUnicodeTextOut(X, Y, angle, @Text[1], OutLen);
  StoreCurrentFont;
  CurrentFontObj.IsUsed := false;
end;

procedure THPDFPage.InternUnicodeTextOut(X, Y: Single; Angle: Extended; Text: PWORD; TextLength: Integer);
var
  I: Integer;
  XBuff, YBuff: Single;
  MtxA, MtxB: Single;
  Abscissa: Single;
  StrHeight: Single;
  YProj: Single;
  ComposChar: AnsiString;
  TextString: AnsiString;
  OutText: WideString;
  OutCh: WORD;
  TableStream: TTrueTypeTables;
begin
  ProcessFont(true);
  XBuff := X;
  YBuff := Y;
  TableStream := TTrueTypeTables.Create;
  try
    if not (CurrentFontObj.FActive) then
    begin
      TableStream.CharacterDescription(@CurrentFontObj.Symbols,
        String(CurrentFontObj.OldName), CurrentFontObj.FontStyle);
    end;
    CurrentFontObj.FActive := true;
    OutText := '';
    TextString := '';
    for I := 0 to TextLength - 1 do
    begin
      TextString := TextString +
        AnsiString(chr(TableStream.ConvertFromUnicode(Text^)));
      OutCh := Text^;
      ComposChar := ProcChar(OutCh);
      OutText := OutText + WideString(ComposChar);
      Inc(Text);
    end;
  finally
    TableStream.Free;
  end;
  BeginText;
  StrHeight := TextHeight(TextString);
  if FTopTextPosition then
    YProj := 0
  else
    YProj := StrHeight;
  Abscissa := angle * Pi / 180;
  X := XProjection(X) + StrHeight * sin(Abscissa);
  Y := (YProjection(Y)) - StrHeight * cos(Abscissa);
  MtxA := cos(Abscissa);
  MtxB := sin(Abscissa);
  SetTextMatrix(MtxA, MtxB, -MtxB, MtxA, X, Y);
  ShowUnicodeText(OutText);
  EndText;
  if (fsUnderline in CurrentFontObj.FontStyle) then
  begin
    RectangleRotate(XBuff + 2 * sin((PI / 180) * Angle), YBuff + 2 * cos(Angle * (PI / 180)) + YProj { + StrHeight},
      TextWidth(TextString), CurrentFontObj.Size * 0.07, -1 * angle);
    Fill;
  end;
  if (fsStrikeOut in CurrentFontObj.FontStyle) then
  begin
    RectangleRotate(XBuff + 2 * sin((PI / 180) * Angle), YBuff + 2 * cos(Angle * (PI / 180)) + YProj + (StrHeight / 2),
      TextWidth(TextString), CurrentFontObj.Size * 0.02, -1 * angle);
    Fill;
  end;
end;

function THPDFPage.ProcChar(ComposChar: WORD): AnsiString;
var
  ResInt: Integer;
  NewChar: boolean;
  SmallLength: boolean;
  UnTableLength: Integer;
begin
  NewChar := true;
  SmallLength := false;
  UnTableLength := Length(CurrentFontObj.UnicodeTable);
  if UnTableLength > ComposChar then
  begin
    if CurrentFontObj.UnicodeTable[ComposChar].CharCode then
      NewChar := false;
  end
  else
    SmallLength := true;
  if NewChar then
  begin
    if SmallLength then
      SetLength(CurrentFontObj.UnicodeTable, ComposChar + 1);
    Inc(CurrentFontObj.UniLen);
    CurrentFontObj.UnicodeTable[ComposChar].CharCode := true;
    CurrentFontObj.UnicodeTable[ComposChar].Index := CurrentFontObj.Symbols[ComposChar].Index;
    ResInt := CurrentFontObj.Symbols[ComposChar].Index;
  end
  else
  begin
    ResInt := CurrentFontObj.UnicodeTable[ComposChar].Index;
  end;
  result := AnsiString(IntToHex(ResInt, 4));
end;

procedure THPDFPage.AddTextAnnotation(Contents: AnsiString; Rectangle: TRect; Open: boolean; Name: THPDFTextAnnotationType; Color: TColor = clRed);
var
  RGBCol: Integer;
  TmpAnnot: THPDFAnnotation;
begin
  TmpAnnot := THPDFAnnotation.Create;
  try
    TmpAnnot.FPage := Self;
    TmpAnnot.FType := asTextNotes;
    TmpAnnot.FOpen := Open;
    case Name of
      taComment: TmpAnnot.FName := 'Comment';
      taKey: TmpAnnot.FName := 'Key';
      taNote: TmpAnnot.FName := 'Note';
      taHelp: TmpAnnot.FName := 'Help';
      taNewParagraph: TmpAnnot.FName := 'NewParagraph';
      taParagraph: TmpAnnot.FName := 'Paragraph';
      taInsert: TmpAnnot.FName := 'Insert';
    end;
    if Color > 0 then
      RGBCol := Integer(Color)
    else
      RGBCol := 0;
    Move(RGBCol, TmpAnnot.FColor[0], 4);
    TmpAnnot.FLeftTop.X := Rectangle.Left;
    TmpAnnot.FLeftTop.Y := Rectangle.Top;
    TmpAnnot.FRightBottom.X := Rectangle.Right;
    TmpAnnot.FRightBottom.Y := Rectangle.Bottom;
    TmpAnnot.FContents := Contents;
    TmpAnnot.AddAnnotationObject;
  finally
    TmpAnnot.Free;
  end;
end;

procedure THPDFPage.AddFreeTextAnnotation(Contents: AnsiString; Rectangle: TRect; Quadding: THPDFFreeTextAnnotationJust; Color: TColor = clRed);
var
  RGBCol: Integer;
  TmpAnnot: THPDFAnnotation;
begin
  TmpAnnot := THPDFAnnotation.Create;
  try
    TmpAnnot.FPage := Self;
    TmpAnnot.FType := asFreeText;
    TmpAnnot.FContents := Contents;
    if Color > 0 then
      RGBCol := Integer(Color)
    else
      RGBCol := 0;
    Move(RGBCol, TmpAnnot.FColor[0], 4);
    TmpAnnot.FLeftTop.X := Rectangle.Left;
    TmpAnnot.FLeftTop.Y := Rectangle.Top;
    TmpAnnot.FRightBottom.X := Rectangle.Right;
    TmpAnnot.FRightBottom.Y := Rectangle.Bottom;
    TmpAnnot.FQuadding := Integer(Quadding);
    TmpAnnot.AddAnnotationObject;
  finally
    TmpAnnot.Free;
  end;
end;

procedure THPDFPage.AddLineAnnotation(Contents: AnsiString; BeginPoint: THPDFCurrPoint; EndPoint: THPDFCurrPoint; Color: TColor = clRed);
var
  RGBCol: Integer;
  TmpAnnot: THPDFAnnotation;
begin
  TmpAnnot := THPDFAnnotation.Create;
  try
    TmpAnnot.FPage := Self;
    TmpAnnot.FType := asLine;
    TmpAnnot.FLeftTop := BeginPoint;
    TmpAnnot.FRightBottom := EndPoint;
    TmpAnnot.FContents := Contents;
    if Color > 0 then
      RGBCol := Integer(Color)
    else
      RGBCol := 0;
    Move(RGBCol, TmpAnnot.FColor[0], 4);
    TmpAnnot.AddAnnotationObject;
  finally
    TmpAnnot.Free;
  end;
end;

procedure THPDFPage.AddCircleSquareAnnotation(Contents: AnsiString; Rectangle: TRect; CSType: THPDFCSAnnotationType; Color: TColor = clRed);
var
  RGBCol: Integer;
  TmpAnnot: THPDFAnnotation;
begin
  TmpAnnot := THPDFAnnotation.Create;
  try
    TmpAnnot.FPage := Self;
    if (CSType = csCircle) then
      TmpAnnot.FType := asCircle
    else
      if (CSType = csSquare) then
      TmpAnnot.FType := asSquare;
    if Color > 0 then
      RGBCol := Integer(Color)
    else
      RGBCol := 0;
    Move(RGBCol, TmpAnnot.FColor[0], 4);
    TmpAnnot.FLeftTop.X := Rectangle.Left;
    TmpAnnot.FLeftTop.Y := Rectangle.Top;
    TmpAnnot.FRightBottom.X := Rectangle.Right;
    TmpAnnot.FRightBottom.Y := Rectangle.Bottom;
    TmpAnnot.FContents := Contents;
    TmpAnnot.AddAnnotationObject;
  finally
    TmpAnnot.Free;
  end;
end;

procedure THPDFPage.AddStampAnnotation(Contents: AnsiString; Rectangle: TRect; StampType: THPDFStampAnnotationType; Color: TColor = clRed);
var
  RGBCol: Integer;
  TmpAnnot: THPDFAnnotation;
begin
  TmpAnnot := THPDFAnnotation.Create;
  try
    TmpAnnot.FType := asStamp;
    TmpAnnot.FPage := Self;
    case StampType of
      satApproved:
        TmpAnnot.FTypeState := 'Approved';
      satExperimental:
        TmpAnnot.FTypeState := 'Experimental';
      satNotApproved:
        TmpAnnot.FTypeState := 'NotApproved';
      satAsIs:
        TmpAnnot.FTypeState := 'AsIs';
      satExpired:
        TmpAnnot.FTypeState := 'Expired';
      satNotForPublicRelease:
        TmpAnnot.FTypeState := 'NotForPublicRelease';
      satConfidential:
        TmpAnnot.FTypeState := 'Confidential';
      satFinal:
        TmpAnnot.FTypeState := 'Final';
      satSold:
        TmpAnnot.FTypeState := 'Sold';
      satDepartmental:
        TmpAnnot.FTypeState := 'Departmental';
      satForComment:
        TmpAnnot.FTypeState := 'ForComment';
      satTopSecret:
        TmpAnnot.FTypeState := 'TopSecret';
      satDraft:
        TmpAnnot.FTypeState := 'Draft';
      satForPublicRelease:
        TmpAnnot.FTypeState := 'ForPublicRelease';
    end;
    if Color > 0 then
      RGBCol := Integer(Color)
    else
      RGBCol := 0;
    Move(RGBCol, TmpAnnot.FColor[0], 4);
    TmpAnnot.FContents := Contents;
    TmpAnnot.FLeftTop.X := Rectangle.Left;
    TmpAnnot.FLeftTop.Y := Rectangle.Top;
    TmpAnnot.FRightBottom.X := Rectangle.Right;
    TmpAnnot.FRightBottom.Y := Rectangle.Bottom;
    TmpAnnot.AddAnnotationObject;
  finally
    TmpAnnot.Free;
  end;
end;

procedure THPDFPage.AddFileAttachmentAnnotation(Contents: AnsiString; FileName: AnsiString; Rectangle: TRect; Color: TColor = clRed);
var
  RGBCol: Integer;
  TmpAnnot: THPDFAnnotation;
begin
  TmpAnnot := THPDFAnnotation.Create;
  try
    TmpAnnot.FType := asFileAttachment;
    TmpAnnot.FPage := Self;
    if Color > 0 then
      RGBCol := Integer(Color)
    else
      RGBCol := 0;
    Move(RGBCol, TmpAnnot.FColor[0], 4);
    TmpAnnot.FName := FileName;
    TmpAnnot.FContents := Contents;
    TmpAnnot.FLeftTop.X := Rectangle.Left;
    TmpAnnot.FLeftTop.Y := Rectangle.Top;
    TmpAnnot.FRightBottom.X := Rectangle.Right;
    TmpAnnot.FRightBottom.Y := Rectangle.Bottom;
    TmpAnnot.AddAnnotationObject;
  finally
    TmpAnnot.Free;
  end;
end;

procedure THPDFPage.AddSoundAnnotation(Contents: AnsiString; FileName: AnsiString; Rectangle: TRect; Color: TColor = clRed);
var
  RGBCol: Integer;
  TmpAnnot: THPDFAnnotation;
begin
  TmpAnnot := THPDFAnnotation.Create;
  try
    TmpAnnot.FType := asSound;
    TmpAnnot.FPage := Self;
    if Color > 0 then
      RGBCol := Integer(Color)
    else
      RGBCol := 0;
    Move(RGBCol, TmpAnnot.FColor[0], 4);
    TmpAnnot.FName := FileName;
    TmpAnnot.FContents := Contents;
    TmpAnnot.FLeftTop.X := Rectangle.Left;
    TmpAnnot.FLeftTop.Y := Rectangle.Top;
    TmpAnnot.FRightBottom.X := Rectangle.Right;
    TmpAnnot.FRightBottom.Y := Rectangle.Bottom;
    TmpAnnot.AddAnnotationObject;
  finally
    TmpAnnot.Free;
  end;
end;

procedure THPDFPage.AddMovieAnnotation(Contents: AnsiString; FileName: AnsiString; Rectangle: TRect; Color: TColor = clRed);
var
  RGBCol: Integer;
  TmpAnnot: THPDFAnnotation;
begin
  TmpAnnot := THPDFAnnotation.Create;
  try
    TmpAnnot.FType := asMovie;
    TmpAnnot.FPage := Self;
    if Color > 0 then
      RGBCol := Integer(Color)
    else
      RGBCol := 0;
    Move(RGBCol, TmpAnnot.FColor[0], 4);
    TmpAnnot.FName := FileName;
    TmpAnnot.FContents := Contents;
    TmpAnnot.FLeftTop.X := Rectangle.Left;
    TmpAnnot.FLeftTop.Y := Rectangle.Top;
    TmpAnnot.FRightBottom.X := Rectangle.Right;
    TmpAnnot.FRightBottom.Y := Rectangle.Bottom;
    TmpAnnot.AddAnnotationObject;
  finally
    TmpAnnot.Free;
  end;
end;

procedure THPDFPage.NoDash;
var
  DashArray: array[0..1] of Byte;
  Phase: byte;
begin
  DashArray[0] := 0;
  Phase := 0;
  SetDash(DashArray, Phase);
end;

procedure THPDFPage.SetDash(DashArray: array of byte; Phase: byte);
var
  S: AnsiString;
  I: Integer;
begin
  S := '[';
  if (High(DashArray) >= 0) and (DashArray[0] <> 0) then
    for I := 0 to High(DashArray) do
      S := S + AnsiString(IntToStr(DashArray[I])) + ' ';
  S := S + '] ' + AnsiString(IntToStr(Phase)) + ' d';
  SaveToPageStream(S);
end;

procedure THPDFPage.SetFlat(Flatness: byte);
var
  S: AnsiString;
begin
  S := AnsiString(IntToStr(Flatness)) + ' i';
  SaveToPageStream(S);
end;

procedure THPDFPage.SetLineCap(Linecap: TLineCapStyle);
var
  S: AnsiString;
begin
  S := AnsiString(IntToStr(Ord(Linecap))) + ' J';
  SaveToPageStream(S);
end;

procedure THPDFPage.SetLineJoin(LineJoin: TLineJoinStyle);
var
  S: AnsiString;
begin
  S := AnsiString(IntToStr(Ord(LineJoin))) + ' j';
  SaveToPageStream(S);
end;

procedure THPDFPage.SetLineWidth(Width: Single);
var
  S: AnsiString;
begin
  Width := XProjection(Width);
  S := _FloatToStrR(Width) + ' w';
  SaveToPageStream(S);
end;

procedure THPDFPage.SetMiterLimit(MiterLimit: byte);
var
  S: AnsiString;
begin
  S := AnsiString(IntToStr(MiterLimit)) + ' M';
  SaveToPageStream(S);
end;

procedure THPDFPage.MFRectangle(X, Y, X1, Y1: Single);
begin
  Rectangle(X, Y, X1 - X, Y1 - Y);
end;

procedure THPDFPage.Rectangle(X, Y, Width, Height: Single);
var
  S: AnsiString;
begin
  X := XProjection(X);
  Y := YProjection(Y);
  Width := XProjection(Width);
  Height := XProjection(Height);
  S := _FloatToStrR(x) + ' ' +
    _FloatToStrR(y - Height) + ' ' +
    _FloatToStrR(Width) + ' ' +
    _FloatToStrR(Height) + ' re';
  SaveToPageStream(S);
end;

procedure THPDFPage.RectangleRotate(X, Y, Width, Height: Single;
  angle: Extended);
var
  XR, YR: Extended;
begin
  MoveTo(X, Y);
  RotateCoordinate(Width, 0, angle, XR, YR);
  LineTo(X + XR, Y + YR);
  RotateCoordinate(Width, Height, angle, XR, YR);
  LineTo(X + XR, Y + YR);
  RotateCoordinate(0, Height, angle, XR, YR);
  LineTo(X + XR, Y + YR);
end;

procedure THPDFPage.SetCharacterSpacing(Spacing: Single);
begin
  if FCharSpace = Spacing then
    Exit;
  FCharSpace := Spacing;
  SaveToPageStream(_FloatToStrR(Spacing) + ' Tc');
end;

procedure THPDFPage.SetHorizontalScaling(Scaling: Single);
begin
  if FHorizontalScaling = Scaling then
    Exit;
  FHorizontalScaling := Scaling;
  SaveToPageStream(_FloatToStrR(Scaling) + ' Tz');
end;

procedure THPDFPage.SetLeading(Leading: Single);
begin
  if FLeading = Leading then
    Exit;
  FLeading := Leading;
  SaveToPageStream(_FloatToStrR(Leading) + ' TL');
end;

procedure THPDFPage.SetWordSpacing(Spacing: Single);
begin
  if FWordSpace = Spacing then
    Exit;
  FWordSpace := Spacing;
  SaveToPageStream(_FloatToStrR(Spacing) + ' Tw');
end;

procedure THPDFPage.MoveTextPoint(X, Y: Single);
var
  S: AnsiString;
begin
  X := XProjection(X);
  Y := YProjection(Y);
  S := _FloatToStrR(X) + ' ' + _FloatToStrR(Y) + ' Td';
  SaveToPageStream(S);
end;

procedure THPDFPage.SetTextRenderingMode(Mode: TPDFTextRenderingMode);
begin
  SaveToPageStream(AnsiString(IntToStr(Ord(Mode))) + ' Tr');
end;

procedure THPDFPage.SetTextRise(Rise: SmallInt);
begin
  SaveToPageStream(AnsiString(IntToStr(Rise)) + ' Ts');
end;

procedure THPDFPage.MoveToNextLine;
begin
  SaveToPageStream('T*');
end;

procedure THPDFPage.SetRGBHyperlinkColor(Value: TColor);
begin
  FHyperColor := Value;
end;

procedure THPDFPage.SetTextMatrix(A, B, C, D, X, Y: Single);
var
  S: AnsiString;
begin
  S := _CutFloat(a) + ' ' +
    _CutFloat(b) + ' ' +
    _CutFloat(c) + ' ' +
    _CutFloat(d) + ' ' +
    _CutFloat(x) + ' ' +
    _CutFloat(y) + ' Tm';
  SaveToPageStream(S);
end;

procedure THPDFPage.ShowUnicodeText(Text: WideString);
var
  FString: AnsiString;
begin
  FString := '<' + AnsiString(Text) + '>';
  SaveToPageStream(FString + ' Tj');
end;

procedure THPDFPage.ShowText(Text: AnsiString; IsHexadecimal: boolean);
var
  FString: AnsiString;
begin
  if IsHexadecimal then
    FString := '<' + _StrToHex(Text) + '>'
  else
    FString := '(' + _EscapeText(Text) + ')';
  SaveToPageStream(FString + ' Tj');
end;

procedure THPDFPage.Circle(X, Y, Radius: Single);
begin
  Ellipse(X - Radius * DPI, Y - Radius * DPI, Radius * 2, Radius * 2);
end;

procedure THPDFPage.FMEllipse(X, Y, X1, Y1: Single);
begin
  Ellipse(X, Y, X1 - X, Y1 - Y);
end;

procedure THPDFPage.Ellipse(X, Y, Width, Height: Single);
begin
  Width := Width / DocScale;
  Height := Height / DocScale;
  FEllipse(X, Y, Width, Height);
end;

procedure THPDFPage.RoundRect(X1, Y1, X2, Y2, X3, Y3: Integer);
const
  B = 0.5522847498;
var
  RX, RY: Extended;
begin
  RX := X3 / 2;
  RY := Y3 / 2;
  MoveTo(X1 + RX, Y1);
  LineTo(X2 - RX, Y1);
  CurveToC(X2 - RX + B * RX, Y1, X2, (Y1 + RY - B * RY), X2, (Y1 + RY));
  LineTo(X2, Y2 - RY);
  CurveToC(X2, (Y2 - RY + B * RY), X2 - RX + B * RX, Y2, X2 - RX, Y2);
  LineTo(X1 + RX, Y2);
  CurveToC(X1 + RX - B * RX, Y2, X1, (Y2 - RY + B * RY), X1, (Y2 - RY));
  LineTo(X1, Y1 + RY);
  CurveToC(X1, (Y1 + RY - B * RY), X1 + RX - B * RX, Y1, X1 + RX, Y1);
  ClosePath;
end;

procedure THPDFPage.SetGrayColor(GrayColor: Extended);
begin
  SetGrayStrokeColor(GrayColor);
  SetGrayFillColor(GrayColor);
end;

procedure THPDFPage.SetGrayFillColor(GrayColor: Extended);
var
  S: AnsiString;
begin
  FFillColor := Round($FF * GrayColor * $FF * $100) + Round($FF * GrayColor * $FF) + Round($FF * GrayColor);
  S := _FloatToStrR(GrayColor) + ' g';
  SaveToPageStream(S);
end;

procedure THPDFPage.SetGrayStrokeColor(GrayColor: Extended);
var
  S: AnsiString;
begin
  FStrokeColor := Round($FF * GrayColor * $FF * $100) + Round($FF * GrayColor * $FF) + Round($FF * GrayColor);
  S := _FloatToStrR(GrayColor) + ' G';
  SaveToPageStream(S);
end;

procedure THPDFPage.DrawBarcode(BCType: THPDFBarcodeType; X, Y, Height,
  MUnit: Integer; Angle: Single; Info: AnsiString; UseCheckSum: Boolean; BarColor, Color: TColor);
var
  BarCode: THPDFBarcode;
begin
  BarCode := THPDFBarcode.Create(nil);
  try
    BarCode.Typ := Integer(BCType);
    BarCode.Top := Y;
    BarCode.Left := X;
    BarCode.Height := Height;
    BarCode.Modul := MUnit;
    BarCode.angle := angle;
    BarCode.Text := Info;
    BarCode.Checksum := UseCheckSum;
    BarCode.Color := Color;
    BarCode.ColorBar := BarColor;
    BarCode.DrawBarcode(FCanvas);
  finally
    BarCode.Free;
  end;
end;

procedure THPDFPage.PrintHyperlink(X, Y: Single; Text, Link: AnsiString);
var
  OldFillColor, OldStrokeColor: TColor;
begin
  OldFillColor := FFillColor;
  OldStrokeColor := FStrokeColor;
  SetRGBColor(FHyperColor);
{$IFDEF BCB}
  OutText(X, Y, 0, Text);
{$ELSE}
  TextOut(X, Y, 0, Text);
{$ENDIF}
  SetUriObject(Link, XProjection(X), YProjection(Y), XProjection(X + TextWidth(Text)), YProjection(Y + TextHeight(Text)));
  SetRGBFillColor(OldFillColor);
  SetRGBStrokeColor(OldStrokeColor);
end;

procedure THPDFPage.SetURIObject(UriLink: AnsiString; Left, Top, Right, Bottom: Extended);
var
  FURIObj: THPDFDictionaryObject;
  FAnnotation: THPDFDictionaryObject;
  FArray: THPDFArrayObject;
  FBorder: THPDFArrayObject;
  AnnotsObject: THPDFArrayObject;
  AnnotsIndex: Integer;
  AnnotObj: THPDFObject;
  AnnotTmpLink: THPDFDictionaryObject;
begin
  Inc(FParent.FMaxObjNum);
  FURIObj := THPDFDictionaryObject.Create(nil);
  FURIObj.IsIndirect := true;
  FURIObj.ID.ObjectNumber := FParent.FMaxObjNum;
  FParent.IndirectObjects.Add(FURIObj);
  FURIObj.AddNameValue('S', 'URI');
  FURIObj.AddStringValue('URI', UriLink);
  Inc(FParent.FMaxObjNum);
  FAnnotation := THPDFDictionaryObject.Create(nil);
  FAnnotation.IsIndirect := true;
  FAnnotation.ID.ObjectNumber := FParent.FMaxObjNum;
  FAnnotation.AddNameValue('Type', 'Annot');
  FAnnotation.AddNameValue('Subtype', 'Link');
  FParent.IndirectObjects.Add(FAnnotation);
  FArray := THPDFArrayObject.Create(nil);
  FBorder := THPDFArrayObject.Create(nil);
  FAnnotation.AddValue('Border', THPDFObject(FBorder));
  FArray.AddNumericValue(Left);
  FArray.AddNumericValue(Top);
  FArray.AddNumericValue(Right);
  FArray.AddNumericValue(Bottom);
  FAnnotation.AddValue('Rect', THPDFObject(FArray));
  FAnnotation.AddValue('A', THPDFObject(FURIObj));
  AnnotsIndex := PageObj.FindValue('Annots');
  if (AnnotsIndex >= 0) then
  begin
    AnnotObj := PageObj.GetIndexedItem(AnnotsIndex);
    if (AnnotObj.ObjectType = otLink) then
    begin
      AnnotTmpLink := THPDFDictionaryObject(FParent.GetObjectByLink(THPDFLink(AnnotObj)));
      AnnotObj := THPDFObject(AnnotTmpLink);
    end;
    AnnotsObject := THPDFArrayObject(AnnotObj);
  end
  else
  begin
    AnnotsObject := THPDFArrayObject.Create(nil);
    PageObj.AddValue('Annots', AnnotsObject);
  end;
  AnnotsObject.AddObject(FAnnotation);
end;

{ THPDFDocOutlineObjectObject }

function THPDFDocOutlineObject.AddChild(Title: AnsiString; X, Y: Single): THPDFDocOutlineObject;
var
  Ind: Integer;
  DestArray: THPDFArrayObject;
  TmpEntry: THPDFDocOutlineObject;
begin
  result := THPDFDocOutlineObject.Create;
  TmpEntry := Self;
  while TmpEntry <> nil do
  begin
    TmpEntry.FCount := TmpEntry.FCount + 1;
    Ind := TmpEntry.LinkedObj.FindValue('Count');
    if (Ind >= 0) then
    begin
      if ((TmpEntry.Opened) or (TmpEntry.FCount = 0)) then
        THPDFNumericObject(TmpEntry.LinkedObj.GetIndexedItem(Ind)).Value := TmpEntry.FCount
      else
        THPDFNumericObject(TmpEntry.LinkedObj.GetIndexedItem(Ind)).Value := (TmpEntry.FCount * (-1));
    end;
    TmpEntry := TmpEntry.Parent;
  end;
  Inc(FDoc.OutlineEnsLen);
  SetLength(FDoc.OutlineEnsemble, FDoc.OutlineEnsLen);
  FDoc.OutlineEnsemble[FDoc.OutlineEnsLen - 1] := result;
  Inc(FDoc.FMaxObjNum);
  result.FParent := Self;
  result.LinkedObj := THPDFDictionaryObject.Create(nil);
  result.LinkedObj.IsIndirect := true;
  result.LinkedObj.ID.ObjectNumber := FDoc.FMaxObjNum;
  FDoc.IndirectObjects.Add(result.LinkedObj);
  if (result.Opened) then
    result.LinkedObj.AddNumericValue('Count', result.FCount)
  else
    result.LinkedObj.AddNumericValue('Count', (result.FCount * (-1)));
  result.LinkedObj.AddStringValue('Title', Title);
  if FFirst = nil then
  begin
    FFirst := result;
    LinkedObj.AddValue('First', result.LinkedObj);
  end;
  if FLast <> nil then
  begin
    FLast.FNext := result;
    FLast.LinkedObj.AddValue('Next', result.LinkedObj);
    result.LinkedObj.AddValue('Prev', FLast.LinkedObj);
  end;
  result.FDoc := FDoc;
  result.FPrev := FLast;
  LinkedObj.AddValue('Last', result.LinkedObj);
  FLast := result;
  result.Title := Title;
  result.FLeft := trunc(result.FDoc.FCurrentPage.XProjection(X));
  result.FTop := trunc(result.FDoc.FCurrentPage.YProjection(Y));
  DestArray := THPDFArrayObject.Create(nil);
  DestArray.AddObject(FDoc.FCurrentPage.PageObj);
  DestArray.AddNameValue('XYZ');
  DestArray.AddNumericValue(result.FLeft);
  DestArray.AddNumericValue(result.FTop);
  DestArray.AddNumericValue(0);
  result.LinkedObj.AddValue('Dest', DestArray);
end;

procedure THPDFDocOutlineObject.Init(AOwner: THotPDF);
begin
  FCount := 0;
  FDoc := AOwner;
  FParent := nil;
  FNext := nil;
  FPrev := nil;
  FFirst := nil;
  FLast := nil;
  FFirst := nil;
  FLast := nil;
  LinkedObj := THPDFDictionaryObject.Create(nil);
  LinkedObj.AddNumericValue('Count', 0);
end;

{ THPDFParagraph }

constructor THPDFPara.Create(Parent: THotPDF);
begin
  FParent := Parent;
end;


procedure THPDFPara.InternShowUnicodeText(Text: WideString; var APrText: WideString; var APrVax: Single);
var
  I: Integer;
  Xi, h: Single;
  Wides: Single;
  SDadicate: Single;
  NisHeight: Single;
  SeLen: Integer;
  TextString: AnsiString;
  PrNewString: AnsiString;
  PreFontName: AnsiString;
  PreFontStyle: TFontStyles;
  PreASize: Single;
  PreIsVertical: boolean;
  PreFontCharset: TFontCharset;
  TableStream: TTrueTypeTables;
begin
  if (FParent.FActivePara > 0) then
  begin
    FParent.CurrentPage.ProcessFont(true);
    TableStream := TTrueTypeTables.Create;
    try
      if not (FParent.CurrentPage.CurrentFontObj.FActive) then
      begin
        TableStream.CharacterDescription(@FParent.CurrentPage.CurrentFontObj.Symbols,
          String(FParent.CurrentPage.CurrentFontObj.OldName),
          FParent.CurrentPage.CurrentFontObj.FontStyle);
      end;
      FParent.CurrentPage.CurrentFontObj.FActive := true;
      TextString := '';
      Xi := 0;
      PrNewString := '';
      SeLen := Length(Text);
      h := 0;
      NisHeight := FParent.CurrentPage.TextHeight('W');
      while (SeLen > 0) do
      begin
        I := 0;
        if (FParent.FNewParaLine) then
          SDadicate := (Fparent.FCurrentPage.FWidth - LeftMargin - RightMargin - Indention)
        else
          if (FParent.FPrevParaOffset = 0) then
          SDadicate := (FParent.FCurrentPage.FWidth - LeftMargin - RightMargin)
        else
            SDadicate := (Fparent.FCurrentPage.FWidth - FParent.FPrevParaOffset - RightMargin);
        PrNewString := '';
        TextString := '';
        Wides := 0;
        while ((Wides <= SDadicate) and (I < SeLen)) do
        begin
          Inc(I);
          PrNewString := PrNewString + AnsiChar(Text[I]);
          TextString := TextString + AnsiChar(chr(TableStream.ConvertFromUnicode(WOrd(Text[I]))));
          Wides := FParent.FCurrentPage.TextWidth(TextString);
        end;
        Text := Copy(Text, I + 1, SeLen - I);
        SeLen := Length(Text);
        if (FJustification = jtCenter) then
        begin
          Xi := (Fparent.FCurrentPage.FWidth / 2) + LeftMargin / 2 - (Wides / 2);
        end
        else
        begin
          if (FParent.FNewParaLine) then
            Xi := Indention + LeftMargin
          else
            if (FParent.FPrevParaOffset = 0) then
            Xi := LeftMargin
          else
          begin
            Xi := FParent.FPrevParaOffset;
            FParent.FPrevParaOffset := 0;
          end;
        end;
{$IFDEF BCB}
        Fparent.FCurrentPage.UnicodeTextOutStr(Xi, Fparent.FCurrentParaLine + H, 0, WideString(PrNewString));
{$ELSE}
        Fparent.FCurrentPage.UnicodeTextOut(Xi, Fparent.FCurrentParaLine + H, 0, WideString(PrNewString));
{$ENDIF}
        h := h + NisHeight;
        FParent.FNewParaLine := false;
        if (SeLen > 0) then
        begin
          if ((Fparent.FCurrentParaLine + H + NisHeight) > (Fparent.FCurrentPage.FHeight - BottomMargin)) then
          begin
            PreFontName := FParent.FCurrentPage.CurrentFontObj.Name;
            PreFontStyle := FParent.FCurrentPage.CurrentFontObj.FontStyle;
            PreASize := FParent.FCurrentPage.CurrentFontObj.Size;
            PreFontCharset := FParent.FCurrentPage.CurrentFontObj.FontCharset;
            PreIsVertical := FParent.FCurrentPage.CurrentFontObj.IsVertical;
            FParent.AddPage;
            FParent.FCurrentParaLine := TopMargin;
            FParent.FPrevParaOffset := 0;
            FParent.CurrentPage.SetFont(PreFontName, PreFontStyle, PreASize, PreFontCharset, PreIsVertical);
            InternShowUnicodeText(Text, APrText, APrVax);
            PrNewString := AnsiString(APrText);
            Xi := APrVax;
            break;
          end;
        end;
      end;
      FParent.FCurrentParaLine := FParent.FCurrentParaLine + h;
      APrText := WideString(PrNewString);
      APrVax := Xi;
      FParent.FPrevParaOffset := 0;
    finally
      TableStream.Free;
    end;
  end
  else
  begin
    raise exception.Create('Current paragraph is nil');
  end;
end;

procedure THPDFPara.ShowUnicodeText(Text: WideString);
var
  XiLine: Single;
  SmTextString: WideString;
begin
  InternShowUnicodeText(Text, SmTextString, XiLine);
end;

procedure THPDFPara.SetJustification(const Value: THPDFJustificationType);
begin
  FJustification := Value;
  FParent.FNewParaLine := true;
end;

procedure THPDFPara.InternShowText(Text: AnsiString; var APrText: AnsiString; var APrVax: Single);
var
  I: Integer;
  Xi, h: Single;
  SDadicate: Single;
  NisHeight: Single;
  SeLen: Integer;
  PrNewString: AnsiString;
  PreFontName: AnsiString;
  PreFontStyle: TFontStyles;
  PreASize: Single;
  PreFontCharset: TFontCharset;
  PreIsVertical: boolean;
begin
  if (FParent.FActivePara > 0) then
  begin
    Xi := 0;
    PrNewString := '';
    SeLen := Length(Text);
    h := 0;
    NisHeight := FParent.CurrentPage.TextHeight('W');
    while (SeLen > 0) do
    begin
      if (FParent.FNewParaLine) then
        SDadicate := (Fparent.FCurrentPage.FWidth - LeftMargin - RightMargin - Indention)
      else
        if (FParent.FPrevParaOffset = 0) then
        SDadicate := (FParent.FCurrentPage.FWidth - LeftMargin - RightMargin)
      else
          SDadicate := (Fparent.FCurrentPage.FWidth - FParent.FPrevParaOffset - RightMargin);
      PrNewString := '';
      I := 0;
      while ((Fparent.FCurrentPage.TextWidth(PrNewString) <= SDadicate) and (I < SeLen)) do
      begin
        Inc(I);
        PrNewString := PrNewString + Text[I];
      end;
      if (I < SeLen) then
      begin
        Text := Copy(Text, I + 1, SeLen - I);
        SeLen := Length(Text);
      end
      else
      begin
        Text := '';
        SeLen := 0;
      end;
      if (FJustification = jtCenter) then
      begin
        Xi := (Fparent.FCurrentPage.FWidth / 2) + LeftMargin / 2 - (Fparent.FCurrentPage.TextWidth(PrNewString) / 2);
        if (H = 0) then NewLine;
      end
      else
      begin
        if (FJustification = jtLeft) then
        begin
          if (FParent.FNewParaLine) then
            Xi := Indention + LeftMargin
          else
            if (FParent.FPrevParaOffset = 0) then
            Xi := LeftMargin
          else
          begin
            Xi := FParent.FPrevParaOffset;
            FParent.FPrevParaOffset := 0;
          end;
        end
        else
        begin
          Xi := Fparent.FCurrentPage.FWidth - RightMargin - Fparent.FCurrentPage.TextWidth(PrNewString);
          if (FParent.FPrevParaOffset <> 0) then
            NewLine;
        end
      end;
{$IFDEF BCB}
      Fparent.FCurrentPage.PrintText(Xi, Fparent.FCurrentParaLine + H, 0, PrNewString);
{$ELSE}
      Fparent.FCurrentPage.TextOut(Xi, Fparent.FCurrentParaLine + H, 0, PrNewString);
{$ENDIF}
      h := h + NisHeight;
      FParent.FNewParaLine := false;
      if (SeLen > 0) then
      begin
        if ((Fparent.FCurrentParaLine + H + NisHeight) > (Fparent.FCurrentPage.FHeight - BottomMargin)) then
        begin
          PreFontName := FParent.FCurrentPage.CurrentFontObj.Name;
          PreFontStyle := FParent.FCurrentPage.CurrentFontObj.FontStyle;
          PreASize := FParent.FCurrentPage.CurrentFontObj.Size;
          PreFontCharset := FParent.FCurrentPage.CurrentFontObj.FontCharset;
          PreIsVertical := FParent.FCurrentPage.CurrentFontObj.IsVertical;
          FParent.AddPage;
          FParent.FCurrentParaLine := TopMargin;
          FParent.FPrevParaOffset := 0;
          FParent.CurrentPage.SetFont(PreFontName, PreFontStyle, PreASize, PreFontCharset, PreIsVertical);
          InternShowText(Text, APrText, APrVax);
          PrNewString := APrText;
          Xi := APrVax;
          break;
        end;
      end;
    end;
    if (FJustification = jtLeft) then
      FParent.FCurrentParaLine := FParent.FCurrentParaLine + h - NisHeight
    else
      FParent.FCurrentParaLine := FParent.FCurrentParaLine + h;
    APrVax := Xi;
    APrText := PrNewString;
    if (FJustification = jtLeft) then
      FParent.FPrevParaOffset := Xi + Fparent.FCurrentPage.TextWidth(PrNewString)
    else
      FParent.FPrevParaOffset := 0;
  end
  else
  begin
    raise exception.Create('Current paragraph is nil');
  end;
end;

{$IFDEF BCB}

procedure THPDFPara.PrintText(Text: AnsiString);
{$ELSE}

procedure THPDFPara.ShowText(Text: AnsiString);
{$ENDIF}
var
  XiLine: Single;
  SmTextString: AnsiString;
begin
  InternShowText(Text, SmTextString, XiLine);
end;

procedure THPDFPara.NewLine;
begin
  FParent.FNewParaLine := true;
  FParent.FPrevParaOffset := 0;
  FParent.FCurrentParaLine := FParent.FCurrentParaLine + FParent.FCurrentPage.TextHeight('W');
end;

function THPDFPara.GetCurrentLine: Single;
begin
  result := FParent.FCurrentParaLine;
end;

procedure THPDFPara.SetCurrentLine(const Value: Single);
begin
  FParent.FCurrentParaLine := Value;
end;

procedure Register;
begin
  RegisterComponents('losLab', [THotPDF]);
end;

end.
