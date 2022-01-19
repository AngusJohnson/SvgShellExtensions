unit SvgPreview;

(*******************************************************************************
* Author    :  Angus Johnson                                                   *
* Version   :  1.0                                                            *
* Date      :  18 January 2022                                                 *
* Website   :  http://www.angusj.com                                           *
* Copyright :  Angus Johnson 2022                                              *
*                                                                              *
* Purpose   :  IPreviewHandler and IThumbnailProvider for SVG image files      *
*                                                                              *
* License   :  Use, modification & distribution is subject to                  *
*              Boost Software License Ver 1                                    *
*              http://www.boost.org/LICENSE_1_0.txt                            *
*******************************************************************************)

interface

uses
  Windows, Messages, ActiveX, Classes, ComObj, ComServ, ShlObj,
  PropSys, Types, SysUtils, Math, Img32, Img32.SVG.Reader, Img32.Text;

{$WARN SYMBOL_PLATFORM OFF}

{$R dialog.res}

const
  extension = 'svg';
  appId = 'SVGShellExtensions';
  appDescription = 'SVG Shell Extensions';
  SID_EXT_ShellExtensions = '{B2980224-58B3-478C-B596-7D2B23F2C041}';
  IID_EXT_ShellExtensions: TGUID = SID_EXT_ShellExtensions;

  SID_IThumbnailProvider = '{E357FCCD-A995-4576-B01F-234630154E96}';
  IID_IThumbnailProvider: TGUID = SID_IThumbnailProvider;

type
  TWTS_ALPHATYPE = (WTSAT_UNKNOWN, WTSAT_RGB, WTSAT_ARGB);
  IThumbnailProvider = interface(IUnknown)
    [SID_IThumbnailProvider]
    function GetThumbnail(cx: Cardinal; out hbmp: HBITMAP;
      out at: TWTS_ALPHATYPE): HRESULT; stdcall;
  end;

  TSvgShellExt = class(TComObject,
    IPreviewHandler, IThumbnailProvider, IInitializeWithStream)
  strict private
    function IInitializeWithStream.Initialize = IInitializeWithStream_Init;
    //IPreviewHandler
    function DoPreview: HRESULT; stdcall;
    function QueryFocus(var phwnd: HWND): HRESULT; stdcall;
    function SetFocus: HRESULT; stdcall;
    function SetRect(var prc: TRect): HRESULT; stdcall;
    function SetWindow(hwnd: HWND; var prc: TRect): HRESULT; stdcall;
    function TranslateAccelerator(var pmsg: tagMSG): HRESULT; stdcall;
    function Unload: HRESULT; stdcall;
    //IThumbnailProvider
    function GetThumbnail(cx: Cardinal; out hbmp: HBITMAP; out at: TWTS_ALPHATYPE): HRESULT; stdcall;
    //IInitializeWithStream
    function IInitializeWithStream_Init(const pstream: IStream;
      grfMode: DWORD): HRESULT; stdcall;
  private
    FBounds   : TRect;
    fParent   : HWND;
    fDialog   : HWND;
    fSvgRead  : TSvgReader;
    fStream   : IStream;
    procedure CleanupObjects;
    procedure RedrawDialog;
  public
    destructor Destroy; override;
  end;

implementation

//------------------------------------------------------------------------------
// Miscellaneous functions
//------------------------------------------------------------------------------

function GetStreamSize(stream: IStream): Cardinal;
var
  statStg: TStatStg;
begin
  if stream.Stat(statStg, STATFLAG_NONAME) = S_OK then
    Result := statStg.cbSize else
    Result := 0;
end;
//------------------------------------------------------------------------------

function SetStreamPos(stream: IStream; pos: Int64): Int64;
var
  res: LargeUInt;
begin
  stream.Seek(pos, STREAM_SEEK_SET, res);
  Result := res;
end;
//------------------------------------------------------------------------------

procedure CheckAlpha(var img: TImage32);
var
  i: integer;
  pc: PARGB;
begin
  pc := PARGB(img.PixelBase);
  for i := 0 to High(img.Pixels) do
    if pc.A > 0 then Exit else
    inc(pc);
  img.SetAlpha(255);
end;
//------------------------------------------------------------------------------

function CalcStride(width, bpp: integer): integer;
begin
  Result := (((width * bpp) + 31) and not 31) shr 3;
end;
//------------------------------------------------------------------------------

type
  PRgbTriple = ^TRgbTriple;
  TRgbTriple = packed record
    r, g, b: byte;
  end;

function Make24BitBitmapFromPxls(const img: TImage32): HBitmap;
var
  i,j, len, stride, wSpace: integer;
  src : PARGB;
  dst : PRgbTriple;
  bi  : TBitmapInfoHeader;
begin
  Result := 0;
  len := Length(img.pixels);
  if len <> img.width * img.height then Exit;
  stride := CalcStride(img.Width, 24);
  wSpace := stride - (img.Width *3);

  FillChar(bi, sizeof(bi), #0);
  bi.biSize := sizeof(bi);
  bi.biWidth := img.width;
  bi.biHeight := -img.height;
  bi.biPlanes := 1;
  bi.biBitCount := 24;
  bi.biSizeImage := img.Height * stride;
  bi.biCompression := BI_RGB;
  src := PARGB(img.PixelBase);
  Result := CreateDIBSection(0,
    PBitmapInfo(@bi)^, DIB_RGB_COLORS, Pointer(dst), 0, 0);
  for i := 0 to img.Height -1 do
  begin
    for j := 0 to img.Width -1 do
    begin
      dst.r := src.R; dst.g := src.G; dst.b := src.B;
      inc(src); inc(dst);
    end;
    PByte(dst) := PByte(dst) + wSpace;
  end;
end;
//------------------------------------------------------------------------------

function Make32BitBitmapFromPxls(const img: TImage32): HBitmap;
var
  len : integer;
  dst : PARGB;
  bi  : TBitmapV4Header;
begin
  Result := 0;
  len := Length(img.pixels);
  if len <> img.width * img.height then Exit;
  FillChar(bi, sizeof(bi), #0);
  bi.bV4Size := sizeof(TBitmapV4Header);
  bi.bV4Width := img.width;
  bi.bV4Height := -img.height;
  bi.bV4Planes := 1;
  bi.bV4BitCount := 32;
  bi.bV4SizeImage := len *4;
  bi.bV4V4Compression := BI_RGB;
  bi.bV4RedMask       := $FF shl 16;
  bi.bV4GreenMask     := $FF shl 8;
  bi.bV4BlueMask      := $FF;
  bi.bV4AlphaMask     := Cardinal($FF) shl 24;

  Result := CreateDIBSection(0,
    PBitmapInfo(@bi)^, DIB_RGB_COLORS, Pointer(dst), 0, 0);
  Move(img.pixels[0], dst^, len * 4);
end;
//------------------------------------------------------------------------------

function ClampByte(val: double): byte; inline;
begin
  if val <= 0 then result := 0
  else if val >= 255 then result := 255
  else result := Round(val);
end;

//------------------------------------------------------------------------------
// TSvgShellExt
//------------------------------------------------------------------------------

destructor TSvgShellExt.Destroy;
begin
  CleanupObjects;
  if Assigned(fSvgRead) then
    FreeAndNil(fSvgRead);
  fStream := nil;
  inherited Destroy;
end;
//------------------------------------------------------------------------------

procedure TSvgShellExt.CleanupObjects;
var
  imgCtrl: HWnd;
  bm: HBitmap;
begin
  if fDialog <> 0 then
  begin
    imgCtrl := GetDlgItem(fDialog, 101);
    //https://devblogs.microsoft.com/oldnewthing/20140219-00/?p=1713
    bm := SendMessage(imgCtrl, STM_SETIMAGE, IMAGE_BITMAP, 0);
    DeleteObject(bm);
    DestroyWindow(fDialog);
    fDialog := 0;
  end;
end;
//------------------------------------------------------------------------------

procedure TSvgShellExt.RedrawDialog;
var
  l,t,w,h : integer;
  imgCtrl : HWnd;
  img     : TImage32;
  bm,oldBm: HBitmap;
begin
  if fDialog = 0 then Exit;
  SetWindowPos(fDialog, 0, FBounds.left, FBounds.top,
    RectWidth(FBounds), RectHeight(FBounds),
    SWP_NOZORDER or SWP_NOACTIVATE);

  w := RectWidth(FBounds);
  h := RectHeight(FBounds);
  img := TImage32.Create(w, h);
  try
    fSvgRead.DrawImage(img, true);
    l := (w - img.Width) div 2;
    t := (h - img.Height) div 2;
    CheckAlpha(img);
    bm := Make32BitBitmapFromPxls(img);
  finally
    img.Free;
  end;
  imgCtrl := GetDlgItem(fDialog, 101);
  SetWindowPos(imgCtrl, 0, l,t,w,h, SWP_NOZORDER or SWP_NOACTIVATE);

  //https://devblogs.microsoft.com/oldnewthing/20140219-00/?p=1713
  oldBm := SendMessage(imgCtrl, STM_SETIMAGE, IMAGE_BITMAP, bm);
  if oldBm <> 0 then DeleteObject(oldBm);
  DeleteObject(bm);
end;
//------------------------------------------------------------------------------

function DlgProc(dlg: HWnd; msg, wPar: WPARAM; lPar: LPARAM): Bool; stdcall;
begin
  case msg of
    WM_CTLCOLORDLG, WM_CTLCOLORSTATIC:
      Result := Bool(GetStockObject(WHITE_BRUSH));
    else
      Result := False;
  end;
end;
//------------------------------------------------------------------------------

function TSvgShellExt.DoPreview: HRESULT;
var
  size,dum  : Cardinal;
  ms: TMemoryStream;
begin
  //MessageBox(0, 'DoPreview', '',0);
  result := S_OK;
  if (fParent = 0) or FBounds.IsEmpty then Exit;

  if not Assigned(fSvgRead) then
    fSvgRead := TSvgReader.Create;

  CleanupObjects;

  size := GetStreamSize(fStream);
  if size = 0 then Exit;
  ms := TMemoryStream.Create;
  try
    ms.SetSize(size);
    SetStreamPos(fStream, 0);
    fStream.Read(ms.Memory, size, @dum);
    fSvgRead.LoadFromStream(ms);
  finally
    ms.Free;
  end;

  if fSvgRead.IsEmpty then Exit;

  //create the display dialog containing an image control
  fDialog := CreateDialog(hInstance,
    MAKEINTRESOURCE(1), fParent, @DlgProc);
  //draw and show the display dialog
  RedrawDialog;
  ShowWindow(fDialog, SW_SHOW);
end;
//------------------------------------------------------------------------------

function TSvgShellExt.QueryFocus(var phwnd: HWND): HRESULT;
begin
  phwnd := GetFocus;
  result := S_OK;
end;
//------------------------------------------------------------------------------

function TSvgShellExt.SetFocus: HRESULT;
begin
  result := S_OK;
end;
//------------------------------------------------------------------------------

function TSvgShellExt.SetRect(var prc: TRect): HRESULT;
begin
  FBounds := prc;
  RedrawDialog;
  result := S_OK;
end;
//------------------------------------------------------------------------------

function TSvgShellExt.SetWindow(hwnd: HWND; var prc: TRect): HRESULT;
begin
  //MessageBox(0, 'SetWindow', '',0);
  if (hwnd <> 0) then fParent := hwnd;
  if (@prc <> nil) then FBounds := prc;
  CleanupObjects;
  result := S_OK;
end;
//------------------------------------------------------------------------------

function TSvgShellExt.TranslateAccelerator(var pmsg: tagMSG): HRESULT;
begin
  result := S_FALSE
end;
//------------------------------------------------------------------------------

function TSvgShellExt.Unload: HRESULT;
begin
  CleanupObjects;
  if Assigned(fSvgRead) then FreeAndNil(fSvgRead);
  fStream := nil;
  fParent := 0;
  result := S_OK;
end;
//------------------------------------------------------------------------------

function TSvgShellExt.IInitializeWithStream_Init(const pstream: IStream;
  grfMode: DWORD): HRESULT;
begin
  //MessageBox(0, 'Init', '',0);
  fStream := nil;
  fStream := pstream;
  result := S_OK;
end;
//------------------------------------------------------------------------------

function TSvgShellExt.GetThumbnail(cx: Cardinal;
  out hbmp: HBITMAP; out at: TWTS_ALPHATYPE): HRESULT;
var
  size, dum : Cardinal;
  w,h       : integer;
  scale     : double;
  img       : TImage32;
  ms        : TMemoryStream;
  svgr      : TSvgReader;
begin
  result := S_FALSE;
  if fStream = nil then Exit;

  //get file contents and put into qoiBytes
  size := GetStreamSize(fStream);
  SetStreamPos(fStream, 0);

  img := TImage32.Create(1024, 1024);
  try

    svgr := TSvgReader.Create;
    ms := TMemoryStream.Create;
    try
      ms.SetSize(size);
      result := fStream.Read(ms.Memory, size, @dum);
      svgr.LoadFromStream(ms);
      svgr.DrawImage(img, true);
    finally
      ms.Free;
      svgr.Free;
    end;
    if img.IsEmpty then Exit;

    at := WTSAT_ARGB;
    scale := Min(cx/img.width, cx/img.height);
    w := Round(img.width * scale);
    h := Round(img.height * scale);
    img.Resize(w, h);
    //img.SetBackgroundColor(clWhite32);
    hbmp := Make32BitBitmapFromPxls(img);
  finally
    img.Free;
  end;
end;

initialization
  FontManager.Load('Arial');
  FontManager.Load('Arial Bold');
  FontManager.Load('Times New Roman');

  TComObjectFactory.Create(ComServer,
    TSvgShellExt, IID_EXT_ShellExtensions,
    appId, appDescription, ciMultiInstance, tmApartment);
end.
