unit SvgPreview;

(*******************************************************************************
* Author    :  Angus Johnson                                                   *
* Version   :  1.1                                                            *
* Date      :  21 January 2022                                                 *
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
  Windows, Messages, ActiveX, Classes, ComObj, ComServ, ShlObj, Registry,
  PropSys, Types, SysUtils, Math, Img32, Img32.SVG.Reader, Img32.Text;

{$WARN SYMBOL_PLATFORM OFF}

{$R dialog.res}

const
  extension = '.svg';
  appId = 'SVGShellExtensions';
  appDescription = 'SVG Shell Extensions';
  SID_EXT_ShellExtensions = '{B2980224-58B3-478C-B596-7D2B23F2C041}';
  IID_EXT_ShellExtensions: TGUID = SID_EXT_ShellExtensions;

  SID_IThumbnailProvider = '{E357FCCD-A995-4576-B01F-234630154E96}';
  IID_IThumbnailProvider: TGUID = SID_IThumbnailProvider;

  darkBkColor = $202020;
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
    fBounds   : TRect;
    fParent   : HWND;
    fDialog   : HWND;
    fDarkBrush: HBrush;
    fSvgRead  : TSvgReader;
    fStream   : IStream;
    fDarkModeChecked: Boolean;
    fDarkModeEnabled: Boolean;
    procedure CleanupDialog;
    procedure RedrawDialog;
    procedure CheckDarkMode;
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
  if assigned(stream) and
    (stream.Stat(statStg, STATFLAG_NONAME) = S_OK) then
      Result := statStg.cbSize
  else
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

function Make32BitBitmapFromPxls(const img: TImage32): HBitmap;
var
  len : integer;
  dst : PARGB;
  bi  : TBitmapInfoHeader;
begin
  Result := 0;
  len := Length(img.pixels);
  if len <> img.width * img.height then Exit;
  FillChar(bi, sizeof(bi), #0);
  bi.biSize := sizeof(bi);
  bi.biWidth := img.width;
  bi.biHeight := -img.height;
  bi.biPlanes := 1;
  bi.biBitCount := 32;
  bi.biCompression := BI_RGB;
  Result := CreateDIBSection(0,
    PBitmapInfo(@bi)^, DIB_RGB_COLORS, Pointer(dst), 0, 0);
  Move(img.pixels[0], dst^, len * 4);
end;
//------------------------------------------------------------------------------

function DlgProc(dlg: HWnd; msg, wPar: WPARAM; lPar: LPARAM): Bool; stdcall;
var
  svgShellExt: TSvgShellExt;
begin
  case msg of
    WM_CTLCOLORDLG, WM_CTLCOLORSTATIC:
      begin
        svgShellExt := Pointer(GetWindowLongPtr(dlg, GWLP_USERDATA));
        if Assigned(svgShellExt) and (svgShellExt.fDarkBrush <> 0) then
          Result := Bool(svgShellExt.fDarkBrush) else
          Result := Bool(GetSysColorBrush(COLOR_WINDOW));
      end;
    else
      Result := False;
  end;
end;

//------------------------------------------------------------------------------
// TSvgShellExt
//------------------------------------------------------------------------------

destructor TSvgShellExt.Destroy;
begin
  CleanupDialog;
  if Assigned(fSvgRead) then
    FreeAndNil(fSvgRead);
  fStream := nil;
  inherited Destroy;
end;
//------------------------------------------------------------------------------

procedure TSvgShellExt.CleanupDialog;
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
    if fDarkBrush <> 0 then DeleteObject(fDarkBrush);
    fDarkBrush := 0;
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
  SetWindowPos(fDialog, 0, fBounds.left, fBounds.top,
    RectWidth(fBounds), RectHeight(fBounds),
    SWP_NOZORDER or SWP_NOACTIVATE);

  w := RectWidth(fBounds);
  h := RectHeight(fBounds);
  img := TImage32.Create(w, h);
  try
    fSvgRead.DrawImage(img, true);
    l := (w - img.Width) div 2;
    t := (h - img.Height) div 2;
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

  ShowWindow(fDialog, SW_SHOW);
end;
//------------------------------------------------------------------------------

procedure TSvgShellExt.CheckDarkMode;
var
  reg: TRegistry;
begin
  fDarkModeChecked := true;
  reg := TRegistry.Create(KEY_READ); //specific access rights important here
  try
    reg.RootKey := HKEY_CURRENT_USER;
    fDarkModeEnabled := reg.OpenKey(
      'SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize', false) and
      reg.ValueExists('SystemUsesLightTheme') and
      (reg.ReadInteger('SystemUsesLightTheme') = 0);
  finally
    reg.Free;
  end;
end;
//------------------------------------------------------------------------------

function TSvgShellExt.DoPreview: HRESULT;
var
  size, dummy  : Cardinal;
  ms: TMemoryStream;
begin
  result := S_OK;
  if (fParent = 0) or fBounds.IsEmpty then Exit;

  if not Assigned(fSvgRead) then
    fSvgRead := TSvgReader.Create;

  CleanupDialog;

  if not fDarkModeChecked then
    CheckDarkMode;

  size := GetStreamSize(fStream);
  if size = 0 then Exit;
  ms := TMemoryStream.Create;
  try
    ms.SetSize(size);
    SetStreamPos(fStream, 0);
    fStream.Read(ms.Memory, size, @dummy);
    fSvgRead.LoadFromStream(ms);
  finally
    ms.Free;
  end;

  if fSvgRead.IsEmpty then Exit;

  //create the display dialog containing an image control
  fDialog := CreateDialog(hInstance, MAKEINTRESOURCE(1), fParent, @DlgProc);
  SetWindowLongPtr(fDialog, GWLP_USERDATA, NativeInt(self));
  if fDarkModeEnabled then
    fDarkBrush := CreateSolidBrush(darkBkColor);
  //draw and show the display dialog
  RedrawDialog;
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
  fBounds := prc;
  RedrawDialog;
  result := S_OK;
end;
//------------------------------------------------------------------------------

function TSvgShellExt.SetWindow(hwnd: HWND; var prc: TRect): HRESULT;
begin
  if (hwnd <> 0) then fParent := hwnd;
  if (@prc <> nil) then fBounds := prc;
  CleanupDialog;
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
  CleanupDialog;
  if Assigned(fSvgRead) then FreeAndNil(fSvgRead);
  fStream := nil;
  fParent := 0;
  result := S_OK;
end;
//------------------------------------------------------------------------------

function TSvgShellExt.IInitializeWithStream_Init(const pstream: IStream;
  grfMode: DWORD): HRESULT;
begin
  fStream := nil;
  fStream := pstream;
  result := S_OK;
end;
//------------------------------------------------------------------------------

function TSvgShellExt.GetThumbnail(cx: Cardinal;
  out hbmp: HBITMAP; out at: TWTS_ALPHATYPE): HRESULT;
var
  size  : Cardinal;
  dummy : Cardinal;
  w,h   : integer;
  scale : double;
  img   : TImage32;
  ms    : TMemoryStream;
  svgr  : TSvgReader;
begin
  result := S_FALSE;
  if fStream = nil then Exit;

  //get file contents and put into qoiBytes
  size := GetStreamSize(fStream);
  SetStreamPos(fStream, 0);

  img := TImage32.Create(256, 256);
  try
    svgr := TSvgReader.Create;
    ms := TMemoryStream.Create;
    try
      ms.SetSize(size);
      result := fStream.Read(ms.Memory, size, @dummy);
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
    hbmp := Make32BitBitmapFromPxls(img);
  finally
    img.Free;
  end;
end;

//------------------------------------------------------------------------------
//------------------------------------------------------------------------------

procedure LoadFonts;
begin
  FontManager.Load('Segoe UI');
  FontManager.Load('Segoe UI Black');
  FontManager.Load('Times New Roman');
  FontManager.Load('Segoe UI Symbol');
end;
//------------------------------------------------------------------------------

var
  res: HRESULT;

initialization
  res := OleInitialize(nil);

  LoadFonts; //needed when displaying SVG text
  TComObjectFactory.Create(ComServer,
    TSvgShellExt, IID_EXT_ShellExtensions,
    appId, appDescription, ciMultiInstance, tmApartment);

finalization
  if res = S_OK then OleUninitialize();

end.
