unit ClipViewU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, ExtCtrls;

type
  TForm1 = class(TForm)
    RichEdit: TRichEdit;
    Image: TImage;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure RichEditKeyPress(Sender: TObject; var Key: Char);
    procedure RichEditKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    NextCBViewer: HWnd;
    procedure UpdateViewOfClipboard;
  protected
    procedure WMChangeCBChain(var Msg: TWMChangeCBChain);
      message WM_CHANGECBCHAIN;
    procedure WMDrawClipboard(var Msg: TWMDrawClipboard);
      message WM_DRAWCLIPBOARD;
  public
    { Public declarations }
  end;
var
  Form1: TForm1;

implementation

{$R *.dfm}

uses
  ClipBrd, ComObj, ActiveX, RichEdit;

var
  CF_RTF: Cardinal;

procedure TForm1.FormCreate(Sender: TObject);
begin
  NextCBViewer := SetClipboardViewer(Handle)
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  ChangeClipboardChain(Handle, NextCBViewer)
end;

procedure TForm1.WMChangeCBChain(var Msg: TWMChangeCBChain);
begin
  //if next window in the clipboard chain is the window being removed then
  if Msg.Remove = NextCBViewer then
    //save the handle specified by hwndNext as the next window in the chain
    NextCBViewer := Msg.Next
  else
    //if there is a next window in the chain then
    if NextCBViewer <> 0 then
      //pass the message to it
      SendMessage(NextCBViewer, Msg.Msg, Msg.Remove, Msg.Next)
end;

procedure TForm1.WMDrawClipboard(var Msg: TWMDrawClipboard);
begin
  //Use new clipboard content if appropriate
  UpdateViewOfClipboard;
  //if there is a next window in the chain then
  if NextCBViewer <> 0 then
    with TMessage(Msg) do
      //pass the message to it
      SendMessage(NextCBViewer, Msg, WParam, LParam)
end;

type
  // Structure passed to GetObject and InsertObject
  TREObject = record
    cbStruct: DWord;          // size of structure in bytes
    cp: Longint;              // character position of object
    ClsID: TClsID;            // class identifier of object
    pOleObj: IOleObject;      // OLE object interface
    pStg: IStorage;           // associated storage interface
    pOleSite: IOleClientSite; // associated client site interface
    sizel: TSize;             // size of object (may be 0,0)
    dvaspect,                 // display aspect to use
    dwFlags,                  // object status flags
    dwUser: DWord;            // user-defined value
  end;

  IRichEditOle = interface
    ['{00020D00-0000-0000-C000-000000000046}']
    procedure GetClientSite(out lplpOleSite: IOleClientSite); stdcall;
    function GetObjectCount: Longint; stdcall;
    function GetLinkCount: Longint; stdcall;
    function GetObject(iObj: Longint; out reobject: TREObject;
      dwFlags: DWord): HResult; stdcall;
    function InsertObject(const reobject: TREObject): HResult; stdcall;
    function ConvertObject(iObj: Longint; const clsidNew: TClsId;
      lpStrUserTypeNew: lpCStr): HResult; stdcall;
    function ActivateAs(const clsId, clsIdAs: TClsId): HResult; stdcall;
    function SetHostNames(lpstrContainerApp, 
      lpstrContainerObj: lpCStr): HResult; stdcall;
    function SetLinkAvailable(iObj: Longint; fAvailable: Bool): HResult; stdcall;
    function SetDvaspect(iObj: Longint; dvaspect: DWord): HResult; stdcall;
    function HandsOffStorage(iObj: Longint): HResult; stdcall;
    function SaveCompleted(iObj: Longint; stg: IStorage): HResult; stdcall;
    function InPlaceDeactivate: HResult; stdcall;
    function ContextSensitiveHelp(fEnterMode: Bool): HResult; stdcall;
    function GetClipboardData(const chrg: TCharRange; reco: DWord;
      out dataobj: IDataObject): HResult; stdcall;
    function ImportDataObject(dataobj: IDataObject; cf: TClipFormat;
      hMetaPict: HGlobal): HResult; stdcall;
  end;

procedure TForm1.UpdateViewOfClipboard;
var
  DataObj: IDataObject;
  RichOle: IRichEditOle;
  Data: THandle;
  Palette: HPalette;
begin
  RichEdit.Clear;
  RichEdit.Hide;
  if Clipboard.HasFormat(CF_BITMAP) then
  begin
    Caption := 'Clipboard Viewer: Bitmap';
    Clipboard.Open;
    Palette := Clipboard.GetAsHandle(CF_PALETTE);
    Data := Clipboard.GetAsHandle(CF_BITMAP);
    Image.Picture.LoadFromClipboardFormat(CF_BITMAP, Data, Palette);
    Clipboard.Close;
  end
  else
  if Clipboard.HasFormat(CF_RTF) then
  begin
    Caption := 'Clipboard Viewer: Rich Text';
    RichEdit.Show;
    RichEdit.Perform(EM_GETOLEINTERFACE, 0, LParam(@RichOle));
    if Assigned(RichOle) then
    begin
      OleCheck(OleGetClipboard(DataObj));
      RichOle.ImportDataObject(DataObj, 0, 0);
    end
  end
  else
  if Clipboard.HasFormat(CF_TEXT) then
  begin
    Caption := 'Clipboard Viewer: Text';
    RichEdit.Show;
    RichEdit.Lines.Text := Clipboard.AsText
  end
end;

procedure TForm1.RichEditKeyPress(Sender: TObject; var Key: Char);
begin
  Key := #0
end;

procedure TForm1.RichEditKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key in [VK_DELETE, VK_BACK] then
    Key := 0
end;

initialization
  CF_RTF := RegisterClipboardFormat('Rich Text Format');
  
end.
