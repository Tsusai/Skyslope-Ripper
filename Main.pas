unit Main;

(*
2019-2020 "Tsusai": Skyslope Ripper, a crappy buggy tool to export transactions
from Skyslope
*)

interface

uses
	Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
	Vcl.Controls, Vcl.Forms, Vcl.Dialogs, REST.Types, Vcl.StdCtrls, REST.Client,
	Data.Bind.Components, Data.Bind.ObjectScope, Vcl.OleCtrls, SHDocVw, MSHTML,
	Clipbrd, URLMon, Vcl.ExtCtrls;

type
		TForm1 = class(TForm)
		Button1: TButton;
		Button2: TButton;
		WebBrowser2: TWebBrowser;
		Memo1: TMemo;
		TimerNextPage: TTimer;
		MemoSaveTimer: TTimer;
		Button3: TButton;
		Timer1: TTimer;
		TimeoutTimer: TTimer;
		Memo2: TMemo;
		WebBrowser1: TWebBrowser;
		procedure Button1Click(Sender: TObject);
		procedure Button2Click(Sender: TObject);
		procedure WebBrowser1NewWindow2(ASender: TObject; var ppDisp: IDispatch;
			var Cancel: WordBool);
		procedure WebBrowser2BeforeNavigate2(ASender: TObject;
			const pDisp: IDispatch; const URL, Flags, TargetFrameName, PostData,
			Headers: OleVariant; var Cancel: WordBool);
		procedure TimerNextPageTimer(Sender: TObject);
		procedure FormCreate(Sender: TObject);
		procedure MemoSaveTimerTimer(Sender: TObject);
		procedure Button3Click(Sender: TObject);
		procedure Timer1Timer(Sender: TObject);
		procedure TimeoutTimerTimer(Sender: TObject);
		procedure FormDestroy(Sender: TObject);
	private
		//myFile : TStringList;
		{ Private declarations }
	public
		{ Public declarations }
	end;


var
	Form1: TForm1;


implementation
uses
	StrUtils,
	System.IOUtils,
	System.Win.Registry;

type
	TSkyStatus = (waiting,clear);

	type TBrowserEmulationAdjuster = class
	private
		class function GetExeName(): String; inline;
	public const
		// Quelle: https://msdn.microsoft.com/library/ee330730.aspx, Stand: 2017-04-26
		IE11_default   = 11000;
		IE11_Quirks    = 11001;
		IE10_force     = 10001;
		IE10_default   = 10000;
		IE9_Quirks     = 9999;
		IE9_default    = 9000;
		/// <summary>
		/// Webpages containing standards-based !DOCTYPE directives are displayed in IE7
		/// Standards mode. Default value for applications hosting the WebBrowser Control.
		/// </summary>
		IE7_embedded   = 7000;
	public
		class procedure SetBrowserEmulationDWORD(const value: DWORD);
	end platform;

var
	PageNumber : integer = 1;
	LogName : string;
	DlURL : string;
	SkyStatus : TSkyStatus;

{$R *.dfm}

class function TBrowserEmulationAdjuster.GetExeName(): String;
begin
	Result := TPath.GetFileName( ParamStr(0) );
end;

class procedure TBrowserEmulationAdjuster.SetBrowserEmulationDWORD(const value: DWORD);
const registryPath = 'Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION';
var
	registry:   TRegistry;
	exeName:   String;
begin
	exeName := GetExeName();

	registry := TRegistry.Create(KEY_SET_VALUE);
	try
		registry.RootKey := HKEY_CURRENT_USER;
		Win32Check( registry.OpenKey(registryPath, True) );
		registry.WriteInteger(exeName, value)
	finally
		registry.Destroy();
	end;
end;

//Finds the pull down boxs
function SelectOptionByValue(const ADocument: IDispatch; const AElementID,
	AOptionValue: WideString): Integer;
var
	HTMLDocument: IHTMLDocument3;
	HTMLElement: IHTMLSelectElement;

	function IndexOfValue(const AHTMLElement: IHTMLSelectElement;
		const AValue: WideString): Integer;
	var
		I: Integer;
	begin
		Result := -1;
		for I := 0 to AHTMLElement.length - 1 do
			if (AHTMLElement.item(I, I) as IHTMLOptionElement).value = AValue then
			begin
				Result := I;
				Break;
			end;
	end;

begin
	Result := -1;
	if Supports(ADocument, IID_IHTMLDocument3, HTMLDocument) then
	begin
		if Supports(HTMLDocument.getElementById(AElementID), IID_IHTMLSelectElement,
			HTMLElement) then
		begin
			Result := IndexOfValue(HTMLElement, AOptionValue);
			HTMLElement.selectedIndex := Result;
		end;
	end;
end;

//Start browsing
procedure TForm1.Button1Click(Sender: TObject);
begin
	WebBrowser1.Navigate('http://www.skyslope.com');
end;

//Catch the "new window" for the download
procedure TForm1.WebBrowser1NewWindow2(ASender: TObject; var ppDisp: IDispatch;
	var Cancel: WordBool);
begin
	ppDisp := WebBrowser2.DefaultDispatch;
end;

//Catch
procedure TForm1.WebBrowser2BeforeNavigate2(ASender: TObject;
	const pDisp: IDispatch; const URL, Flags, TargetFrameName, PostData,
	Headers: OleVariant; var Cancel: WordBool);
begin
	DlURL := URL;
	Cancel := True;

		//ShowMessage('Here´s the URL: '+URL);
		//Clipboard.AsText := URL;

end;

function GetElementById(const Doc: IDispatch; const Id: string): IDispatch;
var
	Document: IHTMLDocument2;     // IHTMLDocument2 interface of Doc
	Body: IHTMLElement2;          // document body element
	Tags: IHTMLElementCollection; // all tags in document body
	Tag: IHTMLElement;            // a tag in document body
	I: Integer;                   // loops thru tags in document body
begin
	Result := nil;
	// Check for valid document: require IHTMLDocument2 interface to it
	if not Supports(Doc, IHTMLDocument2, Document) then
		raise Exception.Create('Invalid HTML document');
	// Check for valid body element: require IHTMLElement2 interface to it
	if not Supports(Document.body, IHTMLElement2, Body) then
		raise Exception.Create('Can''t find <body> element');
	// Get all tags in body element ('*' => any tag name)
	Tags := Body.getElementsByTagName('*');
	// Scan through all tags in body
	for I := 0 to Pred(Tags.length) do
	begin
		// Get reference to a tag
		Tag := Tags.item(I, EmptyParam) as IHTMLElement;
		// Check tag's id and return it if id matches
		if AnsiSameText(Tag.id, Id) then
		begin
			Result := Tag;
			Break;
		end;
	end;
end;

function GrabAgent(const idx: byte) : string;
var
	Elem : IHTMLElement;
begin
	Result := '';
	Elem := GetElementById(Form1.WebBrowser1.Document, 'ContentPlaceHolder1_gvManageTransacion_lblagent_'+IntToStr(idx)) as IHTMLElement;
	if Assigned(Elem) then
	begin
		Result := Elem.innerHTML;
		Result := StringReplace(Result,', ','_',[rfReplaceAll]);
	end;
end;

function GrabAddress(const idx: byte) : string;
var
	Elem : IHTMLElement;
	bidx : byte;
const
	baddies : array[0..8] of char = ('/','\','>','<',':','"','|','?','*');
begin
	Result := '';
	Elem := GetElementById(Form1.WebBrowser1.Document, 'ContentPlaceHolder1_gvManageTransacion_lblPAddress_'+IntToStr(idx)) as IHTMLElement;
	if Assigned(Elem) then
	begin
		Result := Elem.innerHTML;
		Result := StringReplace(Result,',','',[rfReplaceAll]);
		Result := StringReplace(Result,' ','_',[rfReplaceAll]);
		for bidx := Low(baddies) to High(Baddies) do Result := StringReplace(Result,baddies[bidx],'-',[rfReplaceAll]);
	end;
end;

function GrabDate(const idx: byte) : string;
var
	Elem : IHTMLElement;
	tmp : TDate;
begin
	Result := '';
	Elem := GetElementById(Form1.WebBrowser1.Document, 'ContentPlaceHolder1_gvManageTransacion_lblCreatedDate_'+IntToStr(idx)) as IHTMLElement;
	if Assigned(Elem) then
	begin
		Result := Elem.innerHTML;
		tmp := StrToDate(Result);
		Result := FormatDateTime('yyyymmdd',tmp);
	end;
end;

function InitDownload(const idx: byte) : boolean;
var
	Index : integer;
	HTMLWindow : IHTMLWindow2;
begin
	Result := false;
	Index := SelectOptionByValue(Form1.WebBrowser1.Document, 'ContentPlaceHolder1_gvManageTransacion_ddlOption_'+IntToStr(idx), '1');

	if Index <> -1 then
	begin
		Result := true;
		HTMLWindow := IHTMLDocument2(Form1.WebBrowser1.Document).parentWindow;
		Skystatus := Waiting;
		Form1.TimeoutTimer.Enabled := true;
		HTMLWindow.execScript('skyslope.utilities.showSpinner();setTimeout(''__doPostBack(\''ctl00$ContentPlaceHolder1$gvManageTransacion$ctl03$ddlOption\'',\''\'')'', 0)','javascript');
		//ShowMessage('Option was found and selected on index: ' + IntToStr(Index))
	end else
		ShowMessage('Option was not found or the function failed (probably due to ' +
			'invalid input document)!');
end;

function NextPage : boolean;
var
	HTMLWindow : IHTMLWindow2;
	Elem : IHTMLElement;
begin
	Result := false;
	Elem := GetElementById(Form1.WebBrowser1.Document, 'ContentPlaceHolder1_gvManageTransacion_imgBtnNext') as IHTMLElement;
	if Assigned(Elem) then
	begin
		Inc(PageNumber);
		HTMLWindow := IHTMLDocument2(Form1.WebBrowser1.Document).parentWindow;
		HTMLWindow.execScript(
			'__doPostBack(''ctl00$ContentPlaceHolder1$gvManageTransacion$ctl13$imgBtnNext'','''')'
		,'javascript');
		Result := True;
	end;
	//Form1.myFile.SaveToFile('filenames.txt');
end;


function GrabFile2(const URL, SaveTo : string) : integer;
begin
	Result := URLDownloadToFile(nil,PChar(URL),PChar(SaveTo),0,nil);
	if not (S_OK = Result) then
	begin
		Form1.Memo2.Lines.Add(Format('Page %d - Failed To Grab %s, Result: %d',[PageNumber,SaveTo,Result]));
	end;
	SkyStatus := Clear;
end;

function ExtractUrlFileName(const AUrl: string): string;
var
	i: Integer;
begin
	i := LastDelimiter('/', AUrl);
	Result := Copy(AUrl, i + 1, Length(AUrl) - (i));
end;

procedure TForm1.Button2Click(Sender: TObject);
var
	index : byte;
	Agent : string;
	Address: string;
	CDate : string;
	FName : string;
begin
	SkyStatus := Clear;
	for index := 0 to 9 do
	begin
		DlURL := '';
		FName := '';

		Agent := GrabAgent(index);
		if Agent = '' then
		begin
			Memo2.Lines.Add(Format('Item %d-%d - Error Getting Agent',[PageNumber,index+1]));
			break;
		end;
		Memo1.Lines.Add('Agent: '+Agent);

		Address := GrabAddress(index);
		if Address = '' then
		begin
			Memo2.Lines.Add(Format('Item %d-%d - Error Getting Address',[PageNumber,index+1]));
			break;
		end;
		Memo1.Lines.Add('Address: '+Address);

		CDate := GrabDate(index);
		if Address = '' then
		begin
			Memo2.Lines.Add(Format('Item %d-%d - Error Getting Date',[PageNumber,index+1]));
			break;
		end;
		Memo1.Lines.Add('CloseDate: '+CDate);


		FName := Format('s:\%s_%s_%s.zip',[Agent,CDate,Address]);
		Memo1.Lines.Add('SaveName: '+FName);
		//Form1.myFile.Add(FName);
		if FileExists(FName) then begin
			Memo1.Lines.Add('File Already Exists, Skipping');
			continue;
		end;

		if InitDownload(index) then
		begin
			Memo1.Lines.Add('Waiting on URL');
			while DlURL = '' do Application.ProcessMessages;
			TimeoutTimer.Enabled := false;
			if DlURL = 'Timeout' then
			begin
				Memo1.Lines.Add('URL 4 Minute Timeout Reached, Skipping');
				Memo2.Lines.Add(Format('Item %d-%d - Error Getting URL',[PageNumber,index+1]));
				Continue;
			end else
			begin
				Memo1.Lines.Add('URL = '+ DlURL);
				GrabFile2(DlURL,FName);
				while Skystatus = Waiting do Application.ProcessMessages;
			end;

		end else
		begin
			Continue;
		end;
	end;
	Memo1.Lines.Add('Page Complete');
	//Sleep for 10 seconds to be sure everything caught up
	if NextPage then TimerNextPage.Enabled := true;
end;


procedure TForm1.Button3Click(Sender: TObject);
begin
	PageNumber := 1;
	Timer1.Enabled := true;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
	LogName := FormatDateTime('yyyy_mm_dd_hh_mm_ss',Now) + '.txt';
	TBrowserEmulationAdjuster.SetBrowserEmulationDWORD(TBrowserEmulationAdjuster.IE11_Quirks);
	WebBrowser1.Navigate('http://www.skyslope.com');
	//myFile := TStringList.Create;
	//myFile.Sorted := false;
	//myFile.Duplicates := dupAccept;
	//myFile.Clear;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
	//myFile.Free;
end;

procedure TForm1.MemoSaveTimerTimer(Sender: TObject);
begin
	Memo1.Lines.SaveToFile(LogName);
end;

procedure TForm1.TimeoutTimerTimer(Sender: TObject);
begin
	DlURL := 'Timeout';
	TimeoutTimer.Enabled := false;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin
	NextPage;
	//if PageNumber = 208 then Timer1.Enabled := false;
end;

procedure TForm1.TimerNextPageTimer(Sender: TObject);
begin
	TimerNextPage.Enabled:= False;
	Button2.Click;
end;

end.
