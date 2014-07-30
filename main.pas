unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, dateutils, Forms, Dialogs,
  ExtCtrls, StdCtrls, ComCtrls, IniFiles, IdMessage,
  IdIMAP4, IdSSLOpenSSL, IdExplicitTLSClientServerBase;

type

  { TForm1 }

  TForm1 = class(TForm)
    ListBox1: TListBox;
    Memo1: TMemo;
    Splitter1: TSplitter;
    StatusBarBottom: TStatusBar;
    StatusBarTop: TStatusBar;
    Timer1: TTimer;
    TrayIcon1: TTrayIcon;
    procedure FormCreate(Sender: TObject);
    procedure CustomExceptionHandler(Sender: TObject; E: Exception);
    procedure ListBox1Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure GetMailMsgs();
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  Form1: TForm1;
  ini: TIniFile;
  DB: string;
  LogFile: string;
  Interval: integer;
  IncDays: integer;
  BeepOnNew: boolean;

  IMAPClient: TIdIMAP4;
  OpenSSLHandler: TIdSSLIOHandlerSocketOpenSSL;

  Host, UName, UPass: string;
  Port, TimeoutConnect: integer;


implementation

uses IdText, IdAttachment;

{$R *.lfm}

{ TForm1 }

procedure AddLog(LogString: string; LogFileName: string);
var
  F: TFileStream;
  PStr: PChar;
  Str: string;
  LengthLogString: cardinal;
begin
  Str := DateTimeToStr(Now()) + ': ' + LogString + #13#10;
  LengthLogString := Length(Str);
  try
    if FileExists(LogFileName) then
      F := TFileStream.Create(LogFileName, fmOpenWrite)
    else
      F := TFileStream.Create(LogFileName, fmCreate);
  except
    MessageDlg(Form1.Caption, LogString, mtError, [mbYes], 0);
    Exit;
  end;
  PStr := StrAlloc(LengthLogString + 1);
  try
    try
      StrPCopy(PStr, Str);
      F.Position := F.Size;
      F.Write(PStr^, LengthLogString);
    except
      MessageDlg(Form1.Caption, LogString, mtError, [mbYes], 0);
      Exit;
    end;
  finally
    StrDispose(PStr);
    F.Free;
  end;
end;

procedure TForm1.CustomExceptionHandler(Sender: TObject; E: Exception);
begin
  AddLog(E.Message, LogFile);
end;

procedure TForm1.ListBox1Click(Sender: TObject);
var
  k: integer;
  msg: PChar;
begin
  k := ListBox1.ItemIndex;
  if k > -1 then
  begin
    msg := PChar(ListBox1.Items.Objects[k]);
    Memo1.Lines.Text := msg;
  end;

end;


procedure TForm1.FormCreate(Sender: TObject);
begin
  Application.OnException := @CustomExceptionHandler;
  ini := TIniFile.Create(ChangeFileExt(ParamStr(0), '.ini'));
  DB := ini.ReadString('Global', 'DB', '');
  LogFile := ini.ReadString('Global', 'Log', ChangeFileExt(ParamStr(0), '.log'));
  Interval := ini.ReadInteger('Global', 'Interval', 300000);
  IncDays := ini.ReadInteger('Global', 'IncDays', 0);
  BeepOnNew := ini.ReadBool('Global', 'BeepOnNew', False);

  host := ini.ReadString('MailParams', 'host', '');
  UName := ini.ReadString('MailParams', 'UName', '');
  UPass := ini.ReadString('MailParams', 'UPass', '');
  Port := ini.ReadInteger('MailParams', 'Port', 110);
  TimeoutConnect := ini.ReadInteger('MailParams', 'TimeoutConnect', 5000);

  Timer1.Interval := Interval;
  GetMailMsgs();
  // Timer1.Enabled := True;

end;

procedure TForm1.GetMailMsgs();
var
  IMAPClient: TIdIMAP4;
  UsersFolders: TStringList;
  OpenSSLHandler: TIdSSLIOHandlerSocketOpenSSL;
  res: boolean;
  i, j, k: integer;
  inbox, currUID: string;
  cntMsg: integer;
  msg, msg2: TIdMessage;
  s: string;
  BodyTexts: TStringList;
  flags: TIdMessageFlagsSet;
  fileName_MailSource, TmpFolder: string;
  SearchInfo: array of TIdIMAP4SearchRec;
  Fs: TFormatSettings;
begin
  ListBox1.Clear;
  IMAPClient := TIdIMAP4.Create(nil);
  try
    OpenSSLHandler := TIdSSLIOHandlerSocketOpenSSL.Create(nil);
    try
      IMAPClient.Host := Host;
      IMAPClient.Port := Port;
      IMAPClient.Username := UName;
      IMAPClient.Password := UPass;

      OpenSSLHandler.SSLOptions.Method := sslvSSLv23;
      IMAPClient.IOHandler := OpenSSLHandler;
      IMAPClient.UseTLS := utUseImplicitTLS;

      try
        res := IMAPClient.Connect;
        if not res then
        begin
          AddLog('  Unsuccessful connection.', LogFile);
          exit;
        end;

      except
        on e: Exception do
        begin
          AddLog('  Unsuccessful connection.', LogFile);
          AddLog('  (' + Trim(e.Message) + ')', LogFile);
          exit;
        end;
      end;

      try
        IMAPClient.RetrieveOnSelect := rsDisabled;
        inbox := 'INBOX';
        AddLog('Opening folder "' + inbox + '"...', LogFile);
        res := IMAPClient.SelectMailBox(inbox);

        SetLength(SearchInfo, 1);
        SearchInfo[0].SearchKey := skSince;
        SearchInfo[0].Date := IncDay(Date, IncDays);
        //SearchInfo[1].SearchKey := skUnseen;

        if IMAPClient.SearchMailBox(SearchInfo) then
        begin
          msg := TIdMessage.Create(nil);
          msg2 := TIdMessage.Create(nil);
          try
            for I := High(IMAPClient.MailBox.SearchResult) downto 0 do
            begin
              IMAPClient.RetrieveHeader(IMAPClient.MailBox.SearchResult[i], msg);
              if (pos('.xml', msg.Subject) > 0) or (pos('.table', msg.Subject) > 0) then
              begin
                IMAPClient.Retrieve(IMAPClient.MailBox.SearchResult[i], msg2);
                for j := 0 to msg2.MessageParts.Count - 1 do
                begin
                  if msg2.MessageParts.Items[j] is Tidtext then
                  begin
                    s := TIdText(msg2.MessageParts.Items[j]).Body.Text;
                    ListBox1.Items.AddObject(DateTimeToStr(msg2.Date), TObject(s));
                  end
                  else if (msg2.MessageParts.Items[j] is TIdAttachment) then
                  begin
                    try
                      if TIdAttachment(msg2.MessageParts.Items[j]).FileName =
                        msg2.Subject then
                        TIdAttachment(msg2.MessageParts.Items[j]).SaveToFile(
                          msg2.Subject);

                    except
                      continue;
                    end;
                  end;
                end;
              end;
            end;

          finally
            FreeAndNil(msg);
            //FreeAndNil( msg2 );
            //FreeAndNil( BodyTexts )
          end;

        end;

      finally

        IMAPClient.Disconnect;
      end;
    finally
      OpenSSLHandler.Free;
    end;
  finally
    IMAPClient.Free;
  end;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
var
  msgcnt, i: integer;
begin
  GetMailMsgs();
  {IdPOP3_1.Host := stPOP3;
  IdPOP3_1.Port := inPOP3Port;
  IdPOP3_1.Username := stUName;
  IdPOP3_1.Password := stUPass;
  IdPOP3_1.ConnectTimeout := TimeoutConnect;
  IdPOP3_1.Connect;
  IdIMAP4_1.Host:=Host;
  IdIMAP4_1.Port:=Port;
  IdIMAP4_1.Username:=UName;
  IdIMAP4_1.Password:=UPass;
  IdIMAP4_1.ConnectTimeout:=TimeoutConnect;

  try
    try
      if IdIMAP4_1.Connected then
      begin
        // IdMessageDecoderMIME1;
        msgcnt := IdIMAP4_1.CheckMessages;
        for i := msgcnt downto 1 do
        begin
          IdMessage1.Clear;
          if (IdPOP3_1.RetrieveHeader(i, IdMessage1)) then
          begin
            ListBox1.Items.Add(IdMessage1.Subject);
          end;
        end;
      end;
    except
      exit;
    end;
  finally
    IdPOP3_1.disconnect;
  end; }

end;


end.
