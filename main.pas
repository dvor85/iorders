unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Dialogs, ExtCtrls, StdCtrls, ComCtrls, IniFiles,
  DOM, IdHTTP, IBConnection, Windows, Controls, Menus, Updater;

type



  { TForm1 }

  TForm1 = class(TForm)
    HttpClient: TIdHTTP;
    myIbConnection: TIBConnection;
    ListBox1: TListBox;
    Memo1: TMemo;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    PopupMenu1: TPopupMenu;
    Splitter1: TSplitter;
    StatusBarBottom: TStatusBar;
    StatusBarTop: TStatusBar;
    Timer1: TTimer;
    TrayIcon1: TTrayIcon;
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure ListBox1Click(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure MenuItem3Click(Sender: TObject);
    procedure MenuItem4Click(Sender: TObject);
    procedure Timer1StartTimer(Sender: TObject);
    procedure Timer1StopTimer(Sender: TObject);

    procedure Timer1Timer(Sender: TObject);
    //procedure GetMailMsgs();
    function getListOrders(var list: TStrings): boolean;
    function getXmlOrder(id_order: integer; var dest: TStream): string;
    function setOrderStatus(id_order: integer; status: integer): string;

    function LoadOrders(): integer;
    procedure TrayIcon1Click(Sender: TObject);
    procedure TrayIcon1DblClick(Sender: TObject);
  private
    { private declarations }
    ini: TIniFile;
    LogFile: string;
    Shutdown: boolean;
    Version: string;
    LoadErrorCnt: integer;
    procedure CustomExceptionHandler(Sender: TObject; E: Exception);
    procedure AppEndSession(Sender: TObject);
    function getNodeValue(var xml: TXmlDocument; dname: string): string;
    procedure BeepOnNewOrders;
    procedure formatOrderInfo(id_order: integer; var xml: TXmlDocument; var orderInfo: TStrings);
    function InsertToBase(id_order: integer; fn: string; var src: TMemoryStream): integer;
    procedure ShowBalloon(Msg: string; flag: TBalloonFlags);
    function getDBOrderStatus(id_order: integer): integer;
  public
    { public declarations }
  end;

var
  Form1: TForm1;
  Upd: TUpdater;




implementation

uses IdMessageParts, base64, sqldb, XMLRead, dateutils;

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
    begin
      ForceDirectories(ExtractFileDir(LogFileName));
      F := TFileStream.Create(LogFileName, fmCreate);
    end;
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

procedure TForm1.FormCreate(Sender: TObject);
begin
  Application.OnException := @CustomExceptionHandler;
  Application.OnEndSession := @AppEndSession;
  Upd := TUpdater.Create;
  LoadErrorCnt := 0;
  Version := '2.16';
  Caption := 'Интернет заказы v.' + Version;
  ini := TIniFile.Create(ChangeFileExt(ParamStr(0), '.ini'));
  LogFile := ini.ReadString('Global', 'Log', ChangeFileExt(ParamStr(0), '.log'));
  TrayIcon1.Hint := Form1.Caption;
  with httpClient do
  begin
    ConnectTimeout := ini.ReadInteger('Http', 'TimeoutConnect', 5000);
    Request.Username := ini.ReadString('Http', 'Uname', '');
    Request.Password := DecodeStringBase64(ini.ReadString('Http', 'Upass', ''));
    Request.BasicAuthentication := True;
    Request.UserAgent := ExtractFileName(ParamStr(0)) + ' v.' + Form1.Version;
  end;
  AddLog('START', LogFile);
  Shutdown := False;

  with myIbConnection do
  begin
    DatabaseName := ini.ReadString('Global', 'DB', '');
    UserName := 'SYSDBA';
    Password := 'masterkey';
  end;

  with Upd do
  begin
    LogFilename := LogFile;
    CurrentVersion := Version;
    VersionIndexURI := ini.ReadString('http', 'updurl', '');
    Username := HttpClient.Request.Username;
    Password := HttpClient.Request.Password;
    SelfTimer := False;
  end;

  Timer1.Interval := ini.ReadInteger('Global', 'Interval', 60000);
  Timer1.Enabled := True;
end;

procedure TForm1.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  if not Shutdown then
    CloseAction := caHide
  else
  begin
    CloseAction := caFree;
    ini.Free;
    upd.Free;
    Timer1.Enabled := False;
    AddLog('EXIT', LogFile);
  end;
end;

procedure TForm1.CustomExceptionHandler(Sender: TObject; E: Exception);
begin
  AddLog(E.Message, LogFile);
end;

procedure TForm1.ShowBalloon(Msg: string; flag: TBalloonFlags);
begin
  with TrayIcon1 do
  begin
    BalloonFlags := flag;
    BalloonTitle := Form1.Caption;
    BalloonHint := Msg;
    //It seems is not working properly???
    BalloonTimeout := ini.ReadInteger('Global', 'BalloonTimeout', 3000);
    ShowBalloonHint;
  end;
end;

procedure TForm1.ListBox1Click(Sender: TObject);
var
  k: integer;
  orderInfo: TStrings;
begin
  k := ListBox1.ItemIndex;
  if k < 0 then
    exit;
  orderInfo := TStrings(ListBox1.Items.Objects[k]);
  Memo1.Lines := orderInfo;
end;

procedure TForm1.MenuItem1Click(Sender: TObject);
begin
  MessageDlg(Form1.Caption,
    'Delivery InternetOrders v.' + Version + ', Dmitriy Vorotilin, dvor85@mail.ru',
    mtInformation, [mbYes], '');
end;

procedure TForm1.MenuItem2Click(Sender: TObject);
begin
  Form1.ShowOnTop;
end;

procedure TForm1.MenuItem3Click(Sender: TObject);
begin
  LoadOrders();
end;

procedure TForm1.MenuItem4Click(Sender: TObject);
begin
  if MessageDlg(Form1.Caption, 'Вы уверены что надо выйти?', mtConfirmation, mbYesNo, '') = mrYes then
  begin
    Shutdown := True;
    Close;
  end;
end;

procedure TForm1.Timer1StartTimer(Sender: TObject);
begin
  AddLog('Process resume', LogFile);
end;

procedure TForm1.Timer1StopTimer(Sender: TObject);
begin
  AddLog('Process stop', LogFile);
end;



procedure TForm1.formatOrderInfo(id_order: integer; var xml: TXmlDocument; var orderInfo: TStrings);
var
  s: string;
  j: integer;
  dn: TDOMNode;
begin
  try
    s := 'ЗАКАЗ №' + IntToStr(id_order) + ' НА СУММУ: ' + getNodeValue(xml, 'order_summ') + ' руб.';
    orderInfo.Add(s);
    orderInfo.Add('-----------------------------------------');
    s := 'Имя: ' + getNodeValue(xml, 'f_name') + ' ' + getNodeValue(xml, 'm_name') +
      ' ' + getNodeValue(xml, 'l_name');
    orderInfo.Add(s);
    s := getNodeValue(xml, 'organization');
    if s <> '' then
      orderInfo.Add('Организация: ' + s);
    s := 'Адрес: г.' + getNodeValue(xml, 'town') + ', ул.' + getNodeValue(xml, 'street') +
      ', д.' + getNodeValue(xml, 'house') + ', стр.' + getNodeValue(xml, 'building') +
      ', подъезд: ' + getNodeValue(xml, 'entry') + ', кв.(офис): ' + getNodeValue(xml, 'flat') +
      ', этаж: ' + getNodeValue(xml, 'floor') + ', код входа: ' + getNodeValue(xml, 'codeentry');
    orderInfo.Add(s);
    s := 'Телефон: ' + getNodeValue(xml, 'phone1');
    orderInfo.Add(s);
    s := getNodeValue(xml, 'email');
    if s <> '' then
      orderInfo.Add('Email: ' + s);
    s := getNodeValue(xml, 'adv_info');
    if s <> '' then
      orderInfo.Add('Примечание: ' + s);
    s := 'СОСТАВ:';
    orderInfo.Add(s);
    orderInfo.Add(format('%-10s | %-50s | %-10s | %-10s', ['Код', 'Наименование', 'Кол-во', 'Цена']));
    orderInfo.Add('------------------------------------------------------------');
    dn := xml.DocumentElement.FindNode('menu');
    for j := 0 to dn.ChildNodes.Count - 1 do
      with dn.ChildNodes.Item[j].FindNode('item') do
      begin
        s := format('%-10s | %-50s | %-10s | %-10s',
          [UTF8Encode(Attributes.GetNamedItem('code').TextContent),
          UTF8Encode(Attributes.GetNamedItem('name').TextContent), UTF8Encode(
          Attributes.GetNamedItem('quantity').TextContent), UTF8Encode(
          Attributes.GetNamedItem('price').TextContent)]);
        orderInfo.Add(s);
      end;
    orderInfo.Add('------------------------------------------------------------');
    s := format('%89s', ['Всего: ' + getNodeValue(xml, 'order_summ') + ' руб.']);
    orderInfo.Add(s);

  except
    on e: Exception do
    begin
      AddLog(E.Message + ' in function "formatOrderInfo"', LogFile);
    end;

  end;
end;

function TForm1.getNodeValue(var xml: TXmlDocument; dname: string): string;
begin
  try
    Result := UTF8Encode(xml.DocumentElement.FindNode(dname).TextContent);
  except
    Result := '';
  end;
end;

procedure TForm1.BeepOnNewOrders;
begin
  Windows.Beep(500, 300);
  Windows.Beep(700, 500);
end;

function TForm1.LoadOrders(): integer;
var
  orders: TStrings;
  xmlstream: TMemoryStream;
  xml: TXMLDocument;
  orderInfo: TStrings;
  fn: string;
  i: integer;
  err: boolean;
  RetCode: string;
  id_order: integer;
  delivery_id_order: integer;
  FS: TFormatSettings;
  dateorder: TDateTime;
  status_order: integer;
  sso: integer;
  dbos: integer;
begin
  Result := 0;
  err := False;
  orders := TStringList.Create;
  with FS do
  begin
    LongDateFormat := 'dd.mm.yyyy hh:nn:ss';
    DateSeparator := '.';
    TimeSeparator := ':';
    ShortDateFormat := 'dd.mm.yyyy';
    ShortTimeFormat := 'hh:nn:ss';
  end;
  try
    try
      err := not getListOrders(orders);
      if (not err) and (orders.Count > 0) then
      begin
        myIbConnection.Connected := True;
        for i := 0 to orders.Count - 1 do
        begin
          id_order := StrToInt(orders.Names[i]);
          status_order := StrToInt(orders.ValueFromIndex[i]);
          delivery_id_order :=
            id_order + ini.ReadInteger('Global', 'DeliveryOrderStart', 0);
          xmlstream := TMemoryStream.Create;
          xml := TXMLDocument.Create;
          try
            try
              dbos := getDBOrderStatus(delivery_id_order);
              if (dbos = -$ff) and (status_order = 0) then
              begin
                fn := getXmlOrder(id_order, TStream(xmlstream));
                xmlstream.Position := 0;
                ReadXMLFile(xml, xmlstream);
                InsertToBase(delivery_id_order, fn, xmlstream);
                orderInfo := TStringList.Create;
                formatOrderInfo(id_order, xml, orderInfo);
                dateorder := IncSecond(StrToDateTime(getNodeValue(xml, 'wait_time'), FS),
                  -ini.ReadInteger('Global', 'DeliveryTime', 0));
                ListBox1.Items.InsertObject(0,
                  format('Заказ %d от %s', [delivery_id_order, DateTimeToStr(dateorder)]),
                  TObject(orderInfo));
                AddLog(format('New order %d with name: %s', [delivery_id_order, fn]), LogFile);
                Inc(Result);
              end;
              case dbos of
                -$ff, 0: sso := 1;
                else
                  sso := dbos;
              end;
              RetCode := setOrderStatus(id_order, sso);
              if RetCode <> '0' then
                AddLog(format('Set status %d for order %d return: %s',
                  [sso, delivery_id_order, RetCode]),
                  LogFile);
            except
              on e: Exception do
              begin
                err := True;
                continue;
              end;
            end;
          finally
            xml.Free;
            xmlstream.Free;
            Application.ProcessMessages;
          end;
        end;
      end;

    except
      on e: Exception do
      begin
        AddLog(E.Message + ' in function "LoadOrders"', LogFile);
        err := True;
      end;
    end;
    try
      StatusBarBottom.Panels.Items[0].Text := 'Last check: ' + DateTimeToStr(Now());
      if Result > 0 then
      begin
        StatusBarBottom.Panels.Items[1].Text := Format('Новых заказов: %d', [Result]);
        ListBox1.ItemIndex := 0;
        ListBox1.Click;

        ShowBalloon(StatusBarBottom.Panels.Items[1].Text + #13#10 +
          StatusBarBottom.Panels.Items[0].Text, bfInfo);

        if ini.ReadBool('Global', 'BeepOnNew', False) then
          BeepOnNewOrders;
        LoadErrorCnt := 0;
      end
      else
      if err then
      begin
        if LoadErrorCnt > 4 then
        begin
          StatusBarBottom.Panels.Items[1].Text :=
            'Возможны проблемы при получении заказа!';
          ShowBalloon(StatusBarBottom.Panels.Items[1].Text, bfError);
          LoadErrorCnt := 0;
        end
        else
          Inc(LoadErrorCnt);
      end
      else
      begin
        StatusBarBottom.Panels.Items[1].Text := 'Нет новых заказов';
        LoadErrorCnt := 0;
      end;

    except
      on e: Exception do
      begin
        AddLog(E.Message + ' in function "LoadOrders"', LogFile);
      end;
    end;
  finally
    orders.Free;
    myIbConnection.Connected := False;
  end;
end;

procedure TForm1.TrayIcon1Click(Sender: TObject);
begin
  // ShowBalloon(StatusBarBottom.Panels.Items[1].Text + #13#10 +
  //   StatusBarBottom.Panels.Items[0].Text, bfInfo);
end;

procedure TForm1.TrayIcon1DblClick(Sender: TObject);
begin
  Form1.ShowOnTop;
end;

function TForm1.getDBOrderStatus(id_order: integer): integer;
var
  mysqlquery: TSQLQuery;
  mytransaction: TSQLTransaction;
  order: integer;
  deleted: integer;
  status: variant;
begin
  Result := -$ff;
  mysqlquery := TSQLQuery.Create(Form1);
  mytransaction := TSQLTRansaction.Create(Form1);
  try
    try
      mytransaction.DataBase := myIbConnection;
      mytransaction.Active := True;
      with mysqlquery do
      begin
        DataBase := myIbConnection;
        SQL.Text :=
          'SELECT DLV_INTERNETORDERS.ORDER_ID,DLV_INTERNETORDERS.DELETED,DLV_ORDERS.STATUS FROM DLV_INTERNETORDERS LEFT JOIN DLV_ORDERS ON DLV_INTERNETORDERS.ORDER_ID=DLV_ORDERS.ORDER_ID WHERE INETORDER_ID=:INETORDER_ID';
        ParamByName('INETORDER_ID').AsInteger := id_order;
        Active := True;
        if RecordCount > 0 then
        begin
          mysqlquery.First;
          order := FieldByName('ORDER_ID').AsInteger;
          deleted := FieldByName('DELETED').AsInteger;
          status := FieldByName('STATUS').AsVariant;
          if deleted = 1 then
            Result := -1
          else if status = NULL then
            Result := 0
          else
            Result := status;
        end;
        Close;
      end;
    except
      on e: Exception do
      begin
        AddLog(E.Message + ' in function "getDBOrderStatus"', LogFile);
        raise Exception.Create(E.Message);
      end;
    end;
  finally
    mysqlquery.Free;
  end;
end;

function TForm1.InsertToBase(id_order: integer; fn: string; var src: TMemoryStream): integer;
var
  mysqlquery: TSQLQuery;
  mytransaction: TSQLTransaction;
  buf: PChar;
begin
  Result := 0;
  mysqlquery := TSQLQuery.Create(Form1);
  mytransaction := TSQLTRansaction.Create(Form1);
  buf := AllocMem(src.size);
  try
    try
      mytransaction.DataBase := myIbConnection;
      mytransaction.Active := True;
      with mysqlquery do
      begin
        DataBase := myIbConnection;
        ReadOnly := False;
        Transaction := mytransaction;
        //SQL.Text :=
        //  'INSERT INTO DLV_INTERNETORDERS (INETORDER_ID,INETORDER,FILENAME) VALUES ((select next value for dlv_inetorder_id_gen from RDB$DATABASE),:INETORDER,:FILENAME)';
        SQL.Text :=
          'INSERT INTO DLV_INTERNETORDERS (INETORDER_ID,INETORDER,FILENAME) VALUES (:INETORDER_ID,:INETORDER,:FILENAME)';
        src.Position := 0;
        ParamByName('INETORDER_ID').AsInteger := id_order;
        ParamByName('FILENAME').AsString := fn;
        strlcopy(buf, src.Memory, src.size);
        ParamByName('INETORDER').AsBlob := buf;
        ExecSQL;
        mytransaction.Commit;
        Close;
      end;
    except
      on e: Exception do
      begin
        Result := 2;
        AddLog(E.Message + ' in function "InsertToBase"', LogFile);
        raise Exception.Create(E.Message);
      end;
    end;
  finally
    Freemem(buf);
    mysqlquery.Free;
    mytransaction.Free;
  end;
end;

function TForm1.getXmlOrder(id_order: integer; var dest: TStream): string;
var
  Data: TStrings;
begin
  Data := TStringList.Create;
  Data.Add('getxmlorder='+IntToStr(id_order));
  try
    try
      httpClient.Post(ini.ReadString('Http', 'url', ''), Data, dest);
      Result := ExtractHeaderSubItem(httpClient.Response.ContentDisposition, 'filename=');
    except
      on e: Exception do
      begin
        AddLog(E.Message + ' in function "getXmlOrder"', LogFile);
        raise Exception.Create(E.Message);
      end;
    end;
  finally
    Data.Free;
    HttpClient.Disconnect;
  end;

end;

function TForm1.getListOrders(var list: TStrings): boolean;
var
  Data: TStrings;
begin
  Result := True;
  Data := TStringList.Create;
  Data.Add('getlistorders=1');
  try
    try
      list.Text := Trim(httpClient.Post(ini.ReadString('Http', 'url', ''), Data));
    except
      on e: Exception do
      begin
        AddLog(E.Message + ' in function "getListOrders"', LogFile);
        Result := False;
      end;
    end;
  finally
    Data.Free;
    HttpClient.Disconnect;
  end;

end;

function TForm1.setOrderStatus(id_order: integer; status: integer): string;
var
  Data: TStrings;
  res: TStringList;
begin
  Data := TStringList.Create;
  res := TStringList.Create;
  Data.Add('setorderstatus='+IntToStr(status));
  Data.Add('id_order=' + IntToStr(id_order));
  Result := '';
  try
    try
      res.Text := Trim(httpClient.Post(ini.ReadString('Http', 'url', ''), Data));
      Result := res.Values['status'];
    except
      on e: Exception do
      begin
        AddLog(E.Message + ' in function "setOrderStatus"', LogFile);
        raise Exception.Create(E.Message);
      end;
    end;
  finally
    res.Free;
    Data.Free;
    HttpClient.Disconnect;
  end;

end;

{
procedure TForm1.GetMailMsgs();
var
  IMAPClient: TIdIMAP4;
  OpenSSLHandler: TIdSSLIOHandlerSocketOpenSSL;
  res: boolean;
  i, j: integer;
  inbox: string;
  msg, msg2: TIdMessage;
  s: string;
  SearchInfo: array of TIdIMAP4SearchRec;
begin
  ListBox1.Clear;
  IMAPClient := TIdIMAP4.Create(nil);
  try
    OpenSSLHandler := TIdSSLIOHandlerSocketOpenSSL.Create(nil);
    try
      IMAPClient.Host := ini.ReadString('Mail', 'host', '');
      IMAPClient.Port := ini.ReadInteger('Mail', 'port', 993);
      IMAPClient.Username := ini.ReadString('Mail', 'UName', '');
      IMAPClient.Password := DecodeStringBase64(ini.ReadString('Mail', 'UPass', ''));

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
        SearchInfo[0].Date := IncDay(Date, ini.ReadInteger('Global', 'incDays', 0));
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
end; }

procedure TForm1.Timer1Timer(Sender: TObject);
begin
  if not upd.Checked then
    if Upd.NewVersion > Upd.CurrentVersion then
      Upd.UpdateFiles;
  if LoadOrders() > 0 then
  begin

    Form1.ShowOnTop;
  end;
end;

procedure TForm1.AppEndSession(Sender: TObject);
begin
  Shutdown := True;
  AddLog('Shutdown windows', LogFile);
  Close;
end;


end.
