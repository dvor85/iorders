unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, dateutils, Forms, Dialogs, ExtCtrls, StdCtrls, ComCtrls,
  IniFiles, IdMessage, IdIMAP4, IdSSLOpenSSL, IdExplicitTLSClientServerBase,
  IBConnection, sqldb, XMLRead, DOM, Windows, Controls,
  Menus;

type

  { TForm1 }

  TForm1 = class(TForm)
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
    procedure CustomExceptionHandler(Sender: TObject; E: Exception);
    procedure ListBox1Click(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure MenuItem3Click(Sender: TObject);
    procedure MenuItem4Click(Sender: TObject);
    procedure Timer1StartTimer(Sender: TObject);
    procedure Timer1StopTimer(Sender: TObject);



    procedure Timer1Timer(Sender: TObject);
    procedure GetMailMsgs();
    procedure getListOrders(var list: TStrings);
    function getXmlOrder(id_order: integer; var dest: TStream): string;
    function setOrderStatus(id_order: string; status: string): string;

    procedure LoadOrders();
    procedure TrayIcon1DblClick(Sender: TObject);
  private
    { private declarations }
    ini: TIniFile;

    procedure WMQueryEndSession(var Message: TWMQueryEndSession);
      message WM_QUERYENDSESSION;
    function getNodeValue(var xml: TXmlDocument; dname: string): string;
    procedure BeepOnNewOrders;
    procedure formatOrderInfo(id_order: string; var xml: TXmlDocument;
      var orderInfo: TStrings);
    procedure InsertToBase(fn: string; var src: TMemoryStream);
    procedure ShowBalloon(Msg: string; flag: TBalloonFlags);
  public
    { public declarations }
  end;

var
  Form1: TForm1;
  LogFile: string;
  Shutdown: boolean = False;



implementation

uses IdText, IdAttachment, IdHTTP, IdMessageParts, base64;

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

procedure TForm1.ShowBalloon(Msg: string; flag: TBalloonFlags);
begin
  with TrayIcon1 do
  begin
    BalloonFlags := flag;
    BalloonTitle := Form1.Caption;
    BalloonHint := Msg;
    BalloonTimeout:=0;
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
    'Delivery InternetOrders v.2.0, Dmitriy Vorotilin, dvor85@mail.ru',
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
  if MessageDlg(Form1.Caption, 'Вы уверены что надо выйти?', mtConfirmation,
    mbYesNo, '') = mrYes then
  begin
    Shutdown:=True;
    Close;
  end;
end;

procedure TForm1.Timer1StartTimer(Sender: TObject);
begin
  AddLog('Program resume', LogFile);
end;

procedure TForm1.Timer1StopTimer(Sender: TObject);
begin
  AddLog('Program stop', LogFile);
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  Application.OnException := @CustomExceptionHandler;

  ini := TIniFile.Create(ChangeFileExt(ParamStr(0), '.ini'));
  LogFile := ini.ReadString('Global', 'Log', ChangeFileExt(ParamStr(0), '.log'));
  Timer1.Interval := ini.ReadInteger('Global', 'Interval', 300000);
  TrayIcon1.Hint := Form1.Caption;
  AddLog('Program start', LogFile);

  LoadOrders();
end;

procedure TForm1.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  if not Shutdown then
    CloseAction := caHide
  else
  begin
    CloseAction := caFree;
    ini.Free;
    AddLog('Exit program', LogFile);
  end;
end;

procedure TForm1.formatOrderInfo(id_order: string; var xml: TXmlDocument;
  var orderInfo: TStrings);
var
  s: string;
  j: integer;
  dn: TDOMNode;
begin
  try
    s := 'ЗАКАЗ №' + id_order + ' НА СУММУ: ' + getNodeValue(xml,
      'order_summ') + ' руб.';
    orderInfo.Add(s);
    orderInfo.Add('-----------------------------------------');
    s := 'Имя: ' + getNodeValue(xml, 'f_name') + ' ' + getNodeValue(xml, 'm_name') +
      ' ' + getNodeValue(xml, 'l_name');
    orderInfo.Add(s);
    s := getNodeValue(xml, 'organization');
    if s <> '' then
      orderInfo.Add('Организация: ' + s);
    s := 'Адрес: г.' + getNodeValue(xml, 'town') + ', ул.' +
      getNodeValue(xml, 'street') + ', д.' + getNodeValue(xml, 'house') +
      ', стр.' + getNodeValue(xml, 'building') + ', подъезд: ' +
      getNodeValue(xml, 'entry') + ', кв.(офис): ' + getNodeValue(xml, 'flat') +
      ', этаж: ' + getNodeValue(xml, 'floor') + ', код входа: ' +
      getNodeValue(xml, 'codeentry');
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
    orderInfo.Add(format('%-10s | %-50s | %-10s | %-10s',
      ['Код', 'Наименование', 'Кол-во', 'Цена']));
    orderInfo.Add('------------------------------------------------------------');
    dn := xml.DocumentElement.FindNode('menu');
    for j := 0 to dn.ChildNodes.Count - 1 do
      with dn.ChildNodes.Item[j].FindNode('item') do
      begin
        s := format('%-10s | %-50s | %-10s | %-10s',
          [UTF8Encode(Attributes.GetNamedItem('code').TextContent),
          UTF8Encode(Attributes.GetNamedItem('name').TextContent),
          UTF8Encode(Attributes.GetNamedItem('quantity').TextContent),
          UTF8Encode(Attributes.GetNamedItem('price').TextContent)]);
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

procedure TForm1.LoadOrders();
var
  orders: TStrings;
  xmlstream: TMemoryStream;
  orderInfo: TStrings;
  fn: string;
  i, oc: integer;
  xml: TXMLDocument;
  err: boolean;
begin
  err := False;
  orders := TStringList.Create;
  oc := 0;
  try
    try
      getListOrders(orders);
      for i := 0 to orders.Count - 1 do
      begin
        xmlstream := TMemoryStream.Create;
        try
          try
            fn := getXmlOrder(StrToInt(orders[i]), TStream(xmlstream));
            xmlstream.Position := 0;
            ReadXMLFile(xml, xmlstream);
            InsertToBase(fn, xmlstream);
            orderInfo := TStringList.Create;
            formatOrderInfo(orders[i], xml, orderInfo);
            ListBox1.Items.InsertObject(0, 'Заказ ' + orders[i] +
              ' от ' + getNodeValue(xml, 'wait_time'), TObject(orderInfo));

            AddLog(format('New order %s with name: %s', [orders[i], fn]), LogFile);
            // AddLog(format('Set status return: %s',
            //   [setOrderStatus(orders[i], '1')]), LogFile);
            Inc(oc);
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
        end;
      end;

      StatusBarBottom.Panels.Items[0].Text := 'Last check: ' + DateTimeToStr(Now());
      if oc > 0 then
      begin
        StatusBarBottom.Panels.Items[1].Text := Format('Новых заказов: %d', [oc]);
        ListBox1.ItemIndex := 0;
        ListBox1.Click;
        ShowBalloon(StatusBarBottom.Panels.Items[1].Text + #13#10 +
          StatusBarBottom.Panels.Items[0].Text, bfInfo);
        if ini.ReadBool('Global', 'BeepOnNew', False) then
          BeepOnNewOrders;
      end
      else
      begin
        StatusBarBottom.Panels.Items[1].Text := 'Нет новых заказов';
        if err then
          ShowBalloon('Возникли ошибки при получении заказа!' + #13#10 +
            'Правильная работа не гарантирована!', bfError);
      end;
    except
      on e: Exception do
      begin
        AddLog(E.Message + ' in function "LoadOrders"', LogFile);
      end;
    end;
  finally
    orders.Free;
  end;
end;

procedure TForm1.TrayIcon1DblClick(Sender: TObject);
begin
  Form1.ShowOnTop;
end;

procedure TForm1.InsertToBase(fn: string; var src: TMemoryStream);
var
  mysqlquery: TSQLQuery;
  mytransaction: TSQLTransaction;
  myIbConnection: TIBConnection;
begin
  myIbConnection := TIBConnection.Create(Form1);
  mysqlquery := TSQLQuery.Create(Form1);
  mytransaction := TSQLTRansaction.Create(Form1);
  try
    try
      myIbConnection.DatabaseName := ini.ReadString('Global', 'DB', '');
      myIbConnection.UserName := 'SYSDBA';
      myIbConnection.Password := 'masterkey';
      myIbConnection.Connected := True;
      mytransaction.DataBase := myIbConnection;
      mytransaction.Active := True;
      with mysqlquery do
      begin
        ReadOnly := False;
        Database := myIbConnection;
        Transaction := mytransaction;
        SQL.Text :=
          'INSERT INTO DLV_INTERNETORDERS (INETORDER_ID,INETORDER,FILENAME) VALUES ((select next value for dlv_inetorder_id_gen from RDB$DATABASE),:INETORDER,:FILENAME)';
        src.Position := 0;
        ParamByName('FILENAME').AsString := fn;
        ParamByName('INETORDER').AsBlob := PChar(src.Memory);
        ExecSQL;
        mytransaction.Commit;
        Close;
      end;
    except
      on e: Exception do
      begin
        AddLog(E.Message + ' in function "InsertToBase"', LogFile);
        raise Exception.Create(E.Message);
      end;
    end;
  finally
    mysqlquery.Free;
    mytransaction.Free;
    myIbConnection.Free;
  end;
end;

function TForm1.getXmlOrder(id_order: integer; var dest: TStream): string;
var
  httpClient: TIdHTTP;
  Data: TStrings;
begin
  httpClient := TIdHTTP.Create;
  httpClient.ConnectTimeout := ini.ReadInteger('Http', 'TimeoutConnect', 5000);
  httpClient.Request.Username := ini.ReadString('Http', 'Uname', '');
  httpClient.Request.Password := DecodeStringBase64(ini.ReadString('Http', 'Upass', ''));
  httpClient.Request.BasicAuthentication := True;

  Data := TStringList.Create;
  Data.Add('getxmlorder=1');
  Data.Add('id_order=' + IntToStr(id_order));
  try
    try
      httpClient.Post(ini.ReadString('Http', 'url', ''), Data, dest);
      Result := ExtractHeaderSubItem(httpClient.Response.ContentDisposition,
        'filename=');
    except
      on e: Exception do
      begin
        AddLog(E.Message + ' in function "getXmlOrder"', LogFile);
        raise Exception.Create(E.Message);
      end;
    end;
  finally
    Data.Free;
    httpClient.Free;
  end;

end;

procedure TForm1.getListOrders(var list: TStrings);
var
  httpClient: TIdHTTP;
  Data: TStrings;
begin
  httpClient := TIdHTTP.Create;
  httpClient.ConnectTimeout := ini.ReadInteger('Http', 'TimeoutConnect', 5000);
  httpClient.Request.Username := ini.ReadString('Http', 'Uname', '');
  httpClient.Request.Password := DecodeStringBase64(ini.ReadString('Http', 'Upass', ''));
  httpClient.Request.BasicAuthentication := True;
  Data := TStringList.Create;
  Data.Add('getlistorders=1');
  try
    try
      list.Text := httpClient.Post(ini.ReadString('Http', 'url', ''), Data);
    except
      on e: Exception do
      begin
        AddLog(E.Message + ' in function "getListOrders"', LogFile);
        raise Exception.Create(E.Message);
      end;
    end;
  finally
    Data.Free;
    httpClient.Free;
  end;

end;

function TForm1.setOrderStatus(id_order: string; status: string): string;
var
  httpClient: TIdHTTP;
  Data: TStrings;
  res: TStringList;
begin
  httpClient := TIdHTTP.Create;
  httpClient.ConnectTimeout := ini.ReadInteger('Http', 'TimeoutConnect', 5000);
  httpClient.Request.Username := ini.ReadString('Http', 'Uname', '');
  httpClient.Request.Password := DecodeStringBase64(ini.ReadString('Http', 'Upass', ''));
  httpClient.Request.BasicAuthentication := True;
  Data := TStringList.Create;
  res := TStringList.Create;
  Data.Add('setorderstatus=1');
  Data.Add('id_order=' + id_order);
  Data.Add('status=' + status);
  Result := '';
  try
    try
      res.Text := httpClient.Post(ini.ReadString('Http', 'url', ''), Data);
      Result := res.Text;
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
    httpClient.Free;
  end;

end;

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
end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin
  //

end;

procedure TForm1.WMQueryEndSession(var Message: TWMQueryEndSession);
begin
  Shutdown := True;
  AddLog('Shutdown windows', LogFile);
  Message.Result := 1;
  inherited;
end;


end.
