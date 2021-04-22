unit unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Oracle, DB, OracleData, ComCtrls,StrUtils, ExtCtrls,ComObj;

type
  TForm1 = class(TForm)
    ProgressBar1: TProgressBar;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Button1: TButton;
    StaticText1: TStaticText;
    Edit5: TEdit;
    Label6: TLabel;
    StaticText2: TStaticText;
    StaticText3: TStaticText;
    Label7: TLabel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    Shape1: TShape;
    Edit6: TEdit;
    Label8: TLabel;
    Label9: TLabel;
     procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  year, yearp,month,monthp,day : word;
  pr,empmax,empmin,pr_zap :integer;


implementation
uses unit2;
{$R *.dfm}



 procedure TForm1.FormCreate(Sender: TObject);

 var
  i_np, i_HS : Integer;
  prm, name, psw, HString : string;
begin
  prm := ParamStr(1)+'/'+paramstr(2)+'@'+paramstr(3);
//     prm := ParamStr(1);
  i_np := Pos('/', prm);
  i_HS := Pos('@', prm);
   // есть имя/пароль из ком. строки ? (name/psw@HString)
  if Length(prm) > 2 then begin
    if (i_np = 0) or (i_HS = 0) then begin
     Application.MessageBox('Неверно переданы параметры подключения к БД.' ,
                            'Ошибка', MB_OK + MB_ICONERROR);
     Halt;
    end;
    name :=    LeftStr(prm, i_np - 1);
    psw  :=    MidStr(prm,  i_np + 1, i_HS - i_np - 1);
    HString := RightStr(prm, Length(prm) - i_HS );
    DataModule2.OracleSession1.LogonDatabase := HString;
    DataModule2.OracleSession1.LogonPassword  := psw;
    DataModule2.OracleSession1.LogonUsername:= name;
    DataModule2.OracleSession1.Connected := true;

    if not DataModule2.OracleSession1.Connected then Halt;
    end
   else begin
         DataModule2.OracleLogon1.Execute;
         if not DataModule2.OracleSession1.Connected  then Halt;
   end;


 DecodeDate(Date, Year, Month, Day);
 Edit1.Text := IntToStr(Year);
 Edit2.Text := IntToStr(Month);
 Edit3.Text := IntToStr(Year);
 Edit4.Text := IntToStr(12);
 empmin:=1;
 empmax:=99999;
 Edit5.Text := IntToStr(empmin);
 Edit6.Text := IntToStr(empmax);
       label9.Caption:='База  '+ DataModule2.OracleSession1.LogonDatabase;


  //  ks:=10;   ls:=0;   k:=10;   l:=0; //счетчики для excel

    if VarIsEmpty(XLApp) = true then
  begin
     XLApp:=CreateOleObject('Excel.Application');
     XlApp.Workbooks.Add(ExtractFilePath(ParamStr(0))+'rep801.xls');
     XLApp.Workbooks[1].Worksheets[1].Name:='801';
     Colum:=XLApp.Workbooks[1].WorkSheets['801'].Columns;
     Row:=XLApp.Workbooks[1].WorkSheets['801'].Rows;
     Sheet:=XLApp.Workbooks[1].WorkSheets['801'];
   //  Sheet.Cells[2,5]:=IntToStr(yearp)+' год.';

     k:=10;   l:=0;
    // XLApp.Visible:=true;

  end;

    if VarIsEmpty(XLApps) = true then
  begin
    XLApps:=CreateOleObject('Excel.Application');
    XlApps.Workbooks.Add(ExtractFilePath(ParamStr(0))+'rep801s.xls');
    XLApps.Workbooks[1].Worksheets[1].Name:='801s';

    Colums:=XLApps.Workbooks[1].WorkSheets['801s'].Columns;
    Rows:=XLApps.Workbooks[1].WorkSheets['801s'].Rows;
    Sheets:=XLApps.Workbooks[1].WorkSheets['801s'];
   // Sheets.Cells[2,5]:=IntToStr(yearp)+' год.';
    ks:=10;   ls:=0;
    // XLApps.Visible:=true;
  end;


end;

procedure TForm1.Button1Click(Sender: TObject);
   var    vers,office,FileName,FileNames,mmf:string;
begin
 {
  if VarIsEmpty(XLApp) = true then
  begin
     XLApp:=CreateOleObject('Excel.Application');
     XlApp.Workbooks.Add(ExtractFilePath(ParamStr(0))+'rep801.xls');
     XLApp.Workbooks[1].Worksheets[1].Name:='801';
     Colum:=XLApp.Workbooks[1].WorkSheets['801'].Columns;
     Row:=XLApp.Workbooks[1].WorkSheets['801'].Rows;
     Sheet:=XLApp.Workbooks[1].WorkSheets['801'];
   //  Sheet.Cells[2,5]:=IntToStr(yearp)+' год.';

     k:=10;   l:=0;


  end;

    if VarIsEmpty(XLApps) = true then
  begin
    XLApps:=CreateOleObject('Excel.Application');
    XlApps.Workbooks.Add(ExtractFilePath(ParamStr(0))+'rep801s.xls');
    XLApps.Workbooks[1].Worksheets[1].Name:='801s';

    Colums:=XLApps.Workbooks[1].WorkSheets['801s'].Columns;
    Rows:=XLApps.Workbooks[1].WorkSheets['801s'].Rows;
    Sheets:=XLApps.Workbooks[1].WorkSheets['801s'];
   // Sheets.Cells[2,5]:=IntToStr(yearp)+' год.';
    ks:=10;   ls:=0;
  end;
  }

  year:=StrToInt(Edit1.Text);
  month:=StrToInt(Edit2.Text);
  yearp:=StrToInt(Edit3.Text);
  monthp:=StrToInt(Edit4.Text);
  empmin:=StrToInt(Edit5.Text);
  empmax:=StrToInt(Edit6.Text);


  if(empmin=0) then empmin:=1;
  if(empmax=0) then empmax:=99999;
 // if (CheckBox1.Checked) then pr:=1 else pr:=0;
 pr:=1;
 if RadioButton1.Checked  then pr:=1 ;
 if RadioButton2.Checked  then pr:=0 ;
  pr_zap:=0;

  ProgressBar1.Position:=1;



  DataModule2.raschet1;
 // DataModule2.raschet2;
//  ProgressBar1.Position:=100;
  if pr_zap=0
   then  ShowMessage('Расчет окончен')
   else  showmessage('Нет данных по запросу.');

   {
   if pr_zap<>1 then begin

    vers:=VarToStr(XLApp.version);
    office:='';
    FileName:='';
    office:=Copy(vers,1,Pos('.',vers)-1);

     if length(trim(form1.Edit2.Text))=1
       then  mmf:='0'+ trim(form1.Edit2.Text)
       else  mmf:=trim(form1.Edit2.Text);

     if StrToInt(office)>11 then
      begin
       FileName:='c:\rep801_'+trim(form1.Edit3.Text)+'_'+mmf+'.xlsx';
       FileNames:='c:\rep801s_'+trim(form1.Edit3.Text)+'_'+mmf+'.xlsx';

      end
       else
      begin
       FileName:='c:\rep801_'+trim(form1.Edit3.Text)+'_'+mmf+'.xls';
       FileNames:='c:\rep801_'+trim(form1.Edit3.Text)+'_'+mmf+'.xls';

      end;

     XLApp.Workbooks[1].SaveAs(FileName);
     XLApps.Workbooks[1].SaveAs(FileNames);

       end;

     XLApp.Quit;  XLApp:=Unassigned;
     XLApps.Quit;  XLApps:=Unassigned;
       //end;
    }
  //Close(m
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
 var    vers,office,FileName,FileNames,mmf:string;
 VFalse,VTrue,
VNull,VSNull : OleVariant;
     pr_cf,n_cf:integer; //признак существования файла, N файла

begin
  //  //  ks:=10;   ls:=0;   k:=10;   l:=0; //счетчики для excel
//if (KS>10) OR (K>10) //pr_zap<>1
 // then begin

    vers:=VarToStr(XLApp.version);
    office:='';
    FileName:='';
    office:=Copy(vers,1,Pos('.',vers)-1);

     if length(trim(form1.Edit2.Text))=1
       then  mmf:='0'+ trim(form1.Edit2.Text)
       else  mmf:=trim(form1.Edit2.Text);


      ChDir('c:\report\');

    if not(DirectoryExists('Pensfond'))
      then  createdir('Pensfond');



      pr_cf:=1;  n_cf:=0;
       while pr_cf=1 do
       begin

         if StrToInt(office)>11 then
      begin
       FileName:='c:\report\Pensfond\rep801_'+trim(form1.Edit3.Text)+'_'+mmf+'_'+inttostr(n_cf)+'.xlsx';
       FileNames:='c:\report\Pensfond\rep801s_'+trim(form1.Edit3.Text)+'_'+mmf+'_'+inttostr(n_cf)+'.xlsx';

      end
       else
      begin
       FileName:='c:\report\Pensfond\rep801_'+trim(form1.Edit3.Text)+'_'+mmf+'_'+inttostr(n_cf)+'.xls';
       FileNames:='c:\report\Pensfond\rep801_'+trim(form1.Edit3.Text)+'_'+mmf+'_'+inttostr(n_cf)+'.xls';

      end;

      if (FileExists(FileName)) or (FileExists(FileNames)) then  inc(n_cf) else pr_cf:=0;




       end;      //       while pr_cf=1 do
   //  END; //     if (KS>10) OR (K>10)

  // VFalse:=false; VNull:=0; VSNull:=''; VTrue:=true;


  // IF K>11 THEN
    XLApp.Workbooks[1].SaveAs(FileName);


  // IF KS>11 THEN
  XLApps.Workbooks[1].SaveAs(FileNames);



     XLApp.Quit;  XLApp:=Unassigned;
     XLApps.Quit;  XLApps:=Unassigned;

       // end;


{


    if VarIsEmpty(XLApps) <> true then
  begin
     XLApps.Quit;  XLApps:=Unassigned;
  end;

      if VarIsEmpty(XLApp) <> true then
  begin
     XLApp.Quit;  XLApp:=Unassigned;
 }



end;

end.
