
unit Unit2;

interface

uses
  SysUtils, Classes, Oracle, DB, OracleData,Windows,Dialogs ,  Forms,
  Messages,  Variants,  Controls, Math, ComObj,DateUtils,
  StdCtrls,  ComCtrls;

  const  razmas=100;
type
  TDataModule2 = class(TDataModule)
    OracleLogon1: TOracleLogon;
    OracleSession1: TOracleSession;
    OracleDataSet1: TOracleDataSet;
    OracleQuery1del: TOracleQuery;
    OracleQuery1ins: TOracleQuery;
    OracleDataSetkoef: TOracleDataSet;
    OracleDataSetkoefKOEF: TFloatField;
    OracleDataSetkoefINDATE: TDateTimeField;
    OracleQueryufzpl: TOracleQuery;
    OracleQueryufblfss: TOracleQuery;
    OracleDataSetsecret: TOracleDataSet;
    OracleDataSetsecretPSE: TFloatField;
    OracleDataSetuvol: TOracleDataSet;
    OracleDataSetuvolDSCHDATE: TStringField;
    OracleQueryufbol: TOracleQuery;
    OracleQueryufgpd: TOracleQuery;
    ODScount: TOracleDataSet;
    ODScountMONTHYEAR: TFloatField;
    OracleDataSet1EMP: TIntegerField;
    OracleDataSet1YEARZ: TIntegerField;
    OracleDataSet1MONTHZ: TIntegerField;
    OracleDataSet1SHOP: TFloatField;
    OracleDataSet1EXPEND: TFloatField;
    OracleDataSet1SUM2: TFloatField;
    OracleDataSet1SUM3: TFloatField;
    OracleDataSet1SUM4: TFloatField;
    OracleDataSet1SUM5: TFloatField;
    OracleDataSet1SUM22: TFloatField;
    OracleDataSet1SUM33: TFloatField;
    OracleDataSet1SUM44: TFloatField;
    OracleDataSet1SUM55: TFloatField;
    OracleDataSet1FLAGINV: TFloatField;
    OracleDataSet1FLAGCORR: TFloatField;
    OracleDataSet1MONTHYEAR: TFloatField;
  private
    { Private declarations }
       procedure zap_baza(pay:integer;raz:real);
       procedure  pech;
       procedure  pech_it;
       procedure new_page;
       function RoundEx(const AValue: Double; const ADigit: Integer = -2): Double;
       function vic_koef(date1: TDateTime; koef:real):real;
  public
   procedure raschet1;


 //  procedure raschet2;
    { Public declarations }
  end;

var
  DataModule2: TDataModule2;
  i,flaginv,emp,pay,MONTHz,expend,shop:integer;
date1: TDateTime;
  razgpd,razzpl,razblfss,razbol,raz,koef :real;
  k,ks,l,ls: integer;
//sbol,sblfss,sbolm,sblfssm,st_zpl,st_gpd,st_bl :real;
        dschdate: string;  //дата увольнени€
 //  myFile,myFiles : TextFile;
     XLApp,XLApps:Variant;
      Sheet,Colum,Row:Variant;
      Sheets,Colums,Rows:Variant;
       flaginv_old,monthz_old:Integer;
        yearz,k_st,ks_st: integer;
      sum2,sum3,sum4,sum5,sum22,sum33,sum44,sum55:real;
      sum2_i,sum3_i,sum4_i,sum5_i,sum22_i,sum33_i,sum44_i,sum55_i:real;
      mas_kf:array [1..9,1..razmas] of real;  // 7 -koэф  8- мес€ц 9 год
      emp_old,pse,mnyear,mnyearold,expendold,dept_old,year_old,fl_inv_old,tip_nach,tip_nach_old,tip_nach_i  : integer;
      RUzpl,UFzpl,RUgpd,UFgpd,RUbol,UFbol,RUblfss,UFblfss:real;
      KFgpd, KFblfss ,KFbol,KFblfssi ,KFboli,KFzpli,KFzpl:real;
      yearf,monthf:Word;










implementation
 uses unit1;
{$R *.dfm}

function TDataModule2.RoundEx(const AValue: Double; const ADigit: Integer = -2): Double;
var
  s:  String;
  st: Int64;
  sf: Real;
begin
if(AValue>=0) then begin
  s  := FloatToStr(AValue * IntPower(10, -ADigit));
  st := Trunc(StrToFloat(s));
  sf := Frac(StrToFloat(s));
  if sf <  0.5 then Result := st*IntPower(10, ADigit);
  if sf >= 0.5 then Result := (st+1)*IntPower(10, ADigit);
              end
              else
              begin
  s  := FloatToStr(AValue * IntPower(10, -ADigit));
  st := Trunc(StrToFloat(s));
  sf := Frac(StrToFloat(s));
  if sf >  -0.5 then Result := st*IntPower(10, ADigit);
  if sf <= -0.5 then Result := (st-1)*IntPower(10, ADigit);
              end;


//  if Temp >= 0.5 then ScaledFractPart := ScaledFractPart + 1;
//  if Temp <=-0.5 then ScaledFractPart := ScaledFractPart - 1;

end;




function TDataModule2.vic_koef(date1: TDateTime; koef:real):real;
begin


                 DataModule2.OracleDataSetkoef.Close;
                 DataModule2.OracleDataSetkoef.SetVariable('date1',date1);
                 DataModule2.OracleDataSetkoef.SetVariable('kf',koef);        //kfbol
                 DataModule2.OracleDataSetkoef.Open;
                 DataModule2.OracleDataSetkoef.First;
                   if DataModule2.OracleDataSetkoef.RecordCount<>0
                       then   vic_koef:= DataModule2.OracleDataSetkoefKOEF.AsFloat
                       else   vic_koef:=0 ;

end;



procedure TDataModule2.raschet1;

   var  i,j:Integer;
  {    yearz,k_st: integer;
      sum2,sum3,sum4,sum5,sum22,sum33,sum44,sum55:real;
      sum2_i,sum3_i,sum4_i,sum5_i,sum22_i,sum33_i,sum44_i,sum55_i:real;
      mas_kf:array [1..9,1..razmas] of real;  // 7 -koэф  8- мес€ц 9 год
      emp_old,i,j,pse,mnyear,mnyearold,expendold,dept_old,year_old,month_old,fl_inv_old  : integer;
      RUzpl,UFzpl,RUgpd,UFgpd,RUbol,UFbol,RUblfss,UFblfss:real;
      KFgpd, KFblfss ,KFbol,KFblfssi ,KFboli,KFzpli,KFzpl:real;
      yearf,monthf:Word;
   }
    begin
 Form1.ProgressBar1.Min;
 if pr=0  //если указатель пустой то пишем в базу
  then  begin

                   OracleQuery1del.SetVariable('empmin',empmin);
                   OracleQuery1del.SetVariable('empmax',empmax);
                   OracleQuery1del.setVariable('month',month);
                   OracleQuery1del.SetVariable('year',year);


                   with OracleQuery1del do
                      try
                          Form1.StaticText1.Caption := 'del';
                       try
                          Execute;
                   except
                       on E:EOracleError do begin
                       ShowMessage(E.Message);
                       exit;
                       end;
                       end;

                   except
                       on E:EOracleError do ShowMessage(E.Message);
                   end;
               OracleSession1.Commit;
             end;  // if  pr=0



  OracleDataSet1.Close;
  OracleDataSet1.SetVariable('yearp',yearp);
  OracleDataSet1.SetVariable('monthp',monthp);
  OracleDataSet1.SetVariable('empmin',empmin);
  OracleDataSet1.SetVariable('empmax',empmax);
  OracleDataSet1.Open;
  OracleDataSet1.First;
     //  mnyearold=0;

  if OracleDataSet1.RecordCount<>0
     then  begin
                   Form1.ProgressBar1.Max:=OracleDataSet1.RecordCount;
                   Form1.StaticText2.Caption:='0';
                   Form1.StaticText2.Repaint;
                   Form1.StaticText3.Caption:=Inttostr(OracleDataSet1.RecordCount);
                    Form1.StaticText3.Repaint;

      k:=k+1; ks:=ks+1;
     Sheet.Cells[k,3]:='ѕерерасчет за '+IntToStr(yearp)+' год.';
     Sheets.Cells[ks,3]:='ѕерерасчет за '+IntToStr(yearp)+' год.';


  //вычисление коэф.

    for i:=1 to razmas do
      for j:=1 to 9 do
          mas_kf[j,i]:=0;


     ODScount.Close;
     ODScount.SetVariable('yearp',yearp);
     ODScount.SetVariable('empmin',empmin);
     ODScount.SetVariable('empmax',empmax);
     ODScount.Open;
     ODScount.First;
     if ODScount.RecordCount<>0
       then begin

             for i:=1 to ODScount.RecordCount do
              begin
                if ODScountMONTHYEAR.asinteger<>0
                  then   begin

                      yearf:=ODScountMONTHYEAR.asinteger div 100;
                      monthf:=ODScountMONTHYEAR.asinteger- yearf*100;

                     date1:=startofamonth(yearf,monthf);



                 mas_kf[1,i]:=vic_koef(date1,401); //kfzpl
                 mas_kf[2,i]:=vic_koef(date1,405);//kfzpli
                 mas_kf[3,i]:=vic_koef(date1,434);//kfgpd
                 mas_kf[4,i]:=vic_koef(date1,430);//kfbol
                 mas_kf[5,i]:=vic_koef(date1,435);//kfboli
                 mas_kf[6,i]:=vic_koef(date1,430);//kfblfss
                 mas_kf[7,i]:=vic_koef(date1,435);;//kfblfssi
                 mas_kf[8,i]:=ODScountMONTHYEAR.asinteger;
                 mas_kf[9,i]:=yearf;








                        end;
                   ODScount.Next;
                end ;



       end;
       //

       // ODScount.Close;




      emp_old:=0;
      pse:=0;       expendold:=0;
                             k_st:=0;     ks_st:=0;          tip_nach:=0;    tip_nach_old:=0;
          dept_old:=0; year_old:=0; monthz_old:=0; fl_inv_old:=0;
      RUzpl:=0; UFzpl:=0; RAZzpl:=0; RUgpd:=0; UFgpd:=0; RAZgpd:=0;
    RUbol:=0; UFbol:=0; RAZbol:=0; RUblfss:=0;  UFblfss:=0; RAZblfss:=0;

    
                 OracleDataSetuvol.Close;
                 OracleDataSetuvol.SetVariable('emp',OracleDataSet1EMP.AsInteger);
                 OracleDataSetuvol.Open;
                 OracleDataSetuvol.First;
                   if OracleDataSetuvol.RecordCount<>0
                        then  dschdate:=OracleDataSetuvolDSCHDATE.AsString
                        else  dschdate:='';


                   //секретчики
                 OracleDataSetsecret.Close;
                 OracleDataSetsecret.SetVariable('emp',OracleDataSet1EMP.AsInteger);
                 OracleDataSetsecret.SetVariable('monthr',month);
                 OracleDataSetsecret.SetVariable('yearr',year);
                 OracleDataSetsecret.Open;
                 OracleDataSetsecret.First;
                   if OracleDataSetsecret.RecordCount<>0
                        then pse:=OracleDataSetsecretPSE.AsInteger
                        else pse:=0;




   for i:=1 to OracleDataSet1.RecordCount do
     begin

                Form1.ProgressBar1.Position:=i;
                Form1.StaticText2.Caption:=Inttostr(i);
                Form1.StaticText2.Repaint;


                                {

      emp:=OracleDataSet1EMP.AsInteger;
      yearz:=OracleDataSet1YEARZ.AsInteger;
      monthz:=OracleDataSet1MONTHZ.AsInteger;
      shop:=OracleDataSet1SHOP.AsInteger;
      expend:=OracleDataSet1EXPEND.AsInteger;
      SUM2:=OracleDataSet1SUM2.AsFloat;
      SUM3:=OracleDataSet1SUM3.AsFloat;
      SUM4:=OracleDataSet1SUM4.AsFloat;
      SUM5:=OracleDataSet1SUM5.AsFloat;
      SUM22:=OracleDataSet1SUM22.AsFloat;
      SUM33:=OracleDataSet1SUM33.AsFloat;
      SUM44:=OracleDataSet1SUM44.AsFloat;
      SUM55:=OracleDataSet1SUM55.AsFloat;
      flaginv:=OracleDataSet1FLAGINV.AsInteger;
      mnyear:=OracleDataSet1MONTHYEAR.AsInteger;

                                 }

   //   flagcorr:=OracleDataSet1FLAGCORR.AsInteger;

             { if emp_old<>OracleDataSet1EMP.AsInteger then  begin




                 OracleDataSetuvol.Close;
                 OracleDataSetuvol.SetVariable('emp',OracleDataSet1EMP.AsInteger);
                 OracleDataSetuvol.Open;
                 OracleDataSetuvol.First;
                   if OracleDataSetuvol.RecordCount<>0
                        then  dschdate:=OracleDataSetuvolDSCHDATE.AsString
                        else  dschdate:='';


                   //секретчики
                 OracleDataSetsecret.Close;
                 OracleDataSetsecret.SetVariable('emp',OracleDataSet1EMP.AsInteger);
                 OracleDataSetsecret.SetVariable('monthr',month);
                 OracleDataSetsecret.SetVariable('yearr',year);
                 OracleDataSetsecret.Open;
                 OracleDataSetsecret.First;
                   if OracleDataSetsecret.RecordCount<>0
                        then pse:=OracleDataSetsecretPSE.AsInteger
                        else pse:=0;

                    end; // if emp_old<>emp then

                }

                                   kfzpl:=0;
                                   kfzpli:=0;
                                   kfgpd:=0;
                                   kfbol:=0;
                                   kfblfss:=0;
                                   kfboli:=0;
                                   kfblfssi:=0;




              for j:=1 to ODScount.RecordCount do
                begin
                  if OracleDataSet1MONTHYEAR.AsInteger=mas_kf[8,j]
                     then    begin
                                    kfzpl:=mas_kf[1,j];
                                    kfzpli:=mas_kf[2,j];
                                    kfgpd:=mas_kf[3,j];
                                    kfbol:=mas_kf[4,j];
                                    kfblfss:=mas_kf[6,j];
                                    kfboli:=mas_kf[5,j];
                                    kfblfssi:=mas_kf[7,j];








                                   end;
                 end;



                 if OracleDataSet1MONTHYEAR.AsInteger=0
                     then
                     begin


                     date1:=startofamonth(OracleDataSet1YEARZ.AsInteger,OracleDataSet1MONTHZ.AsInteger);

                       kfzpl:=vic_koef(date1,401); //kfzpl
                       kfzpli:=vic_koef(date1,405);//kfzpli
                       kfgpd:=vic_koef(date1,434);//kfgpd
                       kfbol:=vic_koef(date1,430);//kfbol
                       kfboli:=vic_koef(date1,435);
                       kfblfss:=vic_koef(date1,430);
                       kfblfssi:=vic_koef(date1,435);

                     end;




               {    if  (OracleDataSet1SUM2.AsFloat<>0) or (OracleDataSet1SUM22.AsFloat<>0)
                      then tip_nach:=1;

                 if  (OracleDataSet1SUM3.asFloat<>0) or (OracleDataSet1SUM33.AsFloat<>0)
                      then tip_nach:=2;

                   if  (OracleDataSet1SUM4.AsFloat<>0) or (OracleDataSet1SUM44.AsFloat<>0)
                      then tip_nach:=3;

                 if  (OracleDataSet1SUM5.asFloat<>0) or (OracleDataSet1SUM55.AsFloat<>0)
                      then tip_nach:=4;
                }

              //  if OracleDataSet1TIP.AsInteger<>0 then
               //  begin

                   if (OracleDataSet1SUM2.AsFloat<>0) or  (OracleDataSet1SUM22.AsFloat<>0)
                      then     tip_nach_i:=1;

                    if (OracleDataSet1SUM3.AsFloat<>0) or  (OracleDataSet1SUM33.AsFloat<>0)
                      then     tip_nach_i:=2;

                   if (OracleDataSet1SUM4.AsFloat<>0) or  (OracleDataSet1SUM44.AsFloat<>0)
                      then     tip_nach_i:=3;

                    if (OracleDataSet1SUM5.AsFloat<>0) or  (OracleDataSet1SUM55.AsFloat<>0)
                      then     tip_nach_i:=4;


                  if (expendold<>OracleDataSet1EXPEND.AsInteger)
                   or (dept_old<>OracleDataSet1SHOP.AsInteger)
                   or  (year_old<>OracleDataSet1YEARZ.AsInteger)
                   or  (monthz_old<>OracleDataSet1MONTHZ.AsInteger)
                   or (fl_inv_old<>OracleDataSet1FLAGINV.AsInteger)
                   or (emp_old<>OracleDataSet1EMP.AsInteger)
                   or (tip_nach_old<>tip_nach_i)
         then begin

                 if  i<>1 //перва€ запись
                   then begin
                  RAZzpl:=RUzpl-UFzpl;
                  RAZbol:=RUbol-UFbol;
                  RAZblfss:=RUblfss-UFblfss;
                  RAZgpd:=RUgpd-UFgpd;


                  if pr=0  //если указатель пустой то пишем в базу
           then  begin
              if RoundEx(Razzpl,-2)<>0
                 then  zap_baza(801,RoundEx(Razzpl,-2));//  пишем в базу   pay=801

              if RoundEx(Razbol,-2)<>0
                 then zap_baza(807,RoundEx(Razbol,-2)); // пишем в базу   pay=807

              if RoundEx(RAZblfss,-2)<>0
                 then zap_baza(809,RoundEx(RAZblfss,-2)); // пишем в базу   pay=809

              if RoundEx(RAZgpd,-2)<>0
                 then zap_baza(801,RoundEx(RAZgpd,-2)); // пишем в базу   pay=801


                end;  //      if pr=0

                               pech_it;     //печать итоговой строки

                    end;




                     sum2_i:=0; sum22_i:=0;
            sum3_i:=0; sum33_i:=0;
             sum4_i:=0; sum44_i:=0;
              sum5_i:=0; sum55_i:=0;
          RUzpl:=0; UFzpl:=0; RAZzpl:=0; RUgpd:=0; UFgpd:=0; RAZgpd:=0;
    RUbol:=0; UFbol:=0; RAZbol:=0; RUblfss:=0;  UFblfss:=0; RAZblfss:=0;








                  if emp_old<>OracleDataSet1EMP.AsInteger then  begin
                 OracleDataSetuvol.Close;
                 OracleDataSetuvol.SetVariable('emp',OracleDataSet1EMP.AsInteger);
                 OracleDataSetuvol.Open;
                 OracleDataSetuvol.First;
                   if OracleDataSetuvol.RecordCount<>0
                        then  dschdate:=OracleDataSetuvolDSCHDATE.AsString
                        else  dschdate:='';


                   //секретчики
                 OracleDataSetsecret.Close;
                 OracleDataSetsecret.SetVariable('emp',OracleDataSet1EMP.AsInteger);
                 OracleDataSetsecret.SetVariable('monthr',month);
                 OracleDataSetsecret.SetVariable('yearr',year);
                 OracleDataSetsecret.Open;
                 OracleDataSetsecret.First;
                   if OracleDataSetsecret.RecordCount<>0
                        then pse:=OracleDataSetsecretPSE.AsInteger
                        else pse:=0;

                    end; // if emp_old<>emp then


                   if (emp_old<>OracleDataSet1EMP.AsInteger)
                      then if pse=0 then  begin inc(k) ; new_page; end
                                      else   begin inc(ks); ; new_page; end;


                 tip_nach_old:=tip_nach_i;
                 emp_old:=OracleDataSet1EMP.AsInteger ;
                 monthz_old:=OracleDataSet1MONTHZ.AsInteger;
                 expendold:=OracleDataSet1expend.AsInteger;
                 year_old:=OracleDataSet1YEARZ.AsInteger ;

                 dept_old:=OracleDataSet1SHOP.AsInteger;
                 fl_inv_old:=OracleDataSet1FLAGINV.AsInteger;;

                    sum2_i:=OracleDataSet1SUM2.AsFloat;
                    sum2:=OracleDataSet1SUM2.AsFloat;
                    sum3_i:=OracleDataSet1SUM3.AsFloat;
                    sum3:=OracleDataSet1SUM3.AsFloat;
                    sum4_i:=OracleDataSet1SUM4.AsFloat;
                    sum4:=OracleDataSet1SUM4.AsFloat;
                    sum5_i:=OracleDataSet1SUM5.AsFloat;
                    sum5:=OracleDataSet1SUM5.AsFloat;
                    sum22_i:=OracleDataSet1SUM22.AsFloat;
                    sum22:=OracleDataSet1SUM22.AsFloat;
                    sum33_i:=OracleDataSet1SUM33.AsFloat;
                    sum33:=OracleDataSet1SUM33.AsFloat;
                    sum44_i:=OracleDataSet1SUM44.AsFloat;
                    sum44:=OracleDataSet1SUM44.AsFloat;
                    sum55_i:=OracleDataSet1SUM55.AsFloat;
                    sum55:=OracleDataSet1SUM55.AsFloat;
                        k_st:=0;     ks_st:=0;

                 RUzpl:=0; UFzpl:=0; RAZzpl:=0; RUgpd:=0; UFgpd:=0; RAZgpd:=0;
                 RUbol:=0; UFbol:=0; RAZbol:=0; RUblfss:=0;  UFblfss:=0; RAZblfss:=0;

                  OracleQueryufzpl.SetVariable('emp',OracleDataSet1emp.AsInteger);
                  OracleQueryufzpl.SetVariable('yearp',yearp);
                  OracleQueryufzpl.SetVariable('monthz',OracleDataSet1monthz.AsInteger);
                  OracleQueryufzpl.SetVariable('shop',OracleDataSet1shop.AsInteger);
                  OracleQueryufzpl.SetVariable('expend',OracleDataSet1expend.AsInteger);
                  OracleQueryufzpl.SetVariable('FLAGINV',OracleDataSet1flaginv.AsInteger);

                   with OracleQueryufzpl do
                      try
                          Form1.StaticText1.Caption := 'ufzpl '+IntToStr(emp);
                       try
                          Execute;
                   except
                       on E:EOracleError do begin
                       ShowMessage(E.Message);
                       exit;
                       end;
                       end;
                       ufzpl:=Field(0);
                   except
                       on E:EOracleError do ShowMessage(E.Message);
                   end;
               Form1.StaticText1.Caption:='ufzpl '+IntToStr(emp);

               if OracleDataSet1flaginv.AsInteger<>0 then  RUzpl:=RoundEx(OracleDataSet1SUM2.AsFloat*kfzpli,-2)
                         else RUzpl:=RoundEx((OracleDataSet1SUM2.AsFloat*kfzpl),-2);



                  OracleQueryufbol.SetVariable('emp',OracleDataSet1emp.AsInteger);
                  OracleQueryufbol.SetVariable('yearp',yearp);
                  OracleQueryufbol.SetVariable('monthz',OracleDataSet1monthz.AsInteger);
                  OracleQueryufbol.SetVariable('shop',OracleDataSet1shop.AsInteger);
                  OracleQueryufbol.SetVariable('expend',OracleDataSet1expend.AsInteger);
                  OracleQueryufbol.SetVariable('FLAGINV',OracleDataSet1flaginv.AsInteger);

                   with OracleQueryufbol do
                      try
                          Form1.StaticText1.Caption := 'ufbol '+IntToStr(emp);
                       try
                          Execute;
                   except
                       on E:EOracleError do begin
                       ShowMessage(E.Message);
                       exit;
                       end;
                       end;
                       ufbol:=Field(0);
                   except
                       on E:EOracleError do ShowMessage(E.Message);
                   end;
               Form1.StaticText1.Caption:='ufbol '+IntToStr(emp);

               if OracleDataSet1flaginv.AsInteger<>0 then  RUbol:=RoundEx((OracleDataSet1SUM3.AsFloat*kfboli),-2)
                             else  RUbol:=RoundEx((OracleDataSet1SUM3.AsFloat*kfbol),-2);



                  OracleQueryufblfss.SetVariable('emp',OracleDataSet1emp.AsInteger);
                  OracleQueryufblfss.SetVariable('yearp',yearp);
                  OracleQueryufblfss.SetVariable('monthz',OracleDataSet1monthz.AsInteger);
                  OracleQueryufblfss.SetVariable('shop',OracleDataSet1shop.AsInteger);
                  OracleQueryufblfss.SetVariable('expend',OracleDataSet1expend.AsInteger);
                  OracleQueryufblfss.SetVariable('FLAGINV',OracleDataSet1flaginv.AsInteger);


                   with OracleQueryufblfss do
                      try
                          Form1.StaticText1.Caption := 'ufblfss '+IntToStr(emp);
                       try
                          Execute;
                   except
                       on E:EOracleError do begin
                       ShowMessage(E.Message);
                       exit;
                       end;
                       end;
                       ufblfss:=Field(0);
                   except
                       on E:EOracleError do ShowMessage(E.Message);
                   end;
               Form1.StaticText1.Caption:='ufblfss '+IntToStr(emp);
               if OracleDataSet1flaginv.AsInteger<>0 then RUblfss:=RoundEx((OracleDataSet1SUM4.AsFloat*kfblfssi),-2)
                             else RUblfss:=RoundEx((OracleDataSet1SUM4.AsFloat*kfblfss),-2);




                  OracleQueryufgpd.SetVariable('emp',OracleDataSet1emp.AsInteger);
                  OracleQueryufgpd.SetVariable('yearp',yearp);
                  OracleQueryufgpd.SetVariable('monthz',OracleDataSet1monthz.AsInteger);
                  OracleQueryufgpd.SetVariable('shop',OracleDataSet1shop.AsInteger);
                  OracleQueryufgpd.SetVariable('expend',OracleDataSet1expend.AsInteger);
                  OracleQueryufgpd.SetVariable('FLAGINV',OracleDataSet1flaginv.AsInteger);

                   with OracleQueryufgpd do
                      try
                          Form1.StaticText1.Caption := 'ufgpd '+IntToStr(emp);
                       try
                          Execute;
                   except
                       on E:EOracleError do begin
                       ShowMessage(E.Message);
                       exit;
                       end;
                       end;
                       ufgpd:=Field(0);
                   except
                       on E:EOracleError do ShowMessage(E.Message);
                   end;
               Form1.StaticText1.Caption:='ufgpd '+IntToStr(emp);
              RUgpd:=RoundEx((OracleDataSet1SUM5.AsFloat*kfgpd),-2);








         end
            else begin





                    sum2_i:=OracleDataSet1sum2.AsFloat+sum2_i;
                    sum3_i:=OracleDataSet1sum3.AsFloat+sum3_i;
                    sum4_i:=OracleDataSet1sum4.AsFloat+sum4_i;
                    sum5_i:=OracleDataSet1sum5.AsFloat+sum5_i;
                    sum22_i:=OracleDataSet1sum22.AsFloat+sum22_i;
                    sum33_i:=OracleDataSet1sum33.AsFloat+sum33_i;
                    sum44_i:=OracleDataSet1sum44.AsFloat+sum44_i;
                    sum55_i:=OracleDataSet1sum55.AsFloat+sum55_i;



                    //    if expendold<>expend
                if OracleDataSet1flaginv.AsInteger<>0 then  RUzpl:=RUzpl+RoundEx(OracleDataSet1sum2.AsFloat*kfzpli,-2)
                         else RUzpl:=RUzpl+RoundEx((OracleDataSet1sum2.AsFloat*kfzpl),-2);


                  if OracleDataSet1flaginv.AsInteger<>0 then  RUbol:=RUbol+RoundEx((OracleDataSet1sum3.AsFloat*kfboli),-2)
                             else  RUbol:=RUbol+RoundEx((OracleDataSet1sum3.AsFloat*kfbol),-2);


                if OracleDataSet1flaginv.AsInteger<>0 then RUblfss:=RUblfss+RoundEx((OracleDataSet1sum4.AsFloat*kfblfssi),-2)
                             else RUblfss:=RUblfss+RoundEx((OracleDataSet1sum4.AsFloat*kfblfss),-2);

                 RUgpd:=RUgpd+RoundEx((OracleDataSet1sum5.AsFloat*kfgpd),-2);

            end;



                   pech;

        // end;

                             Application.ProcessMessages;





      OracleDataSet1.Next;
      ////////////////////


     end; //    for i:=1 to OracleDataSet1.RecordCount do



           //последн€€ запись
        RAZzpl:=RUzpl-UFzpl;
                  RAZbol:=RUbol-UFbol;
                  RAZblfss:=RUblfss-UFblfss;
                  RAZgpd:=RUgpd-UFgpd;


                  if pr=0  //если указатель пустой то пишем в базу
           then  begin
              if RoundEx(Razzpl,-2)<>0
                 then  zap_baza(801,RoundEx(Razzpl,-2));//  пишем в базу   pay=801

              if RoundEx(Razbol,-2)<>0
                 then zap_baza(807,RoundEx(Razbol,-2)); // пишем в базу   pay=807

              if RoundEx(RAZblfss,-2)<>0
                 then zap_baza(809,RoundEx(RAZblfss,-2)); // пишем в базу   pay=809

              if RoundEx(RAZgpd,-2)<>0
                 then zap_baza(801,RoundEx(RAZgpd,-2)); // пишем в базу   pay=801


                end;  //      if pr=0


                pech_it;







           end   //if OracleDataSet1.RecordCount<>0 then begin
     else   pr_zap:=1; // showmessage('Ќет данных по запросу.');

    OracleDataSet1.Close;



    end;



 procedure  TDataModule2.zap_baza(pay:integer;raz:real);     //запись в базу
     begin

                   OracleQuery1ins.SetVariable('emp',emp_old);
                   OracleQuery1ins.SetVariable('pay',pay);
                   OracleQuery1ins.SetVariable('year',year);
                   OracleQuery1ins.SetVariable('month',month);
                   OracleQuery1ins.SetVariable('yearp',yearp);
                   OracleQuery1ins.SetVariable('monthz',monthz_old);
                   OracleQuery1ins.SetVariable('expend',expendold);
                   OracleQuery1ins.SetVariable('raz',raz);
                   OracleQuery1ins.SetVariable('flaginv',fl_inv_old);
                   OracleQuery1ins.SetVariable('shop',dept_old);

                    with OracleQuery1ins do
                      try
                          Form1.StaticText1.Caption := 'ins '+IntToStr(emp);
                       try
                          Execute;
                   except
                       on E:EOracleError do begin
                       ShowMessage(E.Message);
                       exit;
                       end;
                       end;

                   except
                       on E:EOracleError do ShowMessage(E.Message);
                                          end;

                              OracleSession1.Commit;
  end; //procedure  TDataModule2.zap_baza;     //запись в базу




  procedure  TDataModule2.pech;


  begin





   if pse=0 then
   begin
           //Inc(k_st);
        if (emp_old<>OracleDataSet1emp.AsInteger )  then    begin
           emp_old:=OracleDataSet1emp.AsInteger;  k:=k+1;

                                          end;


    // if ( (DataModule2.OracleDataSet1sum2.AsFloat<>0) or (DataModule2.OracleDataSet1SUM22.AsFloat<>0)
    // or (RUzpl<>0) or  (Ufzpl<>0)) then begin
   //if ( (SUM2<>0) or (Sum22<>0) or (RUzpl<>0) or  (Ufzpl<>0)) then begin

     if tip_nach_i=1
       then  begin
       if dschdate<>'' then  Sheet.Cells[k,12]:=dschdate;
        Sheet.Cells[k,1]:=OracleDataSet1emp.AsInteger;
    //      k:=k+1;
   //  writeln(myfile,inttostr(emp)+' дата увольнени€ '+dschdate);
   //  writeln(myfile,inttostr(monthz)+' 1 '+floattostr(SUMZPL)+' '+floattostr(mas_kf[1,monthz])+' '+floattostr(RUzpl)+' '+floattostr(STAX_ZPL)+' '+floattostr(RoundEx(RAZzpl,-2)));
        Sheet.Cells[k,2]:=DataModule2.OracleDataSet1MONTHZ.Value; Sheet.Cells[k,3]:='1';
        Sheet.Cells[k,4]:=DataModule2.OracleDataSet1EXPEND.Value;
        Sheet.Cells[k,5]:=DataModule2.OracleDataSet1SHOP.Value;
        Sheet.Cells[k,6]:=DataModule2.OracleDataSet1FLAGINV.Value;
        Sheet.Cells[k,7]:=DataModule2.OracleDataSet1sum2.Value;
        Sheet.Cells[k,8]:=DataModule2.OracleDataSet1sum22.Value  ;

            if OracleDataSet1flaginv.AsInteger<>0 then  Sheet.Cells[k,9]:=RoundEx(OracleDataSet1sum2.AsFloat*kfzpli,-2)
                         else Sheet.Cells[k,9]:=RoundEx((OracleDataSet1sum2.AsFloat*kfzpl),-2);


       // Sheet.Cells[k,10]:=ufzpl; //Sheet.Cells[k,11]:=RUzpl-UFzpl;;
         Sheet.Cells[k,13]:=DataModule2.OracleDataSet1MONTHYEAR.Value;
        if DataModule2.OracleDataSet1flaginv.AsInteger<>0 then  Sheet.Cells[k,14]:=kfzpli
                         else Sheet.Cells[k,14]:=kfzpl;

             Inc(k_st);

         k:=k+1;       // new_page;
         end;


     if  tip_nach_i=2 //( (DataModule2.OracleDataSet1sum3.AsFloat<>0) or (DataModule2.OracleDataSet1SUM33.AsFloat<>0)
         //    or (RUbol<>0) or  (Ufbol<>0))
             then begin
         Sheet.Cells[k,1]:=DataModule2.OracleDataSet1emp.Value;
        Sheet.Cells[k,2]:=DataModule2.OracleDataSet1monthz.Value; Sheet.Cells[k,3]:='2';
       Sheet.Cells[k,4]:=DataModule2.OracleDataSet1expend.Value;
       Sheet.Cells[k,5]:= DataModule2.OracleDataSet1shop.Value;
       Sheet.Cells[k,6]:=DataModule2.OracleDataSet1flaginv.Value;
        Sheet.Cells[k,7]:=DataModule2.OracleDataSet1sum3.Value;
        Sheet.Cells[k,8]:=DataModule2.OracleDataSet1sum33.Value  ;


                  if DataModule2.OracleDataSet1flaginv.AsInteger<>0 then  Sheet.Cells[k,9]:=RoundEx((OracleDataSet1sum3.AsFloat*kfboli),-2)
                             else   Sheet.Cells[k,9]:=RoundEx((OracleDataSet1sum3.AsFloat*kfbol),-2);



       // Sheet.Cells[k,10]:=ufbol; //Sheet.Cells[k,11]:=RUbol-UFbol;



        Sheet.Cells[k,13]:=DataModule2.OracleDataSet1MONTHYEAR.Value;

       if DataModule2.OracleDataSet1flaginv.AsInteger<>0 then  Sheet.Cells[k,14]:=kfboli
                         else Sheet.Cells[k,14]:=kfbol;

        k:=k+1;        Inc(k_st);   //  new_page;

        //        writeln(myfile,inttostr(monthz)+' 2 '+floattostr(SUMBOL)+' '+floattostr(mas_kf[1,monthz])+' '+floattostr(RUbol)+' '+floattostr(STAX_BL)+' '+floattostr(RoundEx(RAZbol,-2)));
                      end;


     if  tip_nach_i=3 {( (DataModule2.OracleDataSet1SUM4.AsFloat<>0) or
           (DataModule2.OracleDataSet1Sum44.AsFloat<>0) or
           (RUblfss<>0) or (Ufbol<>0) ) }

           then  begin
        Sheet.Cells[k,1]:=DataModule2.OracleDataSet1emp.Value;
        Sheet.Cells[k,2]:=DataModule2.OracleDataSet1monthz.Value; Sheet.Cells[k,3]:='3';
        Sheet.Cells[k,4]:=DataModule2.OracleDataSet1expend.Value;
        Sheet.Cells[k,5]:= DataModule2.OracleDataSet1shop.Value;
        Sheet.Cells[k,6]:=DataModule2.OracleDataSet1flaginv.Value;
        Sheet.Cells[k,7]:=DataModule2.OracleDataSet1sum4.Value;
        Sheet.Cells[k,8]:=DataModule2.OracleDataSet1sum44.Value  ;
          if DataModule2.OracleDataSet1flaginv.AsInteger<>0 then Sheet.Cells[k,9]:=RoundEx((OracleDataSet1sum4.AsFloat*kfblfssi),-2)
                             else Sheet.Cells[k,9]:=RoundEx((OracleDataSet1sum4.AsFloat*kfblfss),-2);


       // Sheet.Cells[k,10]:=ufblfss; //Sheet.Cells[k,11]:=RUblfss-UFblfss;
       Sheet.Cells[k,13]:=DataModule2.OracleDataSet1MONTHYEAR.Value;
        if DataModule2.OracleDataSet1flaginv.AsInteger<>0 then  Sheet.Cells[k,14]:=kfblfssi
                         else   Sheet.Cells[k,14]:=kfblfss;

        k:=k+1;     Inc(k_st);        //  new_page;

//                writeln(myfile,inttostr(monthz)+' 3 '+floattostr(SUMBLFSS)+' '+floattostr(mas_kf[1,monthz])+' '+floattostr(RUblfss)+' '+floattostr(STAX_BLFSS)+' '+floattostr(RoundEx(RAZblfss,-2)));
                                                                        END;
     if tip_nach_i=4 {( (DataModule2.OracleDataSet1SUM5.AsFloat<>0)
           or (DataModule2.OracleDataSet1Sum55.AsFloat<>0) or (RUgpd<>0) or  (Ufgpd<>0)) }
           then begin
         Sheet.Cells[k,1]:=DataModule2.OracleDataSet1emp.Value;
        Sheet.Cells[k,2]:=DataModule2.OracleDataSet1monthz.Value;
        Sheet.Cells[k,3]:='4';
        Sheet.Cells[k,4]:=DataModule2.OracleDataSet1expend.Value;
        Sheet.Cells[k,5]:= DataModule2.OracleDataSet1shop.Value;
        Sheet.Cells[k,6]:=DataModule2.OracleDataSet1flaginv.Value;
        Sheet.Cells[k,7]:=DataModule2.OracleDataSet1sum5.Value;
        Sheet.Cells[k,8]:=DataModule2.OracleDataSet1sum55.Value  ;

        Sheet.Cells[k,9]:=RoundEx((OracleDataSet1sum5.AsFloat*kfgpd),-2);
        //Sheet.Cells[k,10]:=ufgpd; //Sheet.Cells[k,11]:=RUgpd-UFgpd;
         Sheet.Cells[k,13]:=DataModule2.OracleDataSet1MONTHYEAR.Value;
        Sheet.Cells[k,14]:=kfgpd; k:=k+1;      Inc(k_st);  // new_page;
                                                                         END;
                  end   // if pse=0 then
       else   begin
          if (emp_old<>OracleDataSet1emp.AsInteger )  then    begin
           emp_old:=OracleDataSet1emp.AsInteger;  ks:=ks+1;  // new_page;
                                    end;



       if tip_nach_i=1 {( (DataModule2.OracleDataSet1sum2.AsFloat<>0) or (DataModule2.OracleDataSet1SUM22.AsFloat<>0)
     or (RUzpl<>0) or  (Ufzpl<>0)) }then begin

            Sheets.Cells[ks,1]:=OracleDataSet1emp.AsInteger;

       if dschdate<>'' then  Sheets.Cells[ks,12]:=dschdate;
        Sheets.Cells[ks,2]:=DataModule2.OracleDataSet1monthz.Value; Sheets.Cells[ks,3]:='1';
        Sheets.Cells[ks,4]:=DataModule2.OracleDataSet1expend.Value;
        Sheets.Cells[ks,5]:= DataModule2.OracleDataSet1shop.Value;
        Sheets.Cells[ks,6]:=DataModule2.OracleDataSet1flaginv.Value;
        Sheets.Cells[ks,7]:=DataModule2.OracleDataSet1sum2.Value;
        Sheets.Cells[ks,8]:=DataModule2.OracleDataSet1sum22.Value  ;

        if OracleDataSet1flaginv.AsInteger<>0 then  Sheets.Cells[ks,9]:=RoundEx(OracleDataSet1sum2.AsFloat*kfzpli,-2)
                         else Sheets.Cells[ks,9]:=RoundEx((OracleDataSet1sum2.AsFloat*kfzpl),-2);

       // Sheets.Cells[ks,9]:=Ruzpl;
      //  Sheets.Cells[ks,10]:=ufzpl; //Sheets.Cells[ks,11]:=razzpl;
         Sheets.Cells[ks,13]:=DataModule2.OracleDataSet1MONTHYEAR.Value;
                if DataModule2.OracleDataSet1flaginv.AsInteger<>0 then  Sheets.Cells[ks,14]:=kfzpli
                         else Sheets.Cells[ks,14]:=kfzpl;
           ks:=ks+1;  Inc(ks_st);  //   new_page;
           end;


     if tip_nach_i=2 {( (DataModule2.OracleDataSet1SUM4.AsFloat<>0) or (DataModule2.OracleDataSet1Sum44.AsFloat<>0)
            or (RUblfss<>0) or (Ufbol<>0) )} then  begin
        Sheets.Cells[ks,1]:=DataModule2.OracleDataSet1emp.Value;
        Sheets.Cells[ks,2]:=DataModule2.OracleDataSet1monthz.Value; Sheets.Cells[ks,3]:='2';
        Sheets.Cells[ks,4]:=DataModule2.OracleDataSet1expend.Value;
        Sheets.Cells[ks,5]:= DataModule2.OracleDataSet1shop.Value;
        Sheets.Cells[ks,6]:=DataModule2.OracleDataSet1flaginv.Value;
        Sheets.Cells[ks,7]:=DataModule2.OracleDataSet1sum3.Value;
        Sheets.Cells[ks,8]:=DataModule2.OracleDataSet1sum33.Value  ;

           if DataModule2.OracleDataSet1flaginv.AsInteger<>0 then  Sheets.Cells[ks,9]:=RoundEx((OracleDataSet1sum3.AsFloat*kfboli),-2)
                             else   Sheets.Cells[ks,9]:=RoundEx((OracleDataSet1sum3.AsFloat*kfbol),-2);



        //Sheets.Cells[ks,10]:=ufbol; //Sheet.Cells[k,11]:=RUbol-UFbol;



        Sheets.Cells[ks,13]:=DataModule2.OracleDataSet1MONTHYEAR.Value;
        if DataModule2.OracleDataSet1flaginv.AsInteger<>0 then  Sheets.Cells[ks,14]:=kfboli
                       else   Sheets.Cells[ks,14]:=kfbol;
         ks:=ks+1;   Inc(ks_st);     // new_page;

//                writeln(myfile,inttostr(monthz)+' 3 '+floattostr(SUMBLFSS)+' '+floattostr(mas_kf[1,monthz])+' '+floattostr(RUblfss)+' '+floattostr(STAX_BLFSS)+' '+floattostr(RoundEx(RAZblfss,-2)));
                                                                        END;
     if  tip_nach_i=3 {( (DataModule2.OracleDataSet1SUM5.AsFloat<>0)
           or (DataModule2.OracleDataSet1Sum55.AsFloat<>0) or (RUgpd<>0) or  (Ufgpd<>0))} then begin
         Sheets.Cells[ks,1]:=DataModule2.OracleDataSet1emp.Value;
        Sheets.Cells[ks,2]:=DataModule2.OracleDataSet1monthz.Value; Sheets.Cells[ks,3]:='3';
        Sheets.Cells[ks,4]:=DataModule2.OracleDataSet1expend.Value;
        Sheets.Cells[ks,5]:= DataModule2.OracleDataSet1shop.Value;
        Sheets.Cells[ks,6]:=DataModule2.OracleDataSet1flaginv.Value;
        Sheets.Cells[ks,7]:=DataModule2.OracleDataSet1sum4.Value;
        Sheets.Cells[ks,8]:=DataModule2.OracleDataSet1sum44.Value  ;
       if DataModule2.OracleDataSet1flaginv.AsInteger<>0 then Sheets.Cells[ks,9]:=RoundEx((OracleDataSet1sum4.AsFloat*kfblfssi),-2)
                             else Sheets.Cells[ks,9]:=RoundEx((OracleDataSet1sum4.AsFloat*kfblfss),-2);


     //   Sheets.Cells[ks,10]:=ufblfss;
        //Sheets.Cells[ks,9]:=rugpd;

          Sheets.Cells[ks,13]:=DataModule2.OracleDataSet1MONTHYEAR.Value;
        if DataModule2.OracleDataSet1flaginv.AsInteger<>0 then  Sheets.Cells[ks,14]:=kfblfssi
                         else   Sheets.Cells[ks,14]:=kfblfss;


        ks:=ks+1;    Inc(ks_st);       //new_page;




                                       end;



         if tip_nach_i=4 {( (DataModule2.OracleDataSet1SUM5.AsFloat<>0)
           or (DataModule2.OracleDataSet1Sum55.AsFloat<>0) or (RUgpd<>0) or  (Ufgpd<>0)) }
           then begin
         Sheets.Cells[ks,1]:=DataModule2.OracleDataSet1emp.Value;
        Sheets.Cells[ks,2]:=DataModule2.OracleDataSet1monthz.Value;
        Sheets.Cells[ks,3]:='4';
        Sheets.Cells[ks,4]:=DataModule2.OracleDataSet1expend.Value;
        Sheets.Cells[ks,5]:= DataModule2.OracleDataSet1shop.Value;
        Sheets.Cells[ks,6]:=DataModule2.OracleDataSet1flaginv.Value;
        Sheets.Cells[ks,7]:=DataModule2.OracleDataSet1sum5.Value;
        Sheets.Cells[ks,8]:=DataModule2.OracleDataSet1sum55.Value  ;

        Sheets.Cells[ks,9]:=RoundEx((OracleDataSet1sum5.AsFloat*kfgpd),-2);
       // Sheets.Cells[ks,10]:=ufgpd; //Sheet.Cells[k,11]:=RUgpd-UFgpd;
         Sheets.Cells[ks,13]:=DataModule2.OracleDataSet1MONTHYEAR.Value;
        Sheets.Cells[ks,14]:=kfgpd; ks:=ks+1;      Inc(ks_st);    //  new_page;
                                                                         END;






                  end;  // if pse=0 else

               //   new_page

     //если записей больше чем макс кол-во строк в Excel
{      if (k>=65000) and ( l<>( i div 65000 ))then begin
                 l:=1+(i div 65000);
                 showmessage(inttistr(l));
                 XLApp.Workbooks[1].Worksheets[l+1].Name:='801_'+inttostr(l);
                 Colum:=XLApp.Workbooks[1].WorkSheets['801_'+inttostr(l)].Columns;
                 Row:=XLApp.Workbooks[1].WorkSheets['801_'+inttostr(l)].Rows;
                 Sheet:=XLApp.Workbooks[1].WorkSheets['801_'+inttostr(l)];
                 k:=1;
                      end;
 }






  end;

   procedure  TDataModule2.new_page;
  begin
   if (k>=60000) then begin
                 l:=l+1;

                 XLApp.Workbooks[1].Worksheets[l+1].Name:='801_'+inttostr(l);
                 Colum:=XLApp.Workbooks[1].WorkSheets['801_'+inttostr(l)].Columns;
                 Row:=XLApp.Workbooks[1].WorkSheets['801_'+inttostr(l)].Rows;
                 Sheet:=XLApp.Workbooks[1].WorkSheets['801_'+inttostr(l)];
                 k:=3;
                  Sheet.select;
                      end;



      if (ks>=60000) then begin
           ls:=ls+1;
           XLApps.Workbooks[1].Worksheets[ls+1].Name:='801s_'+inttostr(ls);
           Colums:=XLApps.Workbooks[1].WorkSheets['801s_'+inttostr(ls)].Columns;
           Rows:=XLApps.Workbooks[1].WorkSheets['801s_'+inttostr(ls)].Rows;
           Sheets:=XLApps.Workbooks[1].WorkSheets['801s_'+inttostr(ls)];
           ks:=3;
           Sheets.select;
                     end;

    {
            if (k>=60000) and ( l<>( i div 60000 ))then begin
                 l:=1+(i div 60000);
                // showmessage(inttostr(l));
                 XLApp.Workbooks[1].Worksheets[l+1].Name:='801_'+inttostr(l);
                 Colum:=XLApp.Workbooks[1].WorkSheets['801_'+inttostr(l)].Columns;
                 Row:=XLApp.Workbooks[1].WorkSheets['801_'+inttostr(l)].Rows;
                 Sheet:=XLApp.Workbooks[1].WorkSheets['801_'+inttostr(l)];
                 k:=1;
                 Sheet.select;
                      end;




                             if (ks>=60000) and ( ls<>( i div 60000 ))then begin
                 ls:=1+(i div 60000);
                // showmessage(inttostr(l));
                 XLApp.Workbooks[1].Worksheets[l+1].Name:='801_'+inttostr(l);
                 Colum:=XLApp.Workbooks[1].WorkSheets['801_'+inttostr(l)].Columns;
                 Row:=XLApp.Workbooks[1].WorkSheets['801_'+inttostr(l)].Rows;
                 Sheet:=XLApp.Workbooks[1].WorkSheets['801_'+inttostr(l)];
                 k:=1;
                 Sheet.select;
                      end;


   {  if (k>=60000) then begin
                 l:=l+1;

                 XLApp.Workbooks[1].Worksheets[l+1].Name:='801_'+inttostr(l);
                 Colum:=XLApp.Workbooks[1].WorkSheets['801_'+inttostr(l)].Columns;
                 Row:=XLApp.Workbooks[1].WorkSheets['801_'+inttostr(l)].Rows;
                 Sheet:=XLApp.Workbooks[1].WorkSheets['801_'+inttostr(l)];
                 k:=1;         Sheet.select;
                      end;



      if (ks>=60000) then begin
           ls:=ls+1;
           XLApps.Workbooks[1].Worksheets[ls+1].Name:='801s_'+inttostr(ls);
           Colums:=XLApps.Workbooks[1].WorkSheets['801s_'+inttostr(ls)].Columns;
           Rows:=XLApps.Workbooks[1].WorkSheets['801s_'+inttostr(ls)].Rows;
           Sheets:=XLApps.Workbooks[1].WorkSheets['801s_'+inttostr(ls)];
           ks:=10;     Sheets.select;
                     end;


    }

  end;



   procedure  TDataModule2.pech_it;
   begin



      if pse=0 then   begin
       if (k_st=1)  then  k:=k-1; //печать итоговой строки


                       if  tip_nach_old=1
                       //(sum2_i<>0) or (sum22_i<>0) or (RUzpl<>0)  or (Ufzpl<>0)
                            then begin

                      Sheet.Rows[k].Font.Bold := True;
                      Sheet.Cells[k,1]:=emp_old;
                       Sheet.Cells[k,2]:=monthz_old;
                       Sheet.Cells[k,3]:=tip_nach_old;

                      Sheet.Cells[k,4]:=expendold;
                       Sheet.Cells[k,5]:=dept_old;
                       Sheet.Cells[k,6]:=fl_inv_old;

                      Sheet.Cells[k,7]:=sum2_i;
                      Sheet.Cells[k,8]:=sum22_i  ;
                      Sheet.Cells[k,9]:=Ruzpl;
                      Sheet.Cells[k,10]:=ufzpl;
                      Sheet.Cells[k,11]:=razzpl;
                      k_st:=0;
                      inc(k);   //new_page;
                         end;
                      if   tip_nach_old=2 //1(sum3_i<>0) or (sum33_i<>0) or (RUbol<>0)  or (Ufbol<>0)
                            then begin
                              Sheet.Rows[k].Font.Bold := True;
                               Sheet.Cells[k,1]:=emp_old;
                       Sheet.Cells[k,2]:=monthz_old;
                       Sheet.Cells[k,3]:=tip_nach_old;

                              Sheet.Cells[k,4]:=expendold;
                             Sheet.Cells[k,5]:=dept_old;
                             Sheet.Cells[k,6]:=fl_inv_old;

                              Sheet.Cells[k,7]:=sum3_i;
                              Sheet.Cells[k,8]:=sum33_i  ;
                              Sheet.Cells[k,9]:=RUbol;
                              Sheet.Cells[k,10]:=Ufbol;
                              sheet.Cells[k,11]:=razbol;
                                   inc(k);  k_st:=0;   // new_page;

                            end;

                          if   tip_nach_old=3 //( (SUM4_i<>0) or (Sum44_i<>0) or (RUblfss<>0) or (RAZblfss<>0)or (UFblfss<>0) )
                           then  begin
                               Sheet.Rows[k].Font.Bold := True;
                              Sheet.Cells[k,1]:=emp_old;
                       Sheet.Cells[k,2]:=monthz_old;
                       Sheet.Cells[k,3]:=tip_nach_old;
                       Sheet.Cells[k,5]:=dept_old;
                       Sheet.Cells[k,6]:=fl_inv_old;



                              Sheet.Cells[k,4]:=expendold;
                              Sheet.Cells[k,7]:=sum4_i;
                              Sheet.Cells[k,8]:=sum44_i  ;
                              Sheet.Cells[k,9]:=RUblfss;
                             Sheet.Cells[k,10]:=UFblfss;
                              Sheet.Cells[k,11]:= razblfss;

                                   inc(k); k_st:=0;   // new_page;
                                 end;

                                if  tip_nach_old=4 //( (SUM5_i<>0) or (Sum55_i<>0) or (RUgpd<>0) or (RAZgpd<>0) or (Ufgpd<>0))
                                then begin
                                   Sheet.Rows[k].Font.Bold := True;
                               Sheet.Cells[k,1]:=emp_old;
                       Sheet.Cells[k,2]:=monthz_old;
                       Sheet.Cells[k,3]:=tip_nach_old;
                       Sheet.Cells[k,5]:=dept_old;
                       Sheet.Cells[k,6]:=fl_inv_old;


                              Sheet.Cells[k,4]:=expendold;
                              Sheet.Cells[k,7]:=sum5_i;
                              Sheet.Cells[k,8]:=sum55_i  ;
                              Sheet.Cells[k,9]:=RUgpd;
                             Sheet.Cells[k,10]:=Ufgpd;
                              Sheet.Cells[k,11]:= RAZgpd;
                               inc(k); k_st:=0;    // new_page;

                                end;


                           end

               else    begin
                     if (ks_st=1)  then  ks:=ks-1;

                     if  tip_nach_old=1 //(sum2_i<>0) or (sum22_i<>0) or (RUzpl<>0)  or (Ufzpl<>0)
                            then begin
                      Sheets.Rows[ks].Font.Bold := True;
                       Sheets.Cells[ks,1]:=emp_old;
                       Sheets.Cells[ks,2]:=monthz_old;
                       Sheets.Cells[ks,3]:=tip_nach_old;
                       Sheets.Cells[ks,5]:=dept_old;
                       Sheets.Cells[ks,6]:=fl_inv_old;


                      Sheets.Cells[ks,4]:=expendold;
                      Sheets.Cells[ks,7]:=sum2_i;
                      Sheets.Cells[ks,8]:=sum22_i  ;
                      Sheets.Cells[ks,9]:=Ruzpl;
                      Sheets.Cells[ks,10]:=ufzpl;
                      Sheets.Cells[ks,11]:=razzpl;
                      ks_st:=0;
                      inc(ks);  // new_page;
                      end;

                      if  tip_nach_old=2 //1(sum3_i<>0) or (sum33_i<>0) or (RUbol<>0)  or (Ufbol<>0)
                            then begin
                              Sheets.Rows[ks].Font.Bold := True;
                                                   Sheets.Cells[ks,1]:=emp_old;
                       Sheets.Cells[ks,2]:=monthz_old;
                       Sheets.Cells[ks,3]:=tip_nach_old;
                        Sheets.Cells[ks,5]:=dept_old;
                       Sheets.Cells[ks,6]:=fl_inv_old;


                              Sheets.Cells[ks,4]:=expendold;
                              Sheets.Cells[ks,7]:=sum3_i;
                              Sheets.Cells[ks,8]:=sum33_i  ;
                              SheetS.Cells[ks,9]:=RUbol;
                              Sheets.Cells[ks,10]:=Ufbol;
                              Sheets.Cells[ks,11]:=razbol;
                                   inc(ks);    ks_st:=0;  //new_page;

                            end;

                          if  tip_nach_old=3 //( (SUM4_i<>0) or (Sum44_i<>0) or (RUblfss<>0) or (RAZblfss<>0)or (UFblfss<>0) )
                           then  begin
                               Sheets.Rows[ks].Font.Bold := True;
                                                  Sheets.Cells[ks,1]:=emp_old;
                       Sheets.Cells[ks,2]:=monthz_old;
                       Sheets.Cells[ks,3]:=tip_nach_old;
                       Sheets.Cells[ks,5]:=dept_old;
                       Sheets.Cells[ks,6]:=fl_inv_old;


                              Sheets.Cells[ks,4]:=expendold;
                              Sheets.Cells[ks,7]:=sum4_i;
                              Sheets.Cells[ks,8]:=sum44_i  ;
                              Sheets.Cells[ks,9]:=RUblfss;
                             Sheets.Cells[ks,10]:=UFblfss;
                              Sheets.Cells[ks,11]:= razblfss;
                                      ks_st:=0; // new_page;
                                   inc(ks);
                                 end;

                                if   tip_nach_old=4 //( (SUM5_i<>0) or (Sum55_i<>0) or (RUgpd<>0) or (RAZgpd<>0) or (Ufgpd<>0))
                                then begin
                                   Sheets.Rows[ks].Font.Bold := True;
                                     Sheets.Cells[ks,1]:=emp_old;
                       Sheets.Cells[ks,2]:=monthz_old;
                       Sheets.Cells[ks,3]:=tip_nach_old;
                              Sheets.Cells[ks,4]:=expendold;
                              Sheets.Cells[ks,5]:=dept_old;
                       Sheets.Cells[ks,6]:=fl_inv_old;

                              Sheets.Cells[ks,7]:=sum5_i;
                              Sheets.Cells[ks,8]:=sum55_i  ;
                              Sheets.Cells[ks,9]:=RUgpd;
                             Sheets.Cells[ks,10]:=Ufgpd;
                              Sheets.Cells[ks,11]:= RAZgpd;
                               ks_st:=0;  //     new_page;
                                   inc(ks);
                                end;


                           end;


           sum2_i:=0; sum22_i:=0;
            sum3_i:=0; sum33_i:=0;
             sum4_i:=0; sum44_i:=0;
              sum5_i:=0; sum55_i:=0;
          RUzpl:=0; UFzpl:=0; RAZzpl:=0; RUgpd:=0; UFgpd:=0; RAZgpd:=0;
    RUbol:=0; UFbol:=0; RAZbol:=0; RUblfss:=0;  UFblfss:=0; RAZblfss:=0;


  end;

end.
