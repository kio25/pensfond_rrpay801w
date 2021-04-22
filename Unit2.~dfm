object DataModule2: TDataModule2
  OldCreateOrder = False
  Left = 351
  Top = 509
  Height = 379
  Width = 776
  object OracleLogon1: TOracleLogon
    Session = OracleSession1
    Options = [ldAuto, ldDatabase, ldDatabaseList, ldLogonHistory, ldConnectAs]
    HistoryIniFile = 'c:\history.ini'
    Left = 696
    Top = 24
  end
  object OracleSession1: TOracleSession
    RollbackOnDisconnect = True
    Left = 696
    Top = 88
  end
  object OracleDataSet1: TOracleDataSet
    SQL.Strings = (
      
        '  SELECT EMP# emp, YEAR yearz, MONTH monthz, NVL(SHOP,0) shop, N' +
        'VL(EXPEND#,0)expend ,'
      
        '              NVL(SUM2,0) sum2,NVL(SUM3,0) sum3, NVL(SUM4,0) sum' +
        '4 ,NVL(SUM5,0) sum5,'
      
        '              NVL(SUM22,0) sum22, NVL(SUM33,0) sum33, NVL(SUM44,' +
        '0)sum44, NVL(SUM55,0)sum55,'
      
        '              NVL(FLAGINV,0) flaginv ,NVL(FLAGCORR,0) flagcorr,n' +
        'vl(monthyear,0) monthyear'
      '                 FROM PAYRECJUR'
      '                 WHERE YEAR=:yearp'
      '                       AND MONTH BETWEEN 1 AND :monthp'
      '                       AND EMP# BETWEEN  :empmin AND :empmax'
      
        '                 ORDER BY EMP#,MONTH,NVL(EXPEND#,0),nvl(monthyea' +
        'r,0),NVL(FLAGINV,0)'
      ''
      ''
      '/*'
      ''
      ''
      
        '         SELECT p.EMP# emp, p.YEAR yearz, p.MONTH monthz, NVL(p.' +
        'SHOP,0) shop, NVL(p.EXPEND#,0)expend ,'
      
        '              NVL(p.SUM2,0) sum2,NVL(p.SUM3,0) sum3, NVL(p.SUM4,' +
        '0) sum4 ,NVL(p.SUM5,0) sum5,'
      
        '              NVL(p.SUM22,0) sum22, NVL(p.SUM33,0) sum33, NVL(p.' +
        'SUM44,0)sum44, NVL(p.SUM55,0)sum55,'
      
        '              NVL(p.FLAGINV,0) flaginv ,NVL(p.FLAGCORR,0) flagco' +
        'rr,nvl(p.monthyear,0) monthyear,'
      '            '
      '               r.tip'
      '                 FROM PAYRECJUR p,'
      
        '                 (select distinct decode(pay#,800,1,801,1,802,1,' +
        '806,2,807,2,808,3,809,3,804,4,805,4) tip ,'
      
        '                    emp#,expend#,factmon,rshop,FLAGINV  from rec' +
        'alcjur'
      '                   where year=:yearp) r'
      '                 '
      '                 '
      '                 WHERE p.YEAR=:yearp'
      '                       AND p.MONTH BETWEEN 1 AND :monthp'
      '                       AND p.EMP# BETWEEN  :empmin AND :empmax'
      ' and p.emp#=r.emp#  and p.FLAGINV=r.FLAGINV '
      '                       and p.expend#=r.expend#'
      '                       and p.shop=r.rshop'
      '                       and r.factmon=:yearp*100+p.month '
      '  ORDER BY 1,3,17,5,16,14'
      '*/')
    Variables.Data = {
      0300000004000000070000003A454D504D494E03000000000000000000000007
      0000003A454D504D4158030000000000000000000000060000003A5945415250
      050000000000000000000000070000003A4D4F4E544850050000000000000000
      000000}
    QBEDefinition.QBEFieldDefs = {
      040000001000000003000000454D5001000000000005000000594541525A0100
      00000000060000004D4F4E54485A0100000000000400000053484F5001000000
      000006000000455850454E440100000000000400000053554D32010000000000
      0400000053554D330100000000000400000053554D3401000000000004000000
      53554D350100000000000500000053554D32320100000000000500000053554D
      33330100000000000500000053554D34340100000000000500000053554D3535
      01000000000007000000464C4147494E5601000000000008000000464C414743
      4F5252010000000000090000004D4F4E544859454152010000000000}
    Cursor = crSQLWait
    Session = OracleSession1
    Left = 48
    Top = 16
    object OracleDataSet1EMP: TIntegerField
      FieldName = 'EMP'
      Required = True
    end
    object OracleDataSet1YEARZ: TIntegerField
      FieldName = 'YEARZ'
      Required = True
    end
    object OracleDataSet1MONTHZ: TIntegerField
      FieldName = 'MONTHZ'
      Required = True
    end
    object OracleDataSet1SHOP: TFloatField
      FieldName = 'SHOP'
    end
    object OracleDataSet1EXPEND: TFloatField
      FieldName = 'EXPEND'
    end
    object OracleDataSet1SUM2: TFloatField
      FieldName = 'SUM2'
    end
    object OracleDataSet1SUM3: TFloatField
      FieldName = 'SUM3'
    end
    object OracleDataSet1SUM4: TFloatField
      FieldName = 'SUM4'
    end
    object OracleDataSet1SUM5: TFloatField
      FieldName = 'SUM5'
    end
    object OracleDataSet1SUM22: TFloatField
      FieldName = 'SUM22'
    end
    object OracleDataSet1SUM33: TFloatField
      FieldName = 'SUM33'
    end
    object OracleDataSet1SUM44: TFloatField
      FieldName = 'SUM44'
    end
    object OracleDataSet1SUM55: TFloatField
      FieldName = 'SUM55'
    end
    object OracleDataSet1FLAGINV: TFloatField
      FieldName = 'FLAGINV'
    end
    object OracleDataSet1FLAGCORR: TFloatField
      FieldName = 'FLAGCORR'
    end
    object OracleDataSet1MONTHYEAR: TFloatField
      FieldName = 'MONTHYEAR'
    end
  end
  object OracleQuery1del: TOracleQuery
    SQL.Strings = (
      'DELETE  FROM RECALCJUR'
      '        WHERE YEAR = :year  AND'
      '              MONTH= :month AND'
      '              PAY# IN (801,805,807,809) AND '
      '              EMP# BETWEEN  :empmin AND :empmax')
    Session = OracleSession1
    Variables.Data = {
      0300000004000000070000003A454D504D494E03000000000000000000000007
      0000003A454D504D4158030000000000000000000000060000003A4D4F4E5448
      030000000000000000000000050000003A594541520300000000000000000000
      00}
    Cursor = crSQLWait
    Left = 160
    Top = 16
  end
  object OracleQuery1ins: TOracleQuery
    SQL.Strings = (
      ' INSERT INTO RECALCJUR'
      
        '             VALUES (:emp, :pay, :year, :month, :yearp*100+:mont' +
        'hz, :expend, :RAZ, '
      
        '                    :flaginv, to_char(SYSDATE,'#39'DD/MM/YYYY HH24:M' +
        'I '#39') || USER, :shop)'
      '       ')
    Session = OracleSession1
    Variables.Data = {
      030000000A000000040000003A454D5003000000000000000000000004000000
      3A504159030000000000000000000000050000003A5945415203000000000000
      0000000000060000003A4D4F4E5448030000000000000000000000070000003A
      455850454E44030000000000000000000000040000003A52415A040000000000
      000000000000080000003A464C4147494E560300000000000000000000000500
      00003A53484F50030000000000000000000000060000003A5945415250030000
      000000000000000000070000003A4D4F4E54485A030000000000000000000000}
    Cursor = crSQLWait
    Left = 72
    Top = 168
  end
  object OracleDataSetkoef: TOracleDataSet
    SQL.Strings = (
      'SELECT NVL(NVAL,0) koef,INDATE '
      #9' FROM   SYSINDEX                                     '
      #9' WHERE  IND#   = :kf    AND                         '
      #9'        ROWNUM = 1       AND                         '
      #9'        INDATE = (SELECT MAX(INDATE)                 '
      #9#9'        FROM   SYSINDEX                      '
      #9#9'        WHERE  IND# = :kf AND               '
      #9#9#9'     INDATE <= :date1)               '
      '')
    Variables.Data = {
      0300000002000000030000003A4B46030000000000000000000000060000003A
      44415445310C0000000000000000000000}
    QBEDefinition.QBEFieldDefs = {
      0400000002000000040000004B4F454601000000000006000000494E44415445
      010000000000}
    Cursor = crSQLWait
    Session = OracleSession1
    Left = 64
    Top = 88
    object OracleDataSetkoefKOEF: TFloatField
      FieldName = 'KOEF'
    end
    object OracleDataSetkoefINDATE: TDateTimeField
      FieldName = 'INDATE'
      Required = True
    end
  end
  object OracleQueryufzpl: TOracleQuery
    SQL.Strings = (
      'SELECT nvl(SUM(SUM),0) UFzpl'
      '               FROM RECALCJUR'
      '              WHERE EMP#=:emp'
      '                AND FACTMON=:yearp*100+:monthz'
      '                AND RSHOP=:shop'
      '                AND EXPEND#=:expend'
      '                AND PAY# IN (800,801,802)'
      '                AND FLAGINV = :flaginv'
      '')
    Session = OracleSession1
    Variables.Data = {
      0300000006000000040000003A454D5003000000000000000000000006000000
      3A5945415250030000000000000000000000070000003A4D4F4E54485A030000
      000000000000000000050000003A53484F500300000000000000000000000700
      00003A455850454E44030000000000000000000000080000003A464C4147494E
      56030000000000000000000000}
    Cursor = crSQLWait
    Left = 176
    Top = 88
  end
  object OracleQueryufblfss: TOracleQuery
    SQL.Strings = (
      'SELECT nvl(SUM(SUM),0) '
      '               FROM RECALCJUR'
      '              WHERE EMP#=:emp'
      '                AND FACTMON=:yearp*100+:monthz'
      '                AND RSHOP=:shop'
      '                AND EXPEND#=:expend'
      '                AND PAY# IN (808,809)'
      'AND FLAGINV = :flaginv')
    Session = OracleSession1
    Variables.Data = {
      0300000006000000040000003A454D5003000000000000000000000006000000
      3A5945415250030000000000000000000000070000003A4D4F4E54485A030000
      000000000000000000050000003A53484F500300000000000000000000000700
      00003A455850454E44030000000000000000000000080000003A464C4147494E
      56030000000000000000000000}
    Cursor = crSQLWait
    Left = 368
    Top = 88
  end
  object OracleDataSetsecret: TOracleDataSet
    SQL.Strings = (
      'sELECT NVL(PRV#,0) pse'
      '               FROM EMPPRV'
      '              WHERE EMP#=:emp'
      '                AND PRV# in(998,999)'
      
        '                AND to_date('#39'01/'#39'||ltrim(to_char(:monthR,'#39'09'#39'))|' +
        '|'#39'/'#39'||ltrim(to_char(:yearR)),'#39'dd/mm/yyyy'#39')>=PRVDATE1'
      
        '                AND to_date('#39'01/'#39'||ltrim(to_char(:monthR,'#39'09'#39'))|' +
        '|'#39'/'#39'||ltrim(to_char(:yearR)),'#39'dd/mm/yyyy'#39')<NVL(PRVDATE2,to_date(' +
        #39'31/12/2999'#39','#39'dd/mm/yyyy'#39'))'
      '')
    Variables.Data = {
      0300000003000000040000003A454D5003000000000000000000000007000000
      3A4D4F4E544852030000000000000000000000060000003A5945415252030000
      000000000000000000}
    QBEDefinition.QBEFieldDefs = {040000000100000003000000505345010000000000}
    Cursor = crSQLWait
    Session = OracleSession1
    Left = 384
    Top = 8
    object OracleDataSetsecretPSE: TFloatField
      FieldName = 'PSE'
    end
  end
  object OracleDataSetuvol: TOracleDataSet
    SQL.Strings = (
      'SELECT  rtrim(NVL(to_char(DSCHDATE,'#39'dd/mm/yyyy'#39'),'#39' '#39')) dschdate'
      '        FROM EMPLOY'
      '        WHERE EMP# = :emp'
      '')
    Variables.Data = {0300000001000000040000003A454D50030000000000000000000000}
    QBEDefinition.QBEFieldDefs = {0400000001000000080000004453434844415445010000000000}
    Session = OracleSession1
    Left = 272
    Top = 16
    object OracleDataSetuvolDSCHDATE: TStringField
      FieldName = 'DSCHDATE'
      Size = 75
    end
  end
  object OracleQueryufbol: TOracleQuery
    SQL.Strings = (
      'SELECT nvl(SUM(SUM),0) '
      '               FROM RECALCJUR'
      '              WHERE EMP#=:emp'
      '                AND FACTMON=:yearp*100+:monthz'
      '                AND RSHOP=:shop'
      '                AND EXPEND#=:expend'
      '                AND PAY# IN (806,807)'
      'AND FLAGINV = :flaginv')
    Session = OracleSession1
    Variables.Data = {
      0300000006000000040000003A454D5003000000000000000000000006000000
      3A5945415250030000000000000000000000070000003A4D4F4E54485A030000
      000000000000000000050000003A53484F500300000000000000000000000700
      00003A455850454E44030000000000000000000000080000003A464C4147494E
      56030000000000000000000000}
    Cursor = crSQLWait
    Left = 264
    Top = 88
  end
  object OracleQueryufgpd: TOracleQuery
    SQL.Strings = (
      'SELECT nvl(SUM(SUM),0) '
      '               FROM RECALCJUR'
      '              WHERE EMP#=:emp'
      '                AND FACTMON=:yearp*100+:monthz'
      '                AND RSHOP=:shop'
      '                AND EXPEND#=:expend'
      '                AND PAY# IN (804,805)'
      'AND FLAGINV = :flaginv')
    Session = OracleSession1
    Variables.Data = {
      0300000006000000040000003A454D5003000000000000000000000006000000
      3A5945415250030000000000000000000000070000003A4D4F4E54485A030000
      000000000000000000050000003A53484F500300000000000000000000000700
      00003A455850454E44030000000000000000000000080000003A464C4147494E
      56030000000000000000000000}
    Cursor = crSQLWait
    Left = 456
    Top = 88
  end
  object ODScount: TOracleDataSet
    SQL.Strings = (
      'select distinct nvl(monthyear,0) monthyear'
      ' FROM PAYRECJUR'
      '                 WHERE YEAR=:yearp'
      '                      '
      '                       AND EMP# BETWEEN  :empmin AND :empmax')
    Variables.Data = {
      0300000003000000060000003A59454152500300000000000000000000000700
      00003A454D504D494E030000000000000000000000070000003A454D504D4158
      030000000000000000000000}
    QBEDefinition.QBEFieldDefs = {0400000001000000090000004D4F4E544859454152010000000000}
    Session = OracleSession1
    Left = 216
    Top = 168
    object ODScountMONTHYEAR: TFloatField
      FieldName = 'MONTHYEAR'
    end
  end
end
