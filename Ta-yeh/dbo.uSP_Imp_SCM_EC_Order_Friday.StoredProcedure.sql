USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_SCM_EC_Order_Friday]    Script Date: 07/24/2017 14:43:59 ******/
DROP PROCEDURE [dbo].[uSP_Imp_SCM_EC_Order_Friday]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_Imp_SCM_EC_Order_Friday]
as
begin
  /***********************************************************************************************************
     2017/04/17 Rickliu
     1.Friday_Order �� �ɶ��bEC���x���� ���x�ϥΡC
  ***********************************************************************************************************/
  Declare @Proc Varchar(50) = 'uSP_Imp_SCM_EC_Order_Friday'
  Declare @Cnt Int =0
  Declare @RowCnt Int =0
  Declare @Msg Varchar(4000) =''
  Declare @Errcode int = -1

  Declare @Kind Varchar(1), @KindName Varchar(10), @OR_Class Varchar(1), @OR_Name Varchar(20), @OR_CName Varchar(20)
  Declare @CompanyName Varchar(100), @CompanyLikeName Varchar(100), @Rm Varchar(100), @Or_maker Varchar(20), @str Varchar(200)
  Declare @Pos Int
  Declare @Or_Date1 varchar(10), @Or_Date2 varchar(10)
  Declare @or_wkno varchar(20) -- ��ڤW�����ʳ渹��@�Ȥ�q��渹 
  Declare @Or_Cnt int

  Declare @xls_head Varchar(100), @TB_head Varchar(200), @TB_xls_Name Varchar(200), @TB_OR_Name Varchar(200), @TB_tmp_name Varchar(200), @TB_Chk_Order Varchar(200)

  Declare @TB_EZCat_List varchar(200)

  Declare @strSQL Varchar(Max)
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @RowCount Table (cnt int)
  Declare @RowData Table (aData Varchar(255))
  
  Declare @XlsFileName Varchar(255) = ''
  Declare @Xls_Imp_Date Varchar(30) = ''
  
  Declare @New_Orno varchar(10)
  Declare @Last_Date Varchar(2)
  Declare @Sct_SubQuery Varchar(1000) -- �P�P�����l�d��

  Declare @F12 varchar(10) --���ʳ渹
  Declare @Prog_Flag Varchar(1000)

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  Begin Try
  Set @CompanyName = '�ɶ��b' --> �ФŶ��ܰ�
    Set @CompanyLikeName = '%'+@CompanyName+'%'
    Set @Or_Maker = 'Admin'
    Set @Sct_SubQuery = '(select Top 1 @@ from SYNC_TA13.dbo.sctsale where 1=1 and ss_ctno LIKE ''%''+Rtrim(m.ct_no)+''%'' and ss_no Like ''%''+Rtrim(d1.sk_no)+''%'' and (ss_edate = ''1900/01/01'' or ss_edate >= getdate()) order by ss_edate desc ,ss_sdate desc)'

    set @Kind = '1'
    set @KindName = ''
    set @OR_Name = ''
    set @OR_CName = '�z�f��q��'
    set @OR_Class = '3'

    Set @xls_head = 'Ori_xls#'
    Set @TB_head = 'EC_Order_Friday'
    Set @TB_xls_Name = @xls_head+@TB_Head -- Ex: Ori_xls#EC_Order_Friday-- �~�����ɶפJ���
    Set @TB_tmp_Name = @TB_Head+'_'+'tmp' -- Ex: EC_Order_Friday_tmp -- �{����J�B�z������
    Set @TB_OR_Name = @TB_Head -- Ex: EC_Order_Friday  -- ����J��V��Ƥ�����
    Set @TB_Chk_Order = '##Chk_'+@TB_Head+'_tmp' -- Ex: ##Chk_EC_Order_Friday_tmp --< ����Ʀs��� tempdb�A�{�����}�ɷ|�M���C
    Set @TB_EZCat_List = 'EC_EZCat_List'
    
    Set @Rm = '�t�ζפJ'+@CompanyName
    set @Or_Date1 = Convert(Varchar(10), CONVERT(Date, getdate()) , 111)
    set @Prog_Flag = @CR+Replicate('=*=', 50)+@CR+
                     '['+@Proc+'] Flag -->>> ### <<<'+
                     @CR+Replicate('=*=', 50)+@CR

    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    Set @Msg = '�ˬd�~�����ɶפJ��ƪ�O�_�s�b ['+@TB_xls_Name+']'
    print Replace(@Prog_Flag, '###', @Msg)
    IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_xls_Name+']') AND type in (N'U'))
    begin
       Set @Msg = '�~�����ɶפJ��ƪ� ['+@TB_xls_Name+']���s�b�A�פ�i�����ɧ@�~�C'
       print Replace(@Prog_Flag, '###', @Msg)
       set @strSQL = ''
       RaisError(@Msg, 16, 1)
    end
    else
    begin
       Set @Msg = '�ˬd�~�����ɶפJ�O�_����ơC'
       print Replace(@Prog_Flag, '###', @Msg)

       Set @strSQL = 'select count(1) from '+@TB_xls_Name
       print @strSQL
       delete @RowCount
       Insert Into @RowCount Exec(@strSQL)
       select @cnt = IsNull(cnt, 0) from @RowCount
       If @Cnt = 0
       begin
          Set @Msg = @Msg + '...�L��ơA�פ����!!'
          print Replace(@Prog_Flag, '###', @Msg)
          Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
          Goto proc_exit
       end
       else
       begin
          Set @Msg = @Msg + '...����ơA�~�����!!'
          print Replace(@Prog_Flag, '###', @Msg)
          Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
       end
    end
    
    Set @Msg = '�ˬd�¿߮榡��ƪ�O�_�s�b!!'
    print Replace(@Prog_Flag, '###', @Msg)
    IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_EZCat_List+']') AND type in (N'U'))
    begin
       Set @Msg = '�s�W Table ['+@TB_EZCat_List+']�C'
       Set @strSQL = 'CREATE TABLE [dbo].['+@TB_EZCat_List+'](' +@CR+
                     '       [UniqueID] [int] IDENTITY(1,1) NOT NULL, ' +@CR+
                     '       [Order_NO] [varchar](100) NULL, ' +@CR+
                     '       [EC_NO] [varchar](100) NULL, ' +@CR+
                     '       [EC_Name] [varchar](100) NULL, ' +@CR+
                     '       [sp_date] [varchar](10) NULL, ' +@CR+
                     '       [sp_date_year] [varchar](4) NULL, ' +@CR+
                     '       [sp_date_month] [varchar](2) NULL, ' +@CR+
                     '       [sp_date_day] [varchar](2) NULL, ' +@CR+
                     '       [sp_date_YM] [varchar](6) NULL, ' +@CR+
                     '       [sp_date_MD] [varchar](4) NULL, ' +@CR+
                     '       [sp_date_YMD] [varchar](8) NULL, ' +@CR+
                     '       [buyer] [varchar](100) NULL, ' +@CR+
                     '       [cust] [varchar](100) NULL, ' +@CR+
                     '       [cust_phone] [varchar](100) NULL, ' +@CR+
                     '       [cust_mobile] [varchar](100) NULL, ' +@CR+
                     '       [cust_addr1] [varchar](100) NULL, ' +@CR+
                     '       [cust_addr2] [varchar](100) NULL, ' +@CR+
                     '       [sk_no] [varchar](100) NULL, ' +@CR+
                     '       [sk_name] [varchar](100) NULL, ' +@CR+
                     '       [source_sk_name] [varchar](250) NULL, ' +@CR+
                     '       [isfound] [varchar](1) NULL, ' +@CR+
                     '       [qty] [varchar](100) NULL, ' +@CR+
                     '       [Price] [varchar](100) NULL, ' +@CR+
                     '       [MEMO] [varchar](255) NULL, ' +@CR+
                     '       [source_file] [varchar](255) NULL, ' +@CR+
                     '       [import_date] [datetime] NULL, ' +@CR+
                     '       [Rowid] [int] NULL ' +@CR+
                     ') ON [PRIMARY] ' 
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    end

    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print ' F01:����               '
print ' F02:�q��s��           '
print ' F03:�q������           '
print ' F04:�q�ʤ�             '
print ' F05:���X�f��           '
print ' F06:�̱ߥX�f��         '
print ' F07:�q�ʤH             '
print ' F08:����H             '
print ' F09:����H���         '
print ' F10:����a�}           '
print ' F11:�Ȥ�q��Ƶ�       '
print ' F12:(�ӫ~�s��)�ӫ~�W�� '
print ' F13:�����ӮƸ�         '
print ' F14:�W��               '
print ' F15:�ƶq               '
print ' F16:����               '
print ' F17:�q�沧�`�O��       '

    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    --�N�C����ƥ[�J�ߤ@���(ps.����ƪ�ФűƧǡA�O�d��l Text �˻��A�H�Q��Უ�X��b��)
    --�إ߼Ȧs�ɥB�إߧǸ��A�إߧǸ��B�J�ܭ��n�A�]������n�̧ǲ��ͩ��W
    IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_tmp_Name+']') AND type in (N'U'))
    begin
       Set @Msg = '�M�� ['+@CompanyName+'] �{����J������ƪ� ['+@TB_tmp_Name+']�C'
       print Replace(@Prog_Flag, '###', @Msg)

       Set @strSQL = 'DROP TABLE [dbo].['+@TB_tmp_Name+']'
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    end

    Set @Msg = '�פJ ['+@CompanyName+'] �{����J������ƪ� ['+@TB_tmp_Name+']�C'
    print Replace(@Prog_Flag, '###', @Msg)
    Set @strSQL = 'select rowid, '+@CR+ -- �C��
                  '       F1, '+@CR+ --����
                  '       '''+@CompanyName+''' as F2, '+@CR+  --���W
                  '       F1 as F3, '+@CR+  --��      
                  '       replace(dbo.uFn_RegexReplace(''[^A-Z0-9]'', '''', F13,1,1),'' '','''') as F4, '+@CR+  --�ӫ~�s��
                  '       F12 as F5, '+@CR+  --�ӫ~�W��
                  '       replace(dbo.uFn_RegexReplace(''[^A-Z0-9]'', '''', F15,1,1),'' '','''') as F6, '+@CR+  --�q�f�q  
                  '       F16 as F7, '+@CR+  --�ث~�_  
                  '       replace(dbo.uFn_RegexReplace(''[^A-Z0-9]'', '''', F15,1,1),'' '','''') as F8, '+@CR+  --�z�f�q  
                  '       replace(dbo.uFn_RegexReplace(''[^A-Z.0-9]'', '''', F16,1,1),'' '','''') as F9, '+@CR+  --�i��    
                  '       Convert(Varchar(100), '''') as F10, '+@CR+ --�p�p    
                  '       Convert(Varchar(100), '''') as F11, '+@CR+ --�Ƶ�    
                  '       Convert(Varchar(100), '''') as F12, '+@CR+ --���ʳ渹    
                  '       F5 as F13, '+@CR+ --���w��f��    
                  '       xlsFileName as source_file,  '+@CR+ --��J�ɮצW��    
                  '       replace(dbo.uFn_RegexReplace(''[^A-Z0-9]'', '''', F2,1,1),'' '','''') as source_order, '+@CR+  
                  '       F7 as buyer, '+@CR+
                  '       F8 as cust, '+@CR+
                  '       F9 as cust_phone, '+@CR+ 
                  '       F9 as cust_mobile, '+@CR+ 
                  '       F10 as cust_addr1, '+@CR+ 
                  '       F10 as cust_addr2, '+@CR+ 
                  '       F11 as MEMO, '+@CR+ 
                  '       print_date='''+@Or_Date1+''', '+@CR+
                  '       import_date=convert(datetime, getdate()) '+@CR+
                  '       into ['+@TB_tmp_Name+']'+@CR+
                  '  from ['+@TB_xls_Name+']'
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    -- 2017/04/07 Rickliu ��� NanLiao ���Ҷq���ӬO�A�C�@�� XLS �|���h�������q�q��A�]�����F��J�Ȥ�q�渹�X�A�ҥH�ĥΦ��覡���o�s�s���C
    Set @Msg = '������ʳ渹��� [SYNC_TA13.dbo.sorder]�C'
    print Replace(@Prog_Flag, '###', @Msg)
    Set @strSQL = ';With QTE_Q1 as( '+@CR+
                  '  select distinct rtrim(or_wkno) as or_wkno '+@CR+
                  '    from SYNC_TA13.dbo.sorder m '+@CR+
                  '   where 1=1 '+@CR+
                  '     and or_class = '''+@Or_Class+''' '+@CR+
                  '     and or_date1 = '''+@Or_Date1+''' '+@CR+
                  '     and or_ctname like '''+@CompanyLikeName+''' '+@CR+
                  '   union '+@CR+
                  '  select distinct rtrim(or_wkno) as or_wkno '+@CR+
                  '    from SYNC_TA13.dbo.sordertmp m '+@CR+
                  '   where 1=1 '+@CR+
                  '     and or_class = '''+@Or_Class+''' '+@CR+
                  '     and or_date1 = '''+@Or_Date1+''' '+@CR+
                  '     and or_ctname like '''+@CompanyLikeName+''' '+@CR+
                  ') '+@CR+
                  'select convert(varchar(10),'+@CR+
                  '               isnull(case'+@CR+
                  '                        when rtrim(max(or_wkno)) = '''''+@CR+
                  '                        then null'+@CR+
                  '                        else rtrim(max(or_wkno))'+@CR+
                  '                      end,'+@CR+
                  '               replace(substring('''+@Or_Date1+''', 3, 8), ''/'', '''') +''0000'') + 1) as Data'+@CR+
                  '  from QTE_Q1'
    print @strSQL
    delete @RowData
    Insert Into @RowData Exec(@strSQL)
    Select @F12 = Isnull(aData, '') From @RowData

    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    Set @Msg = '�ˬd �w/���T�{�� �O�_���g�פJ�L ['+@TB_Chk_Order+']!!!'
    print Replace(@Prog_Flag, '###', @Msg)
    Set @strSQL = 'IF OBJECT_ID(N''tempdb.dbo.'+@TB_Chk_Order+''') Is Not Null '+@CR+
                  '   Drop Table '+@TB_Chk_Order+@CR+
                  'Create Table '+@TB_Chk_Order+' ( '+@CR+
                  '       Kind Varchar(50), '+@CR+
                  '       or_no Varchar(50), '+@CR+
                  '       or_wkno Varchar(50), '+@CR+
                  '       or_ctno Varchar(50), '+@CR+
                  '       or_ctname Varchar(255), '+@CR+
                  '       or_maker Varchar(50), '+@CR+
                  '       or_ispack Varchar(1), '+@CR+
                  '       or_rem Varchar(255) '+@CR+
                  ') '+@CR+
                  'Insert Into '+@TB_Chk_Order+@CR+
                  'select ''�w�T�{'' as Kind, '+@CR+
                  '       rtrim(or_no) as or_no, rtrim(or_wkno) as or_wkno, rtrim(or_ctno) as or_ctno, '+@CR+
                  '       rtrim(or_ctname) as or_ctname, rtrim(or_maker) as or_maker, rtrim(or_ispack) as or_ispack, '+@CR+
                  '       rtrim(substring(or_rem, 1, 255)) as or_rem '+@CR+
                  '  from SYNC_TA13.dbo.sorder m '+@CR+
                  ' where 1=1 '+@CR+
                  '   and or_class = '''+@or_class+''' '+@CR+
                  '   and or_date1 = '''+@Or_Date1+''' '+@CR+
                  '   and or_wkno collate Chinese_Taiwan_Stroke_CI_AS = '''+@F12+''' '+@CR+
                  '   and or_rem like ''%'+@Rm+'%'' '+@CR+
                  ' union '+@CR+
                  'select ''���T�{'' as Kind, '+@CR+
                  '       rtrim(or_no) as or_no, rtrim(or_wkno) as or_wkno, rtrim(or_ctno) as or_ctno, '+@CR+
                  '       rtrim(or_ctname) as or_ctname, rtrim(or_maker) as or_maker, rtrim(or_ispack) as or_ispack, '+@CR+
                  '       rtrim(substring(or_rem, 1, 255)) as or_rem '+@CR+
                  '  from SYNC_TA13.dbo.sordertmp m '+@CR+
                  ' where 1=1 '+@CR+
                  '   and or_class = '''+@Or_class+''' '+@CR+
                  '   and or_date1 = '''+@Or_Date1+''' '+@CR+
                  '   and or_wkno collate Chinese_Taiwan_Stroke_CI_AS = '''+@F12+''' '+@CR+
                  '   and or_rem like ''%'+@Rm+'%'' '
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    Set @strSQL = 'select count(1) from '+@TB_Chk_Order
    print @strSQL
    delete @RowCount
    Insert Into @RowCount Exec(@strSQL)
    select @Or_cnt = IsNull(cnt, 0) from @RowCount

    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    Set @strSQL = 'select top 1 XlsFileName from '+@TB_xls_Name
    Print @strSQL
    delete @RowData
    insert into @RowData exec(@strSQL)
    select @XlsFileName = Isnull(aData, '') from @RowData

    Set @strSQL = 'select top 1 Convert(Varchar(30), imp_date, 120) from '+@TB_xls_Name
    Print @strSQL
    delete @RowData
    insert into @RowData exec(@strSQL)
    select @Xls_Imp_Date = Isnull(aData, '') from @RowData

    if @Or_cnt > 0
    begin
       
       set @Msg =@Proc+'...�ɮ�:'+@XlsFileName+'...�Ȥ�:'+@CompanyName+'[�w/���T�{��w�� ['+@Or_Date1+' '+@OR_CName+'] ��ơA�Y�ݭn����h����M���T�{�Υ��T�{��ګ᭫��Y�i�C'
       print Replace(@Prog_Flag, '###', @Msg)
       exec uSP_Sys_CustomTable2HTML @TB_Chk_Order, @strSQL output

       RaisError(@Msg, 16, 1)
    end

    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    Set @strSQL = 'IF OBJECT_ID(N''tempdb.dbo.'+@TB_Chk_Order+''') Is Not Null '+@CR+
                  '   Drop Table '+@TB_Chk_Order
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    
    Set @Msg = '�B�z�U������� ['+@TB_tmp_Name+']�C'
    print Replace(@Prog_Flag, '###', @Msg)
    Set @strSQL = 'update ['+@TB_tmp_Name+']'+@CR+
                  '   set F3=Replace(LTrim(RTrim(Substring(Convert(Text, F1),   1,  2))), ''('', ''''), '+@CR+ --��
                  '       source_order=case when Charindex(''-'', source_order)-1 <= 0 then source_order else substring(source_order,1,Charindex(''-'',source_order)-1) end, '+@CR+ --�ӫ~�s��
                  '       F4=case when LTrim(RTrim(F4)) > '''' then F4 else ''N'' end, '+@CR+ --�ӫ~�s��
                  '       F7=case when LTrim(RTrim(Substring(Convert(Text, F9),   1, 100))) = ''0'' then ''0'' else ''1'' end, '+@CR+ --�ث~�_  
                  '       F9=case when PATINDEX(''%[0-9]%'',F9)>0 then Convert(Numeric(20, 6), F9)/1.05 else 0 end, '+@CR+ --�i��
                  '       F10=Convert(Numeric(20, 6), PATINDEX(''%[0-9]%'',Convert(Varchar(30), F6))) * Convert(Numeric(20, 6), PATINDEX(''%[0-9]%'',Convert(Varchar(30), F9))), '+@CR+ --�p�p
                  '       F11=LTrim(RTrim(Substring(Convert(Text, F11),   1, 100)))  '  --�Ƶ�
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    Set @Msg = '�R���{����J������ƪ��D���n��� ['+@TB_tmp_Name+']�C'
    print Replace(@Prog_Flag, '###', @Msg)
    Set @strSQL = 'Delete ['+@TB_tmp_Name+'] where PATINDEX(''%[0-9]%'',F6) = 0 '
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    
    Set @strSQL = 'Delete ['+@TB_tmp_Name+'] where LTrim(RTrim(source_order))='''' '
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    
--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
--=-[���ͳ�ک��Ӹ�ơA�ƶq���t�ƪ��n���Ͱh�f��]=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_OR_Name+']') AND type in (N'U'))
    begin
       Set @Msg = '�R������J��ƪ� ['+@TB_OR_Name+']'
       print Replace(@Prog_Flag, '###', @Msg)
       Set @strSQL = 'IF OBJECT_ID(N'''+@TB_OR_Name+''') Is Not Null '+@CR+
                     '   Drop Table '+@TB_OR_Name
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    end

    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    --2013/04/25 ���@���v����J�N�n���Υ[�`�A����n��
    --�b���i����ӫ~�s���A�|�����U�s���H�ΰӫ~�򥻸���ɶi����@�~
    Set @Msg = '�i����ӫ~�s���üg�J ['+@TB_OR_Name+'] ����J��ƪ�C'
    print Replace(@Prog_Flag, '###', @Msg)
    set @strSQL = ';With CTE_Q1 as ( '+@CR+
                  '  select rtrim(ct_no) as ct_no, rtrim(ct_sname) as ct_sname, rtrim(ct_name)+''#''+rtrim(ct_sname) as ct_ssname '+@CR+
                  '    from SYNC_TA13.dbo.PCUST  '+@CR+
                  '   where ct_class =''1'' '+@CR+
                  '     and rtrim(ct_name)+''#''+rtrim(ct_sname) like '''+@CompanyLikeName+''' '+@CR+
                  '     and substring(ct_no, 9, 1) ='''+@Kind+''' '+@CR+ -- -- 1:�ʳf, 2:3C, 3:�H��, 4:�N�uOEM
                  ') '+@CR+
                  'select rowid, ct_no, ct_sname, ct_ssname, '+@CR+
                  '       d1.sk_no as sk_no, '+@CR+
                  '       d1.sk_name as sk_name, '+@CR+
                  -- ��ӫ~�W��
                  '       f5, '+@CR+ 
                  '       d1.sk_bcode as sk_bcode, '+@CR+
                  -- ��ӫ~�s��
                  '       f4, '+@CR+ 
                  -- ��ӫ~�ƶq
                  '       Convert(Numeric(20, 6), F6) as Ori_sale_qty, '+@CR+ -- 2015/08/19 Rickliu �O�d�쥻���A�קK�᭱�ק�
                  '       Convert(Numeric(20, 6), F6) as sale_qty, '+@CR+
                  -- ���
                  '       Convert(Numeric(20, 6), F9) as Ori_sale_amt, '+@CR+ -- 2015/08/19 Rickliu �O�d�쥻���A�קK�᭱�ק�
                  '       Convert(Numeric(20, 6), F9) as sale_amt, '+@CR+
                  -- �p�p
                  '       (Convert(Numeric(20, 6), F9) * Convert(Numeric(20, 6), F6)) as Ori_sale_tot, '+@CR+ -- 2015/08/19 Rickliu �O�d�쥻���A�קK�᭱�ק�
                  '       (Convert(Numeric(20, 6), F9) * Convert(Numeric(20, 6), F6)) as sale_tot, '+@CR+ 
                  -- �Ƶ�
                  '       F11,  '+@CR+ 
                  -- ���ʳ渹
                  '       ''' + @F12 + ''' AS F12,  '+@CR+ 
                  -- ���w��f��
                  '       F13,  '+@CR+ 
                  -- �ӷ��ɦW
                  '       source_file,  '+@CR+ 
                  -- �ӷ��q��s��
                  '       source_order,  '+@CR+ 
                  '       buyer ,cust ,cust_phone ,cust_mobile ,cust_addr1 ,cust_addr2 ,MEMO , '+@CR+ 
                  -- �O�_����
                  '       case when d1.sk_no is null then ''N'' else ''Y'' end as isfound, '+@CR+
                  -- 2015/07/23 Rickliu �W�[�P�P�������A�Φ��g�k�O���F�קK Cursor �ӱ��j�q�귽�A�]���� subQuery ���@�k�A���M���O�̦n���k�C
                  -- �P�P��ץN��
                  '       sct_ss_csno = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_csno')+', ''''), '+@CR+ 
                  -- �P�P��צW��
                  '       sct_ss_csname = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_csname')+', ''''), '+@CR+ 
                  -- ���ӥN�X
                  '       sct_ss_rec = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_rec')+', 0), '+@CR+ 
                  -- �Ȥ������νs��
                  '       sct_ss_ctkind = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_ctkind')+', ''''), '+@CR+ 
                  -- �Ȥ�s��
                  '       sct_ss_ctno = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_ctno')+', ''''), '+@CR+ 
                  -- �f�~�����νs��
                  '       sct_ss_nokind = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_nokind')+', ''''), '+@CR+ 
                  -- �f�~�s��
                  '       sct_ss_no = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_no')+', ''''), '+@CR+ 
                  -- �ث~�s��
                  '       sct_ss_sendno = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_sendno')+', ''''), '+@CR+ 
                  -- �~���ƶq 
                  '       sct_ss_noqty = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_noqty')+', 0), '+@CR+ 
                  -- �ث~�ƶq
                  '       sct_ss_sendqty = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_sendqty')+', 0), '+@CR+ 
                  -- �榸�ث~���� 
                  '       sct_ss_oneqty = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_oneqty')+', 0), '+@CR+ 
                  -- �Ȥ��ث~����
                  '       sct_ss_itmqty = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_itmqty')+', 0), '+@CR+ 
                  -- �`�ث~����
                  '       sct_sendtot = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_sendtot')+', 0), '+@CR+ 
                  -- ���İ_�l���
                  '       sct_ss_sdate = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_sdate')+', getdate()), '+@CR+ 
                  -- ���ĺI����
                  '       sct_ss_edate = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_edate')+', getdate()), '+@CR+ 
                  -- �_�l�ƶq
                  '       sct_ss_sqty = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_sqty')+', 0), '+@CR+ 
                  -- ����ƶq
                  '       sct_ss_eqty = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_eqty')+', 0), '+@CR+ 
                  -- ���椽��
                  '       sct_ss_form = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_form')+', ''''), '+@CR+ 
                  -- �O�_���ث~ 0: �O, 1: �_
                  '       F7 as isSend '+@CR+
                  '       into ['+@TB_OR_Name+'] '+@CR+
                  '  from CTE_Q1 m '+@CR+
                  '       Inner join ['+@TB_tmp_Name+'] d '+@CR+
                  '          on m.ct_ssname collate Chinese_Taiwan_Stroke_CI_AS like ''%''+rtrim(d.f2)+''%'' collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                  '         and rtrim(d.f2)<>'''' '+@CR+ -- 2017/04/27 Rickliu ������Ů����A�]���O�ϥ� LIKE ���覡�A�Y�@���S���Ȥ�W�ٳf�s���|�~�P�Ҧ��Ȥ᳣�n��J�C
                  '        left join SYNC_TA13.dbo.sstock d1 '+@CR+
                  '          on (ltrim(rtrim(d.f4)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d1.sk_no)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                  '          or  ltrim(rtrim(d.f4)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d1.sk_bcode)) collate Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                  '         and d1.sk_bcode <> '''' '+@CR+
                  ' order by 1 '
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    -- 2015/08/19 Rickliu �W�[�䴩 ��@�ث~
    Set @Msg = '�i����ӫ~�s���üg�J ['+@TB_OR_Name+'] ����J��ƪ�C'
    print Replace(@Prog_Flag, '###', @Msg)
    set @strSQL = 'insert into ['+@TB_OR_Name+'] '+@CR+
                  'select rowid, ct_no, ct_sname, ct_ssname, '+@CR+
                  '       sct_ss_sendno as sk_no, '+@CR+
                  '       d.sk_name as sk_name, '+@CR+
                  '       d.sk_name as f5, '+@CR+
                  '       d.sk_bcode as sk_bcode, '+@CR+
                  '       f4, '+@CR+
                  '       Ceiling(sale_qty / sct_ss_sendqty) as Ori_sal_qty, '+@CR+
                  '       Ceiling(sale_qty / sct_ss_sendqty) as sal_qty, '+@CR+
                  '       0 as Ori_sale_amt, '+@CR+
                  '       0 as sale_amt, '+@CR+
                  '       0 as Ori_sale_tot, '+@CR+
                  '       0 as sale_tot, '+@CR+
                  '       f11, f12, f13, '+@CR+
                  '       source_file, source_order, '+@CR+
                  '       buyer, cust, cust_phone, cust_mobile, cust_addr1, cust_addr2, MEMO, '+@CR+ 
                  '       ''Y'' as isfound, '+@CR+
                  '       sct_ss_csno, '+@CR+
                  '       sct_ss_csname, '+@CR+
                  '       sct_ss_rec, '+@CR+
                  '       sct_ss_ctkind, '+@CR+
                  '       sct_ss_ctno, '+@CR+
                  '       sct_ss_nokind, '+@CR+
                  '       sct_ss_no, '+@CR+
                  '       sct_ss_sendno, '+@CR+
                  '       sct_ss_noqty, '+@CR+
                  '       sct_ss_sendqty, '+@CR+
                  '       sct_ss_oneqty, '+@CR+
                  '       sct_ss_itmqty, '+@CR+
                  '       sct_sendtot, '+@CR+
                  '       sct_ss_sdate, '+@CR+
                  '       sct_ss_edate, '+@CR+
                  '       sct_ss_sqty, '+@CR+
                  '       sct_ss_eqty, '+@CR+
                  '       sct_ss_form, '+@CR+
                  '       2 as isSend '+@CR+
                  '  from ['+@TB_OR_Name+'] m '+@CR+
                  '       left join fact_sstock d on m.sct_ss_sendno = d.sk_no '+@CR+
                  ' where 1=1 '+@CR+
                  '   and Rtrim(sct_ss_csno + sct_ss_csname) <> '''' '+@CR+
                  '   and sct_ss_noqty <> 0 '+@CR+
                  '   and sct_ss_sendqty <> 0 '+@CR+
                  ' order by sct_ss_rec '
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

    Set @Msg = '�i���s���ƪ� rowid ['+@TB_OR_Name+'] ����J��ƪ�C'
    print Replace(@Prog_Flag, '###', @Msg)
    set @strSQL = 'update ['+@TB_OR_Name+'] '+@CR+
                  '   set rowid = rowid * 10 '+@CR+
                  ' where isnull(sk_no, '''') = sct_ss_sendno '+@CR+
                  '   and isnull(sct_ss_no, '''') <> '''' ' 
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL            
    
    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    set @Cnt = 0
    Set @Msg = '�ˬd ['+@TB_OR_Name+'] �����ɸ�ƬO�_�s�b�C'
    print Replace(@Prog_Flag, '###', @Msg)
    set @strSQL = 'select count(1) from ['+@TB_OR_Name+']'
    print @strSQL
    delete @RowCount
    insert into @RowCount Exec(@strSQL)

    select @Cnt=cnt from @RowCount
    if @Cnt = 0
    begin
       set @Msg = @Msg + '...���s�b�A�פ���J�{�ǡC'
       RaisError(@Msg, 16, 1)
    end
    else
    begin
       set @Msg = @Msg + '...�s�b�A�N�i����J�{�ǡC'
       Exec uSP_Sys_Write_Log @Proc, @Msg, '', @Cnt
    end

    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    --���o�̷s�X�f�h�^�渹(�����q�w�֭�H�Υ��֭㪺������)
    --�ФŦA�[+1�A�]���s�W�ɷ|�۰ʥ[�J
    --2013/11/24 �אּ @od_date (�C��T�w�ϥ� 25�鰵�����ɤ�)
    --�]���q�\��ƨS�k�o�򨳳t�Y�ɡA�]���u��^�Y�h�����V��Ʈw���渹
    set @Msg = '���o�̷s�q��渹�C'
    print Replace(@Prog_Flag, '###', @Msg)
    set @strSQL = ';With CTE_Q1 as ('+@CR+
                  '  select distinct rtrim(or_no) as or_no'+@CR+
                  '    from tyeipdbs2.lytdbta13.dbo.sorder m '+@CR+
                  '   where 1=1 '+@CR+
                  '     and or_class = '''+@or_class+''' '+@CR+
                  '     and or_date1 = '''+@Or_Date1+''' '+@CR+
                  '   union '+@CR+
                  '  select distinct rtrim(or_no) as or_no '+@CR+
                  '    from tyeipdbs2.lytdbta13.dbo.sordertmp m '+@CR+
                  '   where 1=1 '+@CR+
                  '     and or_class = '''+@Or_class+''' '+@CR+
                  '     and or_date1 = '''+@Or_Date1+''' '+@CR+
                  ')'+@CR+
                  'select convert(varchar(10), isnull(max(or_no), replace(substring('''+@Or_Date1+''', 3, 8), ''/'', '''') +''0000'')) '+@CR+
                  '  from CTE_Q1'
    print @strSQL
    delete @RowData
    insert into @RowData Exec (@strSQL)
    Select @New_Orno = Isnull(aData, '') from @RowData
    set @Msg = '���o�̷s�q��渹(�����q�w�֭�H�Υ��֭㪺������),new_orno=['+@New_Orno+'], or_class=['+@Or_Class+'], or_date=['+@Or_Date1+'].'
    print Replace(@Prog_Flag, '###', @Msg)
    Exec uSP_Sys_Write_Log @Proc, @Msg, '', @Cnt
                         
    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    --2014/04/22 ���`���ܪ�����T�{�q��A�L���A��f�֡C
    set @Msg = '�s�W ['+@Or_Date1+'] �q��D�ɡC'
    print Replace(@Prog_Flag, '###', @Msg)
    set @strSQL =';With CTE_Q1 as ( '+@CR+
                 '  select distinct m.ct_no,m.F12 '+@CR+
                 '    from ['+@TB_OR_Name+'] m '+@CR+
                 '         left join SYNC_TA13.dbo.pcust d1 on m.ct_no = d1.ct_no and ct_class =''1'' '+@CR+
                 '), CTE_Q2 as ( '+@CR+
                 '  select ct_no, F12 as or_wkno,  '+@CR+
                 '         convert(varchar(10), '+@New_Orno+'+row_number() over(order by ct_no)) as od_no '+@CR+ -- �f��s��
                 '    from CTE_Q1 '+@CR+
                 '), CTE_Q3 as ( '+@CR+
                 '  select ct_no , count(1) as cnt '+@CR+
                 '    from ['+@TB_OR_Name+']  '+@CR+
                 '   group by ct_no '+@CR+
                 ') '+@CR+
                 'insert into tyeipdbs2.lytdbta13.dbo.sorder '+@CR+
                 '(  [OR_CLASS], [OR_NO], [OR_DATE1], [OR_DATE2], [OR_CTNO] '+@CR+
                 ' , [OR_CTNAME], [OR_SALES], [OR_DPNO], [OR_MAKER], [OR_TOT] '+@CR+
                 ' , [OR_TAX], [OR_RATE_NM], [OR_RATE], [OR_WKNO], [OR_REM] '+@CR+
                 ' , [or_ispack]  '+@CR+
                 --20161223 add by NanLiao
                 ' , [or_tal_rec],[or_npack]  '+@CR+
                 ') '+@CR+
                 'select distinct '''+@Or_Class+''' as or_class, '+@CR+
                 '       d3.od_no, '+@CR+
                 '       convert(datetime, '''+@Or_Date1+''') as or_date1, '+@CR+
                 '       convert(datetime, '''+@Or_Date1+''') as or_date2, '+@CR+
                 '       d1.ct_no as or_ctno, '+@CR+ -- �Ȥ�s��
                 '       d1.ct_name as or_ctname, '+@CR+ -- �Ȥ�W��
                 '       d1.ct_sales as or_sales, '+@CR+ -- �~�ȭ�
                 '       d1.ct_dept as or_dpno, '+@CR+ -- �����s��
                 '       '''+@Or_Maker+''' as or_maker, '+@CR+ -- �s��H��
                 '       sum(m.sale_tot) as or_tot, '+@CR+ -- �p�p
                 '       sum(m.sale_tot)*0.05 as or_tax, '+@CR+ -- ��~�|(��)
                 '       ''NT'' as or_rate_nm, '+@CR+ -- �ײv�W��
                 '       1 as or_rate, '+@CR+ -- �ײv
                 '       F12 as or_wkno, '+@CR+ -- �Ȥ�q��渹
                 '       '''+@Rm+'-['+@OR_CName+']''+Char(13)+Char(10)+'+@CR+
                 '       ''�J��(''+Convert(Varchar(20), Getdate(), 120)+'')''+Char(13)+Char(10)+ '+@CR+
                 '       ''�{��('+@Proc+')''+Char(13)+Char(10)+ '+@CR+
                 '       ''�ɮ�('+@XlsFileName+')''+Char(13)+Char(10)+ '+@CR+ -- �Ƶ�
                 '       ''����('+@Xls_Imp_Date+')'' as or_rem, '+@CR+ -- �Ƶ�
                 '       ''0'' as or_ispack  '+@CR+ -- ��f��
                 --20161223 add by NanLiao
                 '       ,d4.cnt as or_tal_rec  '+@CR+ -- �����`����
                 '       ,d4.cnt as or_npack  '+@CR+ -- ���涵��
                 '  from ['+@TB_OR_Name+'] m '+@CR+
                 '       left join SYNC_TA13.dbo.sstock d '+@CR+
                 '         on rtrim(m.sk_no) = rtrim(d.sk_no) '+@CR+
                 '       left join SYNC_TA13.dbo.pcust d1 '+@CR+
                 '         on rtrim(m.ct_no) = rtrim(d1.ct_no) '+@CR+
                 '        and ct_class =''1'' '+@CR+
                 '       left join CTE_Q2 d3 '+@CR+
                 '         on d1.ct_no = d3.ct_no and m.F12 = d3.or_wkno '+@CR+
                 --20161223 add by NanLiao
                 '       left join CTE_Q3 d4 '+@CR+
                 '         on m.ct_no = d4.ct_no '+@CR+
                 ' where isfound =''Y'' '+@CR+
                 ' group by d3.od_no, '+@CR+
                 '       d1.ct_no, '+@CR+ -- �Ȥ�s��
                 '       d1.ct_name, '+@CR+ -- �Ȥ�W��
                 '       substring(d1.ct_addr3, 1, 255), '+@CR+ -- �e�f�a�}
                 '       d1.ct_sales, '+@CR+ -- �~�ȭ�
                 '       d1.ct_dept, '+@CR+ -- �����s��
                 '       d1.ct_porter, '+@CR+ -- �f�B���q
                 '       M.F12, '+@CR+ -- �Ȥ�q��渹
                 --20161223 add by NanLiao
                 '       d4.cnt, '+@CR+ -- ���ӵ���
                 '       d1.ct_payfg '
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

    set @Msg = '�s�W ['+@Or_Date1+'] ��q����ӡC'
    print Replace(@Prog_Flag, '###', @Msg)
    --2014/04/22 ���`���ܪ�����T�{�q��A�L���A��f�֡C
    set @strSQL ='insert into tyeipdbs2.lytdbta13.dbo.sorddt '+@CR+
                 '(od_class, or_no, od_date1, od_date2, od_ctno, '+@CR+
                 ' od_skno, od_name, od_price, od_unit, od_qty, '+@CR+
                 ' od_unit_fg, od_dis, od_stot, od_rate, od_ave_p, '+@CR+
                 ' od_lotno, od_seqfld, od_spec, od_sendfg, '+@CR+
                 ' od_csno, od_csrec,od_is_pack '+@CR+
                 ') '+@CR+
                 'select '''+@Or_Class+''' as od_class, '+@CR+ -- ���O
                 '       d2.OR_NO, '+@CR+ -- �f��s��
                 '       d2.OR_DATE1 as od_date1, '+@CR+ -- �f����
                 -- 20160205 modify by Nan ��ޱ���X�A�i�N��f����P�f�����ۦP�A�Y�����ťաA����N�L�k�ק��f���A�C
                 '       d2.OR_DATE2 as od_date2, '+@CR+ -- ��f���
                 '       d2.OR_CTNO as od_ctno, '+@CR+ -- �Ȥ�s��
                 '       d.sk_no as od_skno, '+@CR+ -- �f�~�s��
                 '       d.sk_name as od_name, '+@CR+ -- �~�W�W��
                 '       m.sale_amt as od_price, '+@CR+ -- ����
                 '       d.sk_unit as od_unit, '+@CR+ -- ���
                 '       m.sale_qty as od_qty, '+@CR+ -- ����ƶq
                 '       0 as od_unit_fg, '+@CR+ -- ���X��
                 '       1 as od_dis, '+@CR+ -- ���
                 '       m.sale_tot as od_stot, '+@CR+ -- �p�p
                 '       1 as od_rate, '+@CR+ -- �ײv
                 '       d.s_price4 as od_ave_p, '+@CR+ -- ��즨��
                 '       '''+@Rm+@OR_CName +''' as od_lotno, '+@CR+ -- �ƥ����
                 '       RANK() OVER (ORDER BY m.source_order,d.sk_no) as od_seqfld, '+@CR+ -- ���ӧǸ�
                 '       source_order  as od_spec, '+@CR+ -- ���ӧǸ�, ���Ǹ��|�]����V�ק�ӥ[�H�ܧ�A�ҥH�t�s�@���� od_spec
                 '       Case '+@CR+
                 '         when m.sale_amt = 0 '+@CR+ --2013/12/28 �ȪA�۴f��ܡA�Y���B���s�h�N���ث~
                 '         then 1 '+@CR+
                 '         else 0 '+@CR+
                 '       end as od_sendfg, '+@CR+ -- �O�_���ث~
                 '       sct_ss_csno as od_csno,  '+@CR+ -- �P�P��ץN��
                 '       sct_ss_rec as od_csrec,  '+@CR+ -- �P�P��ש��ӥN�X
                 '       ''0'' as od_is_pack  '+@CR+ -- ��f��
                 '  from ['+@TB_OR_Name+'] m '+@CR+
                 '       left join SYNC_TA13.dbo.sstock d on rtrim(m.sk_no) = rtrim(d.sk_no) '+@CR+
                 '       left join SYNC_TA13.dbo.pcust d1 on rtrim(m.ct_no) = rtrim(d1.ct_no) and ct_class =''1'' '+@CR+					   
                 -- 2017/03/28 Rickliu �]���q�\��ƨS�k�o�򨳳t�Y�ɡA�]���u��^�Y�h�����V��Ʈw
                 '       left join tyeipdbs2.lytdbta13.dbo.sorder d2 on rtrim(m.ct_no) = rtrim(d2.or_ctno) '+@CR+	
                 ' where isfound =''Y'' '+@CR+
                 '   and d2.or_class = '''+@Or_Class+''' ' +@CR+
                 '   and m.F12 COLLATE Chinese_Taiwan_Stroke_CI_AS = rtrim(d2.or_wkno)' +@CR+
                 '   and d2.or_rem like ''%'+@Rm+'%'' ' +@CR+
                 ' Order by 1, 2, 3, m.ct_no,m.source_order,d.sk_no '
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    
    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    set @Msg = '�s�W ['+@Or_Date1+'] �¿ߥD�ɡC'
    print Replace(@Prog_Flag, '###', @Msg)
    --2014/04/22 ���`���ܪ�����T�{�q��A�L���A��f�֡C
    set @strSQL ='insert into '+@TB_EZCat_List+' '+@CR+
                 '([EC_NO], [EC_Name], [sp_date], [sp_date_year], [sp_date_month], [sp_date_day], [sp_date_YM], [sp_date_MD], [sp_date_YMD], '+@CR+
                 ' [Order_NO], [buyer], [cust],[cust_phone], [cust_mobile], [cust_addr1],[cust_addr2],[sk_no],[sk_name],[source_sk_name],[isfound], '+@CR+
                 ' [qty],[Price], [MEMO],[source_file],[import_date],[Rowid] '+@CR+
                 ') '+@CR+
                 'select m.ct_no as EC_NO, '+@CR+
                 '       m.ct_sname as EC_Name, '+@CR+
                 '       '''+@Or_Date1+''' as sp_date, '+@CR+
                 '       substring(replace('''+@Or_Date1+''',''/'',''''),1,4) as sp_date_year, '+@CR+
                 '       substring(replace('''+@Or_Date1+''',''/'',''''),5,2) as sp_date_month, '+@CR+
                 '       substring(replace('''+@Or_Date1+''',''/'',''''),7,2) as sp_date_day, '+@CR+
                 '       substring(replace('''+@Or_Date1+''',''/'',''''),1,6) as sp_date_YM, '+@CR+
                 '       substring(replace('''+@Or_Date1+''',''/'',''''),5,4) as sp_date_MD, '+@CR+
                 '       replace('''+@Or_Date1+''',''/'','''') as sp_date_YMD, '+@CR+
                 '       ltrim(rtrim(replace(m.source_order,'''''''',''''))) as Order_NO, '+@CR+ 
                 '       ltrim(rtrim(replace(m.buyer,'''''''',''''))) as buyer, '+@CR+                        
                 '       ltrim(rtrim(replace(m.cust,'''''''',''''))) as cust, '+@CR+  
                 '       ltrim(rtrim(replace(m.cust_phone,'''''''',''''))) as cust_phone, '+@CR+ 
                 '       ltrim(rtrim(replace(m.cust_mobile,'''''''',''''))) as cust_mobile, '+@CR+ 
                 '       ltrim(rtrim(replace(m.cust_addr1,'''''''',''''))) as cust_addr1, '+@CR+ 
                 '       ltrim(rtrim(replace(m.cust_addr2,'''''''',''''))) as cust_addr2, '+@CR+ 
                 '       m.sk_no, '+@CR+ 
                 '       m.sk_name, '+@CR+
                 '       ltrim(rtrim(replace(m.f5,'''''''',''''))) as source_sk_name, '+@CR+
                 '       m.isfound, '+@CR+
                 '       m.sale_qty, '+@CR+ 
                 '       m.sale_amt, '+@CR+ 
                 '       ltrim(rtrim(replace(m.MEMO,'''''''',''''))) as MEMO, '+@CR+ 
                 '       m.source_file, '+@CR+ 
                 '       import_date=convert(datetime, getdate()), '+@CR+
                 '       Rowid  '+@CR+ 
                 '  from ['+@TB_OR_Name+'] m '+@CR+
                 ' where 1=1 ' +@CR+
                 '   and isSend < 2 ' +@CR+
                 '   and Not Exists '+@CR+
                 '       (select * '+@CR+
                 '          from '+@TB_EZCat_List+' d '+@CR+
                 '         where 1=1 '+@CR+
                 '           and rtrim(m.ct_no)+rtrim(m.source_order) collate Chinese_Taiwan_Stroke_CI_AS = '+@CR+
                 '               rtrim(d.EC_NO)+rtrim(d.Order_NO) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                 '        ) '
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    set @Msg = '�R�� ['+@TB_xls_Name+'] �~�����ɸ�ơA�קK�G����J�C'
    print Replace(@Prog_Flag, '###', @Msg)
    Set @strSQL = 'delete from ['+@TB_xls_Name+']'
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end Try
  begin catch
    Set @Cnt = @Errcode
    Set @Msg = '�i���~ĵ�i�j'+@Proc+'...'+ERROR_MESSAGE()
    print Replace(@Prog_Flag, '###', @Msg)
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

  end catch

proc_exit:
    Set @Msg = @Proc+' �{�ǵ���'
    print Replace(@Prog_Flag, '###', @Msg)

end
GO
