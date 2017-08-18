USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_SCM_Comp_Mall_Consign_Stock_Car1]    Script Date: 08/18/2017 17:43:40 ******/
DROP PROCEDURE [dbo].[uSP_Imp_SCM_Comp_Mall_Consign_Stock_Car1]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_Imp_SCM_Comp_Mall_Consign_Stock_Car1]
as
begin
  /***********************************************************************************************************
     ���R�Ϋ�x�Գ�ɡA�п�ܨ̤��������ơA�p�� XLS �榡�~�|���T
     2013/11/28 �W�[���Ѧ^�ǭ�
  ***********************************************************************************************************/
  Declare @Proc Varchar(50) = 'uSP_Imp_SCM_Comp_Mall_Consign_Stock_Car1'
  Declare @Cnt Int =0
  Declare @RowCnt Int =0
  Declare @Msg Varchar(4000) =''
  Declare @Errcode int = -1

  Declare @Kind Varchar(1), @KindName Varchar(10), @sd_class Varchar(1), @sd_slip_fg Varchar(1), @OD_Name Varchar(20), @OD_CName Varchar(20), @oper Varchar(1)
  Declare @xls_head Varchar(100), @TB_head Varchar(100), @TB_Head_Kind Varchar(100), @TB_xls_Name Varchar(200), @TB_OD_Name Varchar(200), @TB_tmp_name Varchar(200)

  Declare @strSQL Varchar(Max)= ''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @RowCount Table (cnt int)
  Declare @RowData Table (aData Varchar(255))
  
  Declare @new_sdno varchar(10)
  Declare @sd_date varchar(10)
  Declare @sp_date varchar(10)
  Declare @Print_Date varchar(10)
  Declare @Last_Date Varchar(2)
  Declare @sp_cnt int
       
  --���ͳf�����A2013/8/27 �L�ҽT�w�C��T�w�H24�������ɤ�
  set @Last_Date = '24'
  set @sp_date = Substring(Convert(Varchar(10), getdate(), 111), 1, 8)+@Last_Date
  set @sd_date = @sp_date

  Declare @CompanyName Varchar(100), @CompanyLikeCode Varchar(10), @CompanyLikeName Varchar(100), @Rm Varchar(100), @sp_maker Varchar(20), @str Varchar(200)
  Declare @Pos Int
  
  Set @CompanyLikeCode = 'I0002%'
  Set @CompanyName = '���R��'
  Set @CompanyLikeName = '%'+@CompanyName+'%'
  Set @Rm = '�t�ζפJ'+@CompanyName
  Set @sp_maker = 'Admin'

print '1'
  Declare Cur_Car1_Stock_Comp_DataKind cursor for
    select *
      from (select '2' as Kind, '3C' as KindName
             union
            select '3' as Kind, 'Retail' as KindName
           )m
     
print '2'
  open Cur_Car1_Stock_Comp_DataKind
  fetch next from Cur_Car1_Stock_Comp_DataKind into @Kind, @KindName

print '3'
  while @@fetch_status =0
  begin
print '4'
    Set @xls_head = 'Ori_Xls#'
    Set @TB_head = 'Comp_Mall_Consign_Stock_Car1'
    Set @TB_head_Kind = @TB_head+'_'+@KindName
    Set @TB_xls_Name = @xls_head+@TB_Head_Kind -- Ex: Ori_Xls#Comp_Mall_Consign_Stock_Car1_3C -- Excel ���ɸ��
    Set @TB_tmp_Name = @TB_Head_Kind+'_tmp' -- Ex: Comp_Mall_Consign_Stock_Car1_3C_tmp -- �{����J����
    Set @TB_OD_Name = @TB_Head_Kind
    
    -- Check �פJ�ɮ׬O�_�s�b
    IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_xls_Name+']') AND type in (N'U'))
    begin
print '5'
       --Send Message
       set @Cnt = @Errcode
       set @strSQL = ''
       Set @Msg = '�~�� Excel �פJ��ƪ� ['+@TB_xls_Name+']���s�b�A�פ�i�����ɧ@�~�C'
       Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
       -- 2013/11/28 �W�[���Ѧ^�ǭ�
       close Cur_Car1_Stock_Comp_DataKind
       deallocate Cur_Car1_Stock_Comp_DataKind
       Return(@Errcode)
       
--       fetch next from Cur_Car1_Stock_Comp_DataKind into @Kind, @KindName
--       Continue
    end
      
    -- �P�O Excel �C�L��� �O�_�s�b
print '6'
    set @Cnt = 0
    set @Print_Date  = ''
    
    set @strSQL = 'Select Replace(F1, ''�C�L����G'', '''') as Print_date from [dbo].['+@TB_xls_Name+'] where F1 like ''�C�L���%'' '
    set @Msg = '�P�O Excel �C�L��� �O�_�s�b'
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
    
    delete @RowData
    insert into @RowData exec (@strSQL)
    select @Print_Date=Rtrim(isnull(aData, '')) from @RowData
    set @Print_Date = Convert(Varchar(10), CONVERT(Date, @Print_Date) , 111)
    
    --2013/11/25 �L�ҽT�w�C��T�w�H24�������ɤ� 
    set @Print_Date = Substring(Convert(Varchar(10), @Print_Date, 111), 1, 8)+@Last_Date
    
    if Rtrim(@Print_Date) = ''
    begin
print '4'
       set @Cnt = @Errcode
       set @strSQL = ''
       Set @Msg = '�䤣�� Excel ��Ƥ����C�L����A�פ�i�����ɧ@�~�C'
       Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
       -- 2013/11/28 �W�[���Ѧ^�ǭ�
       close Cur_Car1_Stock_Comp_DataKind
       deallocate Cur_Car1_Stock_Comp_DataKind
       Return(@Errcode)

--       fetch next from Cur_Car1_Stock_Comp_DataKind into @Kind, @KindName
--       Continue
    end
    else
    begin
/*    
print '7'
      IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_OD_Name+']') AND type in (N'U'))
      begin
         Set @strSQL = 'select Count(*) as cnt from '+@TB_OD_Name+' where print_date ='''+@Print_Date+''' '
         delete @RowCount
         insert into @RowCount Exec(@strSQL)
    
         if (select cnt from @RowCount) > 0
         begin
print '8'
          set @Msg = '�̪�@���C�L����ۦP�h���i�����ɧ@�~!!'
          set @Cnt = @Errcode
          Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
          -- 2013/11/28 �W�[���Ѧ^�ǭ�
          fetch next from Cur_Car1_Stock_Comp_DataKind into @Kind, @KindName
          Continue
         end
      end
*/

print '9'
      IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_tmp_Name+']') AND type in (N'U'))
      begin
print '10'
         Set @Msg = '�M��'+@CompanyName+'�{����J������ƪ� ['+@TB_tmp_Name+']�C'
    
         Set @strSQL = 'DROP TABLE [dbo].['+@TB_tmp_Name+']'
         Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
      end
    
      --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
      --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
      --�N�C����ƥ[�J�ߤ@���(ps.����ƪ�ФűƧǡA�O�d��l XLS �˻��A�H�Q��Უ�X��b��)
      --�إ߼Ȧs�ɥB�إߧǸ��A�إߧǸ��B�J�ܭ��n�A�]������n�̧ǲ��ͩ��W
print '11'
      Set @Msg = '�إ��{����J������ƪ�C���ߤ@�Ǹ� ['+@TB_tmp_Name+']�C'
      Set @strSQL = 'select *, print_date=Convert(date, rtrim('''+@Print_Date+''')) '+@CR+
                    '       into '+@TB_tmp_Name+@CR+
                    '  from '+@TB_xls_Name
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    
      --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '12'
      Set @Msg = '�R���{����J������ƪ����Y��� ['+@TB_tmp_Name+']�C'
      Set @strSQL = 'delete '+@TB_tmp_Name+@CR+
                    ' where rowid <= 2 '
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    
      --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '13'
      Set @Msg = '�B�z�{����J������ƪ��s�դp�p���W��� ['+@TB_tmp_Name+']�C'
      Set @strSQL = 'declare cur_book cursor for '+@CR+
                    ' select rowid, f1 '+@CR+
                    '   from '+@TB_tmp_Name+@CR+
                    '  order by rowid '+@CR+
                    ''+@CR+
                    'declare @rowid int, @cnt int '+@CR+
                    'declare @f1 varchar(100), @vf1 varchar(100) '+@CR+
                    'set @cnt=1 '+@CR+
                    ''+@CR+
                    'open cur_book '+@CR+
                    'fetch next from cur_book into @rowid, @f1 '+@CR+
                    ''+@CR+
                    'set @cnt=1 '+@CR+
                    'while @@fetch_status =0 '+@CR+
                    'begin '+@CR+
                    ''+@CR+
                    '  if @f1 <> '''' '+@CR+
                    '     set @vf1 = @f1 '+@CR+
                    '  else '+@CR+
                    '  begin '+@CR+
                    '     update '+@TB_tmp_Name+@CR+
                    '      set f1 = @vf1 '+@CR+
                    '    where rowid = @rowid '+@CR+
                    '  end '+@CR+
                    '  fetch next from cur_book into @rowid, @f1 '+@CR+
                    'end '+@CR+
                    'Close cur_book '+@CR+
                    'Deallocate cur_book '
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
        
      --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
      -- �h���{����J������ƪ����W�Ǹ�


-- 2015/01/12 Rickliu ���W�e��X���Ʀr�A�@�ߥh������    
/*
print '14'
      Set @str = ' 1234567890'
      set @Pos = 1
      While @Pos < Len(@str)  
      begin
print '15'
         Set @Msg = '�h�� F1 ��줺�t���Ʀr['+Substring(@str, @Pos, 1)+']����ର�ť� ['+@TB_tmp_Name+']'
    
         Set @strSQL = 'update '+@TB_tmp_Name+' set f1=replace(f1, '''+Substring(@str, @Pos, 1)+''', ''''), f7=replace(f7, '','', ''''), f17=replace(f17, '','', '''') '
         Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
        
         Set @Pos = @Pos +1
      end
*/    
      --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '16'
      Set @Msg = '�h�� F1, F10~F17 ��줺�t��"��" "," �r�����ť� ['+@TB_tmp_Name+']'
      --Set @strSQL = 'update '+@TB_tmp_Name+' set f1 = substring(replace(f1, ''��'', ''''), 1, 2) '
      Set @strSQL = 'update '+@TB_tmp_Name+@CR+
                    ' set F1 = substring(replace(F1, ''��'', ''''), 3, Len(F1)), '+@CR+
                    '     F10= replace(F10, '','', ''''), '+@CR+
                    '     F11= replace(F11, '','', ''''), '+@CR+
                    '     F12= replace(F12, '','', ''''), '+@CR+
                    '     F13= replace(F13, '','', ''''), '+@CR+
                    '     F14= replace(F14, '','', ''''), '+@CR+
                    '     F15= replace(F15, '','', ''''), '+@CR+
                    '     F16= replace(F16, '','', ''''), '+@CR+
                    '     F17= replace(F17, '','', '''') '
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
         
      --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '15'
      Set @Msg = '�R�� F3, F10, F13 ��쬰�ťո�� ['+@TB_tmp_Name+']'
    
      Set @strSQL = 'delete '+@TB_tmp_Name+@CR+
                    ' where (isnull(f3, '''') ='''' '+@CR+
                    '   and isnull(f10, '''') ='''' '+@CR+
                    '   and isnull(f13, '''') ='''') '+@CR+
                    '    Or (F3 Like ''%�ӫ~���X%'') '
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
         
print '17'
      IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_OD_Name+']') AND type in (N'U'))
      begin
print '18'
         --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
         Set @Msg = '����'+@TB_OD_Name+'��b�`�� ['+@TB_head+']�C'
    
         Set @strSQL = 'Create Table '+@TB_OD_Name+' ('+@CR+
                       '  	[Kind1] [varchar](1) NOT NULL, '+@CR+
                       '    [rowid] [int] NOT NULL, '+@CR+
	                   '    [ct_no] [varchar](50) NULL, '+@CR+
	                   '    [ct_sname] [varchar](50) NULL, '+@CR+
	                   '    [ct_ssname] [varchar](50) NULL, '+@CR+
	                   '    [sk_no] [varchar](50) NULL, '+@CR+
	                   '    [sk_name] [varchar](50) NULL, '+@CR+
	                   '    [F7] [varchar](50) NULL, '+@CR+
	                   '    [sk_bcode] [varchar](50) NULL, '+@CR+
	                   '    [F3] [varchar](50) NULL, '+@CR+
	                   '    [fg6_qty] [numeric](38, 6) NOT NULL, '+@CR+
	                   '    [fg7_qty] [numeric](38, 6) NOT NULL, '+@CR+
	                   '    [sum_qty] [numeric](38, 6) NOT NULL, '+@CR+
	                   '    [fg6_amt] [numeric](38, 6) NOT NULL, '+@CR+
	                   '    [fg7_amt] [numeric](38, 6) NOT NULL, '+@CR+
	                   '    [sum_amt] [numeric](38, 6) NOT NULL, '+@CR+
	                   '    [F13] [numeric](20, 6) NOT NULL, '+@CR+
	                   '    [F17] [numeric](20, 6) NOT NULL, '+@CR+
	                   '    [diff_qty] [numeric](38, 6) NOT NULL, '+@CR+
	                   '    [diff_amt] [numeric](38, 6) NOT NULL, '+@CR+
	                   '    [Print_Date] [varchar](10) NOT NULL, '+@CR+
	                   '    [isfound] [varchar](100) NOT NULL, '+@CR+
	                   '    [Exec_DateTime] [datetime] NOT NULL '+@CR+
                       ') ON [PRIMARY] '
         Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
      end

      Set @Msg = '�R����ƪ� ['+@TB_OD_Name+']'
      Set @strSQL = 'Delete [dbo].['+@TB_OD_Name+'] '+@CR+
                    ' where Print_Date = '''+@Print_Date+''' '
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
         
      --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
      --2015/01/13 ���`���ܭn���X���R�ΰӫ~���`��b��A����Ȩ̰ӫ~�i���e���e�{���G�AKind1: 1 ���ӫ~���`�`��, 2 ����+�ӫ~�J�`�`��
      
print '19'
      Set @Msg = '�i����ӫ~�s���üg�J ['+@TB_OD_Name+']��ƪ�C'
      set @strSQL = 
                'Insert Into '+@TB_OD_Name+' '+@CR+
                '(Kind1, rowid, ct_no, ct_sname, ct_ssname, sk_no, '+@CR+
                ' sk_name, f7, sk_bcode, f3, fg6_qty, '+@CR+
                ' fg7_qty, sum_qty, fg6_amt, fg7_amt, sum_amt, '+@CR+
                ' F13, F17, diff_qty, diff_amt, Print_Date, '+@CR+
                ' isfound, Exec_DateTime) '+@CR+

                'select Distinct '+@CR+
                '       ''1'' as Kind1, '+@CR+
                '       0 as rowid, '+@CR+
                '       Convert(Varchar(50), '''') as ct_no, '+@CR+
                '       Convert(Varchar(50), '''') as ct_sname, '+@CR+
                '       Convert(Varchar(50), '''') as ct_ssname, '+@CR+
                '       Convert(Varchar(50), RTrim(isnull(isnull(d.co_skno, d1.sk_no), ''N/A''))) as sk_no, '+@CR+
                -- �ӫ~�W��
                '       Convert(Varchar(50), RTrim(isnull(isnull(d.sk_name, d1.sk_name), ''N/A''))) as sk_name, '+@CR+
                -- �ӫ~�W��(Excel)
                '       Convert(Varchar(50), RTrim(F7)) as F7, '+@CR+ 
                -- �ӫ~���X
                '       Convert(Varchar(50), RTrim(isnull(isnull(d.sk_bcode, d1.sk_bcode), ''N/A''))) as sk_bcode, '+@CR+
                -- �ӫ~���X(Excel)
                '       Convert(Varchar(50), RTrim(F3)) as F3, '+@CR+ 
                -- ����ƶq(��V)
                '       isnull(fg6_qty, 0) as fg6_qty, '+@CR+ 
                -- �^�f�ƶq(��V)
                '       isnull(fg7_qty, 0) as fg7_qty, '+@CR+ 
                -- �X�p�ƶq(��V)
                '       isnull(sum_qty, 0 ) as sum_qty, '+@CR+ 
                -- ������B(��V)
                '       isnull(fg6_amt, 0) as fg6_amt, '+@CR+ 
                -- �^�f���B(��V)
                '       isnull(fg7_amt, 0) as fg7_amt, '+@CR+ 
                -- �X�p���B(��V)
                '       isnull(sum_amt, 0) as sum_amt, '+@CR+ 
                -- �H��w�s (Excel) 
                '       isnull(Convert(Numeric(20, 6), F13), 0) as F13, '+@CR+
                -- �H����B(Excel)
                '       isnull(Convert(Numeric(20, 6), F17), 0) as F17, '+@CR+
                -- �t���ƶq
                '       isnull(sum_qty - Convert(Numeric(20, 6), F13), 0) as diff_qty, '+@CR+
                -- �t�����B
                '       isnull(sum_amt - Convert(Numeric(20, 6), F17), 0) as diff_amt, '+@CR+
                '       '''+@Print_Date+''' as Print_Date, '+@CR+
                -- ���`
                '       case '+@CR+
                '         when (isnull(d.co_skno, d1.sk_no) is null) then ''XLS �s���藍���V�ӫ~���'' '+@CR+
                '         when (fg6_qty is null) then ''XLS ����ơA����V�L�U��^�f���'' '+@CR+
                '         when (F3 is null and sum_qty <>0) then ''��V���U��^�f��ơA��XLS�L���'' '+@CR+
                '         when (sum_qty - F13 <>0) then ''���ƶq���`'' '+@CR+
                '         else '''' '+@CR+
                '       end as isfound, '+@CR+
                '       getdate() as Exec_DateTime '+@CR+
                -- �Ȥ�s��
                '  from (Select F3, F7, '+@CR+
                '               Sum(isnull(Convert(Numeric(20, 6), Replace(F13, '','', '''')), 0)) as F13, '+@CR+
                '               Sum(isnull(Convert(Numeric(20, 6), Replace(F17, '','', '''')), 0)) as F17 '+@CR+
                '          from '+@TB_tmp_Name+@CR+
                '         Group by F3, F7 '+@CR+
                '       )m '+@CR+
                -- �ӫ~���
                '       Full join '+@CR+ 
                '       (select distinct '+@CR+ 
                '               RTrim(co_ctno) as co_ctno, '+@CR+ 
                '               RTrim(co_skno) as co_skno, '+@CR+  
                '               RTrim(co_cono) as co_cono, '+@CR+  
                '               RTrim(sk_no) as sk_no, '+@CR+  
                '               RTrim(sk_name) as sk_name, '+@CR+  
                '               RTrim(sk_bcode) as sk_bcode '+@CR+ 
                '          from SYNC_TA13.dbo.sauxf m '+@CR+  
                '               left join SYNC_TA13.dbo.sstock d '+@CR+  
                '                 on m.co_skno = d.sk_no '+@CR+  
                '                and m.co_ctno like '''+@CompanyLikeCode+''' '+@CR+ 
                '         where 1=1 '+@CR+ 
                '           and co_class=''1'' '+@CR+  
                -- �u�d�U��^�f����
                '           and exists '+@CR+ 
                '               (select * '+@CR+ 
                '                  from SYNC_TA13.dbo.sslpdt d1 '+@CR+ 
                '                 where 1=1 '+@CR+ 
                '                   and sd_slip_fg IN (''6'',''7'') '+@CR+   
                '                   and d.sk_no = d1.sd_skno '+@CR+ 
                '                   and d1.sd_ctno like '''+@CompanyLikeCode+''' '+@CR+ 
                '                   and d1.sd_date <= '''+@Print_Date+'''  '+@CR+
                '               ) '+@CR+
                '         ) d '+@CR+
                '         on 1=1 '+@CR+
                '        and (ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.co_skno)) collate Chinese_Taiwan_Stroke_CI_AS  '+@CR+
                '         or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.co_cono)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                '         or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.sk_no)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                '         or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.sk_bcode)) collate Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                '       left join SYNC_TA13.dbo.sstock d1 '+@CR+
                '         on 1=1 '+@CR+
                '        and (ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d1.sk_no)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                '         or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d1.sk_bcode)) collate Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                --  ��V����^�f��� sp_slip_fg = 6 ����, sp_slip_fg = 7 �^�f
                '       left join  '+@CR+
                '         (select sd_skno,  '+@CR+
                '                 sum(case when sd_slip_fg = ''6'' then sd_qty else 0 end) as fg6_qty, '+@CR+ -- ����ƶq
                '                 sum(case when sd_slip_fg = ''7'' then sd_qty else 0 end) as fg7_qty, '+@CR+ -- �^�f�ƶq
                '                 sum(case when sd_slip_fg = ''7'' then sd_qty * -1 else sd_qty end) as sum_qty, '+@CR+ -- �X�p�ƶq<-------CHECK
                '                 sum(case when sd_slip_fg = ''6'' then sd_stot else 0 end) as fg6_amt, '+@CR+ -- ������B
                '                 sum(case when sd_slip_fg = ''7'' then sd_stot else 0 end) as fg7_amt, '+@CR+ -- ����ƶq
                '                 sum(case when sd_slip_fg = ''7'' then sd_stot * -1 else sd_stot end) as sum_amt '+@CR+ -- �X�p���B
                '            from SYNC_TA13.dbo.sslpdt '+@CR+
                '           where 1=1 '+@CR+
                '             and sd_slip_fg IN (''6'',''7'')  '+@CR+
                '             and sd_date >= ''2012/12/01''  '+@CR+
                -- �d��I���� 2013/01/01 ~ �C�L���
                '             and sd_date <= '''+@Print_Date+'''  '+@CR+
                '             and sd_ctno like '''+@CompanyLikeCode+''' '+@CR+ 
                '           group by sd_skno '+@CR+
                '         ) d2 '+@CR+
                '         on 1=1 '+@CR+
                '        and isnull(d.co_skno, d1.sk_no) = d2.sd_skno '+@CR+
                ' where 1=1 '+@CR+
                -- ����ܤ��e���ʹL���ӫ~�B�ƶq�� 0 ���ӫ~, �]�N�O�u��ܳѾl�w�s����
                --'   and (F3 is not null) '+@CR+
                --'   and (isnull(sum_qty, 0) <> 0) '+@CR+
                ' order by 1, 2 '
                --group by ct_no, ct_sname, ct_ssname, sk_no, sk_name, sk_bcode, f3, f6
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

      --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
      --2013/04/25 ���@���v����J�N�n���Υ[�`�A����n��
      --�b���i����ӫ~�s���A�|�����U�s���H�ΰӫ~�򥻸���ɶi����@�~
print '17'
      Set @Msg = '�i����ӫ~�s���üg�J ['+@TB_OD_Name+']��ƪ�C'
      set @strSQL = 
                'Insert Into '+@TB_OD_Name+' '+@CR+
                '(Kind1, rowid, ct_no, ct_sname, ct_ssname, sk_no, '+@CR+
                ' sk_name, f7, sk_bcode, f3, fg6_qty, '+@CR+
                ' fg7_qty, sum_qty, fg6_amt, fg7_amt, sum_amt, '+@CR+
                ' F13, F17, diff_qty, diff_amt, Print_Date, '+@CR+
                ' isfound, Exec_DateTime) '+@CR+
                
                'select Distinct '+@CR+
                '       ''2'' as Kind1, rowid, '+@CR+
                '       Convert(Varchar(50), RTrim(ct_no)) as ct_no, '+@CR+
                '       Convert(Varchar(50), RTrim(ct_sname)) as ct_sname, '+@CR+
                '       Convert(Varchar(50), RTrim(ct_ssname)) as ct_ssname, '+@CR+
                '       Convert(Varchar(50), RTrim(isnull(isnull(d.co_skno, d1.sk_no), ''N/A''))) as sk_no, '+@CR+
                -- �ӫ~�W��
                '       Convert(Varchar(50), RTrim(isnull(isnull(d.sk_name, d1.sk_name), ''N/A''))) as sk_name, '+@CR+
                -- �ӫ~�W��(Excel)
                '       Convert(Varchar(50), RTrim(F7)) as F7, '+@CR+ 
                -- �ӫ~���X
                '       Convert(Varchar(50), RTrim(isnull(isnull(d.sk_bcode, d1.sk_bcode), ''N/A''))) as sk_bcode, '+@CR+
                -- �ӫ~���X(Excel)
                '       Convert(Varchar(50), RTrim(F3)) as F3, '+@CR+ 
                -- ����ƶq(��V)
                '       isnull(fg6_qty, 0) as fg6_qty, '+@CR+ 
                -- �^�f�ƶq(��V)
                '       isnull(fg7_qty, 0) as fg7_qty, '+@CR+ 
                -- �X�p�ƶq(��V)
                '       isnull(sum_qty, 0 ) as sum_qty, '+@CR+ 
                -- ������B(��V)
                '       isnull(fg6_amt, 0) as fg6_amt, '+@CR+ 
                -- �^�f���B(��V)
                '       isnull(fg7_amt, 0) as fg7_amt, '+@CR+ 
                -- �X�p���B(��V)
                '       isnull(sum_amt, 0) as sum_amt, '+@CR+ 
                -- �H��w�s (Excel) 
                '       isnull(Convert(Numeric(20, 6), F13), 0) as F13, '+@CR+
                -- �H����B(Excel)
                '       isnull(Convert(Numeric(20, 6), F17), 0) as F17, '+@CR+
                -- �t���ƶq
                '       isnull(sum_qty - Convert(Numeric(20, 6), F13), 0) as diff_qty, '+@CR+
                -- �t�����B
                '       isnull(sum_amt - Convert(Numeric(20, 6), F17), 0) as diff_amt, '+@CR+
                '       '''+@Print_Date+''' as Print_Date, '+@CR+
                -- ���`
                '       case '+@CR+
                '         when (isnull(d.co_skno, d1.sk_no) is null) then ''�s���藍��'' '+@CR+
                '         when (fg6_qty is null) then ''�L��V�P�����'' '+@CR+
                '         else '''' '+@CR+
                '       end as isfound, '+@CR+
                '       getdate() as Exec_DateTime '+@CR+
                -- �Ȥ�s��
                '  from (select * '+@CR+
                '          from (select ct_no, rtrim(ct_sname) as ct_sname, rtrim(ct_name)+''#''+rtrim(ct_sname) as ct_ssname '+@CR+
                '                  from SYNC_TA13.dbo.PCUST  '+@CR+
                '                 where ct_class =''1''  '+@CR+
                '                   and ct_name like '''+@CompanyLikeName+''' '+@CR+
                '                   and substring(ct_no, 9, 1) ='''+@Kind+''' '+@CR+ -- 3C �� 2�A �ʳf�� 3
                '                   and ct_no like '''+@CompanyLikeCode+''' '+@CR+ 
                '                ) m, '+@TB_tmp_Name+' d '+@CR+
                '         where m.ct_ssname collate Chinese_Taiwan_Stroke_CI_AS like ''%''+rtrim(d.f1)+''%'' collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                '       )m '+@CR+
                -- �ӫ~���
                '       left join '+@CR+ 
                '       (select distinct '+@CR+ 
                '               RTrim(co_ctno) as co_ctno, '+@CR+ 
                '               RTrim(co_skno) as co_skno, '+@CR+  
                '               RTrim(co_cono) as co_cono, '+@CR+  
                '               RTrim(sk_no) as sk_no, '+@CR+  
                '               RTrim(sk_name) as sk_name, '+@CR+  
                '               RTrim(sk_bcode) as sk_bcode '+@CR+ 
                '          from SYNC_TA13.dbo.sauxf m '+@CR+  
                '               left join SYNC_TA13.dbo.sstock d '+@CR+  
                '                 on m.co_skno = d.sk_no '+@CR+  
                '                and m.co_ctno like '''+@CompanyLikeCode+''' '+@CR+ 
                '         where 1=1 '+@CR+ 
                '           and co_class=''1'' '+@CR+  
                -- �u�d�U��^�f����
                '           and exists '+@CR+ 
                '               (select * '+@CR+ 
                '                  from SYNC_TA13.dbo.sslpdt d1 '+@CR+ 
                '                 where 1=1 '+@CR+ 
                '                   and sd_slip_fg IN (''6'',''7'') '+@CR+   
                '                   and d.sk_no = d1.sd_skno '+@CR+ 
                '                   and d1.sd_ctno like '''+@CompanyLikeCode+''' '+@CR+ 
                '                   and d1.sd_date <= '''+@Print_Date+'''  '+@CR+
                '               ) '+@CR+
                '         ) d '+@CR+
                '         on 1=1 '+@CR+
                '        and m.ct_no=d.co_ctno '+@CR+
                '        and (ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.co_skno)) collate Chinese_Taiwan_Stroke_CI_AS  '+@CR+
                '         or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.co_cono)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                '         or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.sk_no)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                '         or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.sk_bcode)) collate Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                '       left join SYNC_TA13.dbo.sstock d1 '+@CR+
                '         on 1=1 '+@CR+
                '        and (ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d1.sk_no)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                '         or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d1.sk_bcode)) collate Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                --  ��V����^�f��� sp_slip_fg = 6 ����, sp_slip_fg = 7 �^�f
                '       left join  '+@CR+
                '         (select sd_ctno, sd_skno,  '+@CR+
                '                 sum(case when sd_slip_fg = ''6'' then sd_qty else 0 end) as fg6_qty, '+@CR+ -- ����ƶq
                '                 sum(case when sd_slip_fg = ''7'' then sd_qty else 0 end) as fg7_qty, '+@CR+ -- �^�f�ƶq
                '                 sum(case when sd_slip_fg = ''7'' then sd_qty * -1 else sd_qty end) as sum_qty, '+@CR+ -- �X�p�ƶq<-------CHECK
                '                 sum(case when sd_slip_fg = ''6'' then sd_stot else 0 end) as fg6_amt, '+@CR+ -- ������B
                '                 sum(case when sd_slip_fg = ''7'' then sd_stot else 0 end) as fg7_amt, '+@CR+ -- ����ƶq
                '                 sum(case when sd_slip_fg = ''7'' then sd_stot * -1 else sd_stot end) as sum_amt '+@CR+ -- �X�p���B
                '            from SYNC_TA13.dbo.sslpdt '+@CR+
                '           where 1=1 '+@CR+
                '             and sd_slip_fg IN (''6'',''7'')  '+@CR+
                '             and sd_date >= ''2012/12/01''  '+@CR+
                -- �d��I���� 2013/01/01 ~ �C�L���
                '             and sd_date <= '''+@Print_Date+'''  '+@CR+
                '             and sd_ctno like '''+@CompanyLikeCode+''' '+@CR+ 
                '           group by sd_ctno, sd_skno '+@CR+
                '         ) d2 '+@CR+
                '         on 1=1 '+@CR+
                '        and isnull(d.co_skno, d1.sk_no) = d2.sd_skno '+@CR+
                '        and ct_no = d2.sd_ctno '+@CR+
                ' order by 1, 2 '
                --group by ct_no, ct_sname, ct_ssname, sk_no, sk_name, sk_bcode, f3, f6
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    
      --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
      --2013/12/24�N���ƪ���Ƥ]�C�J���`
print '20'
      set @Msg = '�N���ƪ���Ƽе����`'
      set @Cnt = 0
      set @strSQL ='update '+@TB_OD_Name+@CR+
                   '   set isfound = isfound+''�A���Ƥ��'' '+@CR+
                   ' where 1=1 '+@CR+
                   '   and Kind1 = ''2'' '+@CR+
                   '   and rowid in '+@CR+
                   '       (select rowid '+@CR+
                   '          from '+@TB_OD_Name+@CR+
                   '         where print_date = '''+@Print_Date+''' '+@CR+
                   '         group by rowid '+@CR+
                   '        having COUNT(*) > 1) '
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
/*    
print '21'
      --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
      --2014/11/20 Rickliu ���͹�b�`��
      set @Msg = '���͹�b�`��'
      set @Cnt = 0
      set @strSQL ='update '+@TB_OD_Name+@CR+
                   '   set isfound = ''���Ƥ��'' '+@CR+
                   ' where rowid in '+@CR+
                   '       (select rowid '+@CR+
                   '          from '+@TB_OD_Name+@CR+
                   '         where print_date = '''+@Print_Date+''' '+@CR+
                   '         group by rowid '+@CR+
                   '        having COUNT(*) > 1) '
*/    
      
      
      IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_head+']') AND type in (N'U'))
      begin
print '22'
         Set @Msg = '����'+@CompanyName+'��b�`�� ['+@TB_head+']�C'
    
         Set @strSQL = 'Create Table '+@TB_head+' ('+@CR+
                       '  [Kind1] [varchar](10) NULL, '+@CR+
                       '  [Kind2] [varchar](10) NULL, '+@CR+
                       '  [sd_ctno] [varchar](50) NULL, '+@CR+
                       '  [ct_sname] [nchar](50) NULL, '+@CR+
                       '  [sk_no] [varchar](30) NULL, '+@CR+
                       '  [sk_name] [nchar](50) NULL, '+@CR+
                       '  [xls_skno] [varchar](50) NOT NULL, '+@CR+
                       '  [xls_skname] [varchar](50) NOT NULL, '+@CR+
                       '  [sk_bcode] [varchar](50) NULL, '+@CR+
                       '  [xls_bcode] [varchar](50) NOT NULL, '+@CR+
                       '  [rowid] [int] NULL, '+@CR+
                       '  [Chg_sd_Qty] [numeric](18, 2) NULL, '+@CR+
                       '  [Chg_sd_stot] [numeric](18, 2) NULL, '+@CR+
                       '  [xls_qty] [numeric](18, 2) NOT NULL, '+@CR+
                       '  [xls_amt] [numeric](18, 2) NOT NULL, '+@CR+
                       '  [diff_qty] [numeric](18, 2) NULL, '+@CR+
                       '  [diff_amt] [numeric](18, 2) NULL, '+@CR+
                       '  [Print_Date] [varchar](10) NULL, '+@CR+
                       '  [isfound] [varchar](20) NOT NULL, '+@CR+
                       '  [Exec_DateTime] [DateTime] NULL '+@CR+
                       ') ON [PRIMARY] '
         Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
      end
      
      Set @Msg = '�M��'+@CompanyName+'��b�`�� ['+@TB_head+']�C'
      set @strSQL ='Delete '+@TB_head+' '+@CR+
                   ' where Kind2 = '''+@KindName+''' '+@CR+
                   '   and Print_Date = '''+@Print_Date+''' '                   
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
	  
      Set @Msg = '�g�J'+@CompanyName+'��b�`�� ['+@TB_head+']�C'
      set @strSQL ='Insert Into '+@TB_head+' '+@CR+
                   '(Kind1, Kind2, sd_ctno, ct_sname, sk_no, sk_name, '+@CR+
                   ' xls_skno, xls_skname, sk_bcode, xls_bcode, rowid, '+@CR+
                   ' Chg_sd_Qty, Chg_sd_stot, xls_qty, xls_amt, diff_qty, '+@CR+
                   ' diff_amt, print_date, isfound, Exec_DateTime) '+@CR+
                   'select distinct '+@CR+
                   '       Kind1, '+@CR+ -- ��况�a+�ӫ~
                   '       '''+@KindName+''' as Kind2, '+@CR+
                   '       m.sd_ctno, m.ct_sname, '+@CR+
                   '       m.sd_skno as sk_no, d1.sk_name as sk_name,  '+@CR+
                   '       isnull(d.f3, ''N/A'') as xls_skno, '+@CR+
                   '       isnull(d.f7, ''N/A'') as xls_skname, '+@CR+
                   '       m.sk_bcode, '+@CR+
                   '       isnull(d.f3, ''N/A'') as xls_bcode, '+@CR+
                   '       d.rowid, '+@CR+
                   '       m.Chg_sd_Qty, '+@CR+
                   '       m.Chg_sd_stot, '+@CR+
                   '       isnull(d.F13, 0) as xls_qty, '+@CR+
                   '       isnull(d.F17, 0) as xls_amt, '+@CR+
                   '       isnull(d.F13, 0) - isnull(m.Chg_sd_Qty, 0)as diff_qty, '+@CR+
                   '       isnull(d.F17, 0) - isnull(m.Chg_sd_stot, 0) as diff_amt, '+@CR+
                   '       '''+@Print_Date+''' as Print_Date, '+@CR+
                   --'       (select top 1 isnull(print_date, convert(varchar(10), getdate(), 111)) from '+@TB_head_Kind+')as Print_Date, '+@CR+
                   '       isnull(d.isfound, ''�DXLS���'') as isfound, '+@CR+
                   '       GetDate() as Exec_DateTime '+@CR+
                   '  from (select sd_ctno, ct_sname, sd_skno, sk_bcode, '+@CR+
                   '               Sum(Chg_sd_qty) as Chg_sd_Qty, Sum(Chg_sd_stot) as Chg_sd_stot '+@CR+
                   '          from SYNC_TA13.dbo.v_orders m '+@CR+
                   '               left join SYNC_TA13.dbo.pcust d '+@CR+
                   '                 on m.sp_ctno=d.ct_no and d.ct_class=''1'' '+@CR+
                   '         where sd_slip_fg IN (''6'',''7'') '+@CR+
                   '           and sd_date >= ''2012/12/01'' '+@CR+
                   '           and sd_date <= '''+@Print_Date+''' '+@CR+
                   /*
                   '               (select top 1 max(isnull(print_date, convert(varchar(10), getdate(), 111))) '+@CR+
                   '                  from '+@TB_head_Kind+' '+@CR+
                   '                 where isnull(print_date, '''') <> '''')  '+@CR+
                   */
                   '            group by sd_ctno, ct_sname, sd_skno, sk_bcode '+@CR+
                   '       ) m '+@CR+
                   '      LEFT join '+@TB_head_Kind+' d '+@CR+
                   '        on 1=1 '+@CR+
                   '       and m.sd_skno collate Chinese_Taiwan_Stroke_CI_AS = d.sk_no '+@CR+
                   '       and m.sd_ctno collate Chinese_Taiwan_Stroke_CI_AS = d.ct_no '+@CR+
                   '       and d.print_date collate Chinese_Taiwan_Stroke_CI_AS ='''+@Print_Date+''' '+@CR+
                   '      left join SYNC_TA13.dbo.sstock d1 '+@CR+
                   '        on m.sd_skno collate Chinese_Taiwan_Stroke_CI_AS = d1.sk_no '+@CR+
                   ' Where 1=1 '+@CR+
                   '   and m.ct_sname like ''%'+@CompanyName+'%'' '+@CR+
                   '   and substring(m.sd_ctno, len(m.sd_ctno), 1) = '''+@Kind+''' '+@CR+
                   ' order by d.rowid, m.sd_ctno, m.sd_skno '
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

    end
print '23'
    fetch next from Cur_Car1_Stock_Comp_DataKind into @Kind, @KindName
  end
  close Cur_Car1_Stock_Comp_DataKind
  deallocate Cur_Car1_Stock_Comp_DataKind
  -- 2013/11/28 �W�[���Ѧ^�ǭ�
  Return(0)
end
GO
