USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_SCM_Comp_EC_Consign_Stock_MOMO]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_Imp_SCM_Comp_EC_Consign_Stock_MOMO]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_Imp_SCM_Comp_EC_Consign_Stock_MOMO]
as
begin
  /***********************************************************************************************************
   �{���׭q���{�G
   1.2016/01/01 Rickliu �s�W EC MOMO �H�ܮw�s��b�{��
   2.2017/06/22 MOMO �ܧ��b EXCEL �榡, Rickliu �׭q���{��
  ***********************************************************************************************************/
  Declare @Proc Varchar(50) = 'uSP_Imp_SCM_Comp_EC_Consign_Stock_MOMO'
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
  
  Set @CompanyLikeCode = 'I90120011'
  Set @CompanyName = 'MOMO'
  Set @CompanyLikeName = '%'+@CompanyName+'%'
  Set @Rm = '�t�ζפJ'+@CompanyName
  Set @sp_maker = 'Admin'

print '4'
  Set @xls_head = 'Ori_Xls#'
  Set @TB_head = 'Comp_EC_Consign_Stock_'+@CompanyName
  Set @TB_xls_Name = Isnull(@xls_head+@TB_Head, '') -- Ex: Ori_Xls#Comp_EC_Consign_Stock__MOMO -- Excel ���ɸ��
  Set @TB_tmp_Name = Isnull(@TB_Head, '')+'_tmp' -- Ex: Comp_EC_Consign_Stock_MOMO_tmp -- �{����J����
  Set @TB_OD_Name = Isnull(@TB_Head, '')
    
  -- Check �פJ�ɮ׬O�_�s�b
  print @TB_xls_Name
  IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_xls_Name+']') AND type in (N'U'))
  begin
print '5'
     --Send Message
     Set @Msg = '�~�� Excel �פJ��ƪ� ['+@TB_xls_Name+']���s�b�A�פ�i�����ɧ@�~�C'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, @Errcode
     -- 2013/11/28 �W�[���Ѧ^�ǭ�
     Return(@Errcode)
  end
      
    -- �P�O Excel �C�L��� �O�_�s�b
print '6'
  set @Cnt = 0
  set @Print_Date  = ''
    
  set @strSQL = 'Select Top 1 SplitFileName as Print_Date from [dbo].['+@TB_xls_Name+']  '
  set @Msg = '�P�O Excel �C�L��� �O�_�s�b'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
    
  delete @RowData
  insert into @RowData exec (@strSQL)
  select @Print_Date=Rtrim(isnull(aData, '')) from @RowData
  set @Print_Date = Convert(Varchar(10), CONVERT(Date, @Print_Date) , 111)
    
  --2013/11/25 �L�ҽT�w�C��T�w�H24�������ɤ� 
  --set @Print_Date = Substring(Convert(Varchar(10), @Print_Date, 111), 1, 8)+@Last_Date
    
  if Rtrim(@Print_Date) = ''
  begin
print '4'
     set @Cnt = @Errcode
     set @strSQL = ''
     Set @Msg = '�䤣�� Excel ��Ƥ����C�L����A�פ�i�����ɧ@�~�C'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
     -- 2013/11/28 �W�[���Ѧ^�ǭ�
     Return(@Errcode)
  end
  else
  begin
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
                  ' where rowid <= 1 '
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

print '17'
    IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_OD_Name+']') AND type in (N'U'))
    begin
print '18'
       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
       Set @Msg = '����'+@TB_OD_Name+'��b�`�� ['+@TB_head+']�C'
    
       Set @strSQL = 'Create Table '+@TB_OD_Name+' ('+@CR+
                     '    [ct_no] [varchar](50) NULL, '+@CR+
                     '    [ct_sname] [varchar](50) NULL, '+@CR+

                     '    [sk_no] [varchar](50) NULL, '+@CR+
                     '    [sk_name] [varchar](50) NULL, '+@CR+
                     '    [sk_bcode] [varchar](50) NULL, '+@CR+

                     '    [xls_skno] [varchar](50) NULL, '+@CR+
                     '    [xls_cono] [varchar](50) NULL, '+@CR+
                     '    [xls_skname] [varchar](50) NULL, '+@CR+

                     '    [fg6_qty] [numeric](38, 6) NOT NULL, '+@CR+
                     '    [fg7_qty] [numeric](38, 6) NOT NULL, '+@CR+

                     '    [sum_qty] [numeric](38, 6) NOT NULL, '+@CR+
                     '    [xls_qty] [numeric](20, 6) NOT NULL, '+@CR+
                     '    [diff_qty] [numeric](38, 6) NOT NULL, '+@CR+

                     '    [fg6_amt] [numeric](38, 6) NOT NULL, '+@CR+
                     '    [fg7_amt] [numeric](38, 6) NOT NULL, '+@CR+
                     '    [sum_amt] [numeric](38, 6) NOT NULL, '+@CR+

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

print '19'
    /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
    Excel ��l�榡
    20151031_Comp_EC_Consign_Stock_Momo �榡:
    F01:�ܧO          F02:�ӫ~��t�s��(V) F03:�~��(V)     F04:�~�W(V)     F05:�W��
    F06:�ӫ~���A      F07:����w�s        F08:�i�f�q      F09:�Ȱh��      F10:�P�f��
    F11:�ݰh�q        F12:�����w�s        F13:�Y�ɮw�s(V) F14:�H�ܦb�~
    *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/

    Set @Msg = '�i����ӫ~�s���üg�J ['+@TB_OD_Name+']��ƪ�C'
    set @strSQL = -- CTE_Q1 ��l Excel ��ƶȨ��������
                 'With CTE_Q1 as ( '+@CR+
                 '  Select F2 as xls_skno, '+@CR+ -- F02:�ӫ~���X(Excel)
                 '         F4 as xls_skname, '+@CR+ -- F04: �ӫ~�W��(Excel)
                 '         F3 as xls_cono, '+@CR+ -- F03: ���U�s��(Excel)
                 '         Sum(isnull(Convert(Numeric(20, 6), Replace(F13, '','', '''')), 0)) as xls_qty '+@CR+ -- F13:�H��w�s (Excel)
                 '    from '+@TB_tmp_Name+@CR+
                 '   Group by F2, F4, F3 '+@CR+
                 '), CTE_Q2 as ( '+@CR+
                 '  Select distinct '+@CR+ 
                 '         RTrim(ct_no) as ct_no, '+@CR+ 
                 '         RTrim(ct_sname) as ct_sname, '+@CR+ 
                 '         RTrim(co_skno) as co_skno, '+@CR+  
                 '         RTrim(co_cono) as co_cono, '+@CR+  
                 '         RTrim(sk_no) as sk_no, '+@CR+  
                 '         RTrim(sk_name) as sk_name, '+@CR+  
                 '         RTrim(sk_bcode) as sk_bcode '+@CR+ 
                 '    from SYNC_TA13.dbo.sauxf m '+@CR+  
                 '         left join SYNC_TA13.dbo.sstock d '+@CR+  
                 '           on m.co_skno = d.sk_no '+@CR+  
                 '         left join SYNC_TA13.dbo.pcust d1 '+@CR+  
                 '           on m.co_ctno = d1.ct_no '+@CR+  
                 '   where 1=1 '+@CR+ 
                 '     and m.co_class=''1'' '+@CR+  
                 '     and m.co_ctno = '''+@CompanyLikeCode+''' '+@CR+ 
                 -- �u�d�U��^�f����
                 '     and exists '+@CR+ 
                 '         (select * '+@CR+ 
                 '            from SYNC_TA13.dbo.sslpdt dt '+@CR+ 
                 '           where 1=1 '+@CR+ 
                 '             and dt.sd_slip_fg IN (''6'',''7'') '+@CR+   
                 '             and d.sk_no = dt.sd_skno '+@CR+ 
                 '             and dt.sd_ctno = '''+@CompanyLikeCode+''' '+@CR+ 
                 '             and dt.sd_date <= '''+@Print_Date+'''  '+@CR+
                 '         ) '+@CR+
                 '), CTE_Q3 as ( '+@CR+
                 '  Select sd_ctno, ct_sname, sd_skno, '+@CR+
                 '         sum(case when sd_slip_fg = ''6'' then sd_qty else 0 end) as fg6_qty, '+@CR+ -- ����ƶq
                 '         sum(case when sd_slip_fg = ''7'' then sd_qty else 0 end) as fg7_qty, '+@CR+ -- �^�f�ƶq
                 '         sum(case when sd_slip_fg = ''7'' then sd_qty * -1 else sd_qty end) as sum_qty, '+@CR+ -- �X�p�ƶq<-------CHECK
                 '         sum(case when sd_slip_fg = ''6'' then sd_stot else 0 end) as fg6_amt, '+@CR+ -- ������B
                 '         sum(case when sd_slip_fg = ''7'' then sd_stot else 0 end) as fg7_amt, '+@CR+ -- ����ƶq
                 '         sum(case when sd_slip_fg = ''7'' then sd_stot * -1 else sd_stot end) as sum_amt '+@CR+ -- �X�p���B
                 '    from Fact_sslpdt '+@CR+
                 '   where 1=1 '+@CR+
                 '     and sd_slip_fg IN (''6'',''7'')  '+@CR+
                 '     and sd_date >= ''2012/12/01''  '+@CR+
                 -- �d��I���� 2013/01/01 ~ �C�L���
                 '     and sd_date <= '''+@Print_Date+'''  '+@CR+
                 '     and sd_ctno = '''+@CompanyLikeCode+''' '+@CR+ 
                 '   group by sd_ctno, ct_sname, sd_skno '+@CR+
                 ' ) '+@CR+

                 'Insert Into '+@TB_OD_Name+' '+@CR+
                 '  (ct_no, ct_sname, '+@CR+
                 '   sk_no, sk_name, sk_bcode, '+@CR+
                 '   xls_skno, xls_cono, xls_skname, '+@CR+
                 '   fg6_qty, fg7_qty, sum_qty, '+@CR+
                 '   fg6_amt, fg7_amt, sum_amt, '+@CR+
                 '   xls_qty, diff_qty, '+@CR+
                 '   Print_Date, isfound, Exec_DateTime '+@CR+
                 '  ) '+@CR+
   
                 'select Distinct '+@CR+
                 '       Convert(Varchar(50), RTrim(isnull(isnull(d.ct_no, d2.sd_ctno), ''N/A''))) as ct_no, '+@CR+
                 '       Convert(Varchar(50), RTrim(isnull(isnull(d.ct_sname, d2.ct_sname), ''N/A''))) as ct_sname, '+@CR+
   
                 -- �ӫ~�s��
                 '       Convert(Varchar(50), RTrim(isnull(isnull(d.co_skno, d1.sk_no), ''N/A''))) as sk_no, '+@CR+
                 -- �ӫ~�W��
                 '       Convert(Varchar(50), RTrim(isnull(isnull(d.sk_name, d1.sk_name), ''N/A''))) as sk_name, '+@CR+
                 -- �ӫ~���X
                 '       Convert(Varchar(50), RTrim(isnull(isnull(d.sk_bcode, d1.sk_bcode), ''N/A''))) as sk_bcode, '+@CR+
   
                 -- �ӫ~�s��(Excel)
                 '       Convert(Varchar(50), RTrim(Isnull(xls_skno, ''N/A''))) as xls_skno, '+@CR+ 
                 -- �ӫ~���U�s��(Excel)
                 '       Convert(Varchar(50), RTrim(Isnull(xls_cono, ''N/A''))) as xls_cono, '+@CR+ 
                 -- �ӫ~�W��(Excel)
                 '       Convert(Varchar(50), RTrim(isnull(xls_skname, ''N/A''))) as xls_skname, '+@CR+ 
   
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
                 '       isnull(Convert(Numeric(20, 6), xls_qty), 0) as xls_qty, '+@CR+
                 -- �t���ƶq
                 '       isnull(sum_qty - Convert(Numeric(20, 6), xls_qty), 0) as diff_qty, '+@CR+
                 '       '''+@Print_Date+''' as Print_Date, '+@CR+
                 -- ���`
                 '       case '+@CR+
                 '         when (isnull(d.co_skno, d1.sk_no) is null) then ''XLS �s���藍���V�ӫ~���'' '+@CR+
                 '         when (fg6_qty is null) then ''XLS ����ơA����V�L�U��^�f���'' '+@CR+
                 '         when (xls_skno is null and sum_qty <>0) then ''��V���U��^�f��ơA��XLS�L���'' '+@CR+
                 '         when (sum_qty - xls_qty <>0) then ''���ƶq���`'' '+@CR+
                 '         else '''' '+@CR+
                 '       end as isfound, '+@CR+
                 '       getdate() as Exec_DateTime '+@CR+
                 -- �Ȥ�s��
                 '  from CTE_Q1 m '+@CR+
                 -- �ӫ~���
                 '       Full join CTE_Q2 d '+@CR+
                 '         on 1=1 '+@CR+
                 '        and (ltrim(rtrim(m.xls_skno)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.co_skno)) collate Chinese_Taiwan_Stroke_CI_AS  '+@CR+
                 '         or  ltrim(rtrim(m.xls_skno)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.co_cono)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                 '         or  ltrim(rtrim(m.xls_skno)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.sk_no)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                 '         or  ltrim(rtrim(m.xls_skno)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.sk_bcode)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
   
                 '         or  ltrim(rtrim(m.xls_cono)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.co_skno)) collate Chinese_Taiwan_Stroke_CI_AS  '+@CR+
                 '         or  ltrim(rtrim(m.xls_cono)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.co_cono)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                 '         or  ltrim(rtrim(m.xls_cono)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.sk_no)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                 '         or  ltrim(rtrim(m.xls_cono)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.sk_bcode)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                 
                 '         or  ltrim(rtrim(d.sk_name)) collate Chinese_Taiwan_Stroke_CI_AS like ''%''+ltrim(rtrim(m.xls_skno))+''%'' collate Chinese_Taiwan_Stroke_CI_AS  '+@CR+
                 '         or  ltrim(rtrim(d.sk_name)) collate Chinese_Taiwan_Stroke_CI_AS like ''%''+ltrim(rtrim(m.xls_skno))+''%'' collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                 '         or  ltrim(rtrim(d.sk_name)) collate Chinese_Taiwan_Stroke_CI_AS like ''%''+ltrim(rtrim(m.xls_skno))+''%'' collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                 '         or  ltrim(rtrim(d.sk_name)) collate Chinese_Taiwan_Stroke_CI_AS like ''%''+ltrim(rtrim(m.xls_skno))+''%'' collate Chinese_Taiwan_Stroke_CI_AS) '+@CR+
   
                 '       left join SYNC_TA13.dbo.sstock d1 '+@CR+
                 '         on 1=1 '+@CR+
                 '        and (ltrim(rtrim(m.xls_skno)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d1.sk_no)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                 '         or  ltrim(rtrim(m.xls_skno)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d1.sk_bcode)) collate Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                 --  ��V����^�f��� sp_slip_fg = 6 ����, sp_slip_fg = 7 �^�f
                 '       left join CTE_Q3 d2 '+@CR+
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
    --2013/12/24�N���ƪ���Ƥ]�C�J���`
print '20'
/*
    set @Msg = '�N���ƪ���Ƽе����`'
    set @Cnt = 0
    set @strSQL ='update '+@TB_OD_Name+@CR+
                 '   set isfound = isfound+''�A���Ƥ��'' '+@CR+
                 ' where 1=1 '+@CR+
                 '   and sk_no in '+@CR+
                 '       (select sk_no '+@CR+
                 '          from '+@TB_OD_Name+@CR+
                 '         where print_date = '''+@Print_Date+''' '+@CR+
                 '         group by sk_no '+@CR+
                 '        having COUNT(*) > 1) '
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    */
print '23'
  end
  -- 2013/11/28 �W�[���Ѧ^�ǭ�
  Return(0)
end
GO
