USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_SCM_Mall_Normal_Order_800yaoya]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_Imp_SCM_Mall_Normal_Order_800yaoya]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_Imp_SCM_Mall_Normal_Order_800yaoya]
as
begin
  /***********************************************************************************************************
     2013/12/04 -- Rickliu
     �K�ʫ� ��x�Գ�@����q��A�q�����@�ߧ���s�����A�� ���w��f�� �h�g�J �q�檺��f���
     ���O�G
     1.�ѩ� 3C �P �ʳf ���Ȥ�s�����ۦP�A�ҥH�Գ�ɥ������}�B�z�C
     2.���{�ǶȳB�z�Ȥ�z�f�Ҳ��ͪ��z�f��A�z�L�����ন�ڥq���q��A�����H�~�h���b���B�z�d�򤺡C
  ***********************************************************************************************************/
  Declare @Proc Varchar(50) = 'uSP_Imp_SCM_Mall_Normal_Order_800yaoya'
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

  Declare @txt_head Varchar(100), @TB_head Varchar(200), @TB_txt_Name Varchar(200), @TB_OR_Name Varchar(200), @TB_tmp_name Varchar(200)

  Declare @strSQL Varchar(Max)
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @RowCount Table (cnt int)
  Declare @RowData Table (aData Varchar(255))
  
  Declare @New_Orno varchar(10)
  Declare @Last_Date Varchar(2)
  Declare @Sct_SubQuery Varchar(1000) -- �P�P�����l�d��

  Set @CompanyName = '�K�ʫ�' --> �ФŶ��ܰ�
  Set @CompanyLikeName = '%'+@CompanyName+'%'
  Set @Or_Maker = 'Admin'
  Set @Sct_SubQuery = '(select Top 1 @@ from SYNC_TA13.dbo.sctsale where 1=1 and ss_ctno LIKE ''%''+Rtrim(m.ct_no)+''%'' and ss_no Like ''%''+Rtrim(isnull(d.co_skno, d1.sk_no))+''%'' and (ss_edate = ''1900/01/01'' or ss_edate >= getdate()) order by ss_edate desc ,ss_sdate desc)'

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '1'
  Declare Cur_Mall_Normal_Order_800yaoya_DataKind cursor for
    select *
      from (select '2' as Kind, '3C' as KindName
             union
            select '1' as Kind, 'Retail' as KindName
           )m,
           (select 'Order' as OR_Name, '�z�f��q��' as OR_CName, '3' as OR_Class
           )d
    order by kind
     
  open Cur_Mall_Normal_Order_800yaoya_DataKind
  fetch next from Cur_Mall_Normal_Order_800yaoya_DataKind into @Kind, @KindName, @OR_Name, @OR_CName, @OR_Class

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '2'
  while @@fetch_status =0
  begin
print '3'
    Set @txt_head = 'Ori_Txt#'
    Set @TB_head = 'Mall_Normal_Order_800yaoya_'
    Set @TB_txt_Name = @txt_head+@TB_Head+@KindName -- Ex: Ori_txt#Mall_Normal_Order_800yaoya_3C -- Text ���ɸ��
    Set @TB_tmp_Name = @TB_Head+@KindName+'_'+'tmp' -- Ex: Mall_Normal_Order_800yaoya_3C_tmp -- �{����J����
    Set @TB_OR_Name = @TB_Head+@KindName -- Ex: Mall_Normal_Order_800yaoya_3C
    Set @Rm = '�t�ζפJ'+@CompanyName+@KindName

    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    -- Check �פJ�ɮ׬O�_�s�b
    IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_txt_Name+']') AND type in (N'U'))
    begin
print '4'
       --Send Message
       set @Cnt = @Errcode
       set @strSQL = ''
       Set @Msg = '�~�� Text �פJ��ƪ� ['+@TB_txt_Name+']���s�b�A�פ�i�����ɧ@�~�C'
       Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

       close Cur_Mall_Normal_Order_800yaoya_DataKind
       deallocate Cur_Mall_Normal_Order_800yaoya_DataKind
       Return(@Errcode)
    end
    
    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    -- �P�O Text �C�L��� �O�_�s�b
print '5'
    set @Cnt = 0
    set @strSQL = 'Select Distinct SUBSTRING(Convert(Text, F1), 11, 10) as Or_Date1 from [dbo].['+@TB_txt_Name+'] where F1 like ''%�s����%'' '
    set @Msg = '�P�O Text �s���� �O�_�s�b'
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

    --�ѩ�Գ������T�w�A�ҥH�@�ߧ�H��檺�s�������D�C
    --�ثe�Ȥ䴩�P�@�Ѫ��s�����u����@���A�קK�]�H�����y���A�������л\
    delete @RowData
    insert into @RowData exec (@strSQL)
    select @Or_Date1=Rtrim(isnull(aData, '')) from @RowData
    set @Or_Date1 = Convert(Varchar(10), CONVERT(Date, @Or_Date1) , 111)
    
    if @Or_Date1 = ''
    begin
print '6'
       set @Cnt = @Errcode
       set @strSQL = ''
       Set @Msg = '�䤣�� Text ��Ƥ����s�����A�פ�i�����ɧ@�~�C'
       Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

       close Cur_Mall_Normal_Order_800yaoya_DataKind
       deallocate Cur_Mall_Normal_Order_800yaoya_DataKind
       Return(@Errcode)
    end

    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    -- �W�[�P�O �T�{ �� ���T�{ ��� �Y�O�s�b�h���o�i�歫��@�~�C
    -- 2013/10/28 �W�[�P�_�Y���24��w�ѥ��T�{��T�{��h���A����A�קK��V���T�{��ڤ��_�R���s�W�y���渹�ּW���p
print '7'
    -- set @Or_Date1 = Substring(Convert(Varchar(10), @Or_Date1, 111), 1, 8)+@Last_Date
    select @Or_cnt = sum(cnt)
      from (select count(1) as cnt
              from SYNC_TA13.dbo.sorder m
             where or_class = @or_class
               and or_date1 = @Or_Date1
               and or_rem like '%'+@Rm+'%'
             union
            select count(1) as cnt
              from SYNC_TA13.dbo.sordertmp m
             where or_class = @Or_class
               and or_date1 = @Or_Date1
               and or_rem like '%'+@Rm+'%') m

    if @Or_cnt > 0
    begin
print '8'
       set @Cnt = @Errcode
       set @strSQL = ''
       set @Msg ='[ �w�T�{ �� ���T�{ �� ] �w�� ['+@Or_Date1+' '+@OR_CName+'] ��ơA�Y�ݭn����h����M���T�{�Υ��T�{��ګ᭫��Y�i�Csorder, @or_class=['+@OR_class+'], @SP_Date1=['+@Or_Date1+'], @or_lotno=['+@RM+']�C'
       Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
     
       close Cur_Mall_Normal_Order_800yaoya_DataKind
       deallocate Cur_Mall_Normal_Order_800yaoya_DataKind
       Return(@Errcode)
    end
    else
    begin
       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '9'
       IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_tmp_Name+']') AND type in (N'U'))
       begin
print '10'
          Set @Msg = '�M�� '+@CompanyName+' �{����J������ƪ� ['+@TB_tmp_Name+']�C'

          Set @strSQL = 'DROP TABLE [dbo].['+@TB_tmp_Name+']'
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       end

       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
       --�N�C����ƥ[�J�ߤ@���(ps.����ƪ�ФűƧǡA�O�d��l Text �˻��A�H�Q��Უ�X��b��)
       --�إ߼Ȧs�ɥB�إߧǸ��A�إߧǸ��B�J�ܭ��n�A�]������n�̧ǲ��ͩ��W
print '11'
       Set @Msg = '�s�W�ƥ���� F2 ~ F11 ['+@TB_tmp_Name+']�C'
       Set @strSQL = 'select rowid, '+@CR+ -- �C��
                     '       F1, '+@CR+ --����
                     '       Convert(Varchar(100), '''') as F2,  '+@CR+  --���W
                     '       Convert(Varchar(100), '''') as F3,  '+@CR+  --��      
                     '       Convert(Varchar(100), '''') as F4,  '+@CR+  --�ӫ~�s��
                     '       Convert(Varchar(100), '''') as F5,  '+@CR+  --�ӫ~�W��
                     '       Convert(Varchar(100), '''') as F6,  '+@CR+  --�q�f�q  
                     '       Convert(Varchar(100), '''') as F7,  '+@CR+  --�ث~�_  
                     '       Convert(Varchar(100), '''') as F8,  '+@CR+  --�z�f�q  
                     '       Convert(Varchar(100), '''') as F9,  '+@CR+  --�i��    
                     '       Convert(Varchar(100), '''') as F10,  '+@CR+ --�p�p    
                     '       Convert(Varchar(100), '''') as F11,  '+@CR+ --�Ƶ�    
                     '       Convert(Varchar(100), '''') as F12,  '+@CR+ --���ʳ渹    
                     '       Convert(Varchar(100), '''') as F13,  '+@CR+ --���w��f��    
                     '       print_date='''+@Or_Date1+''', '+@CR+
                     /*
                     '       xlsFileName, '+@CR+ -- ��l XLS �ɮצW��
                     '       imp_date, '+@CR+ -- ���� SP_Imp_xls_to_db �פJ���
                     */
                     '       SP_Exec_date=getdate() '+@CR+ -- ���楻�{�Ǥ��
                     '       into ['+@TB_tmp_Name+']'+@CR+
                     '  from ['+@TB_txt_Name+']'
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '12'
       Set @Msg = '������O��� ['+@TB_tmp_Name+']�C'
       Set @strSQL = 'update ['+@TB_tmp_Name+']'+@CR+
                     '   set F2=Replace(Replace(REPLACE(F1, '' '', ''''), ''��)�t�Ӵz�f��(���e����)'', ''''), ''('', '''') '+@CR+
                     ' where F1 like ''%'+SPACE(10)+'(%'''
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       
       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '13'
       -- ���B�|�N�C�@�������W���g�J F2 ���
       Set @Msg = '�B�z�{����J������ƪ��s�դp�p���W��� ['+@TB_tmp_Name+']�C'
       Set @strSQL = 'declare cur_book cursor for '+@CR+
                     '  select rowid, f1, f2 '+@CR+
                     '    from ['+@TB_tmp_Name+'] '+@CR+
                     '   order by rowid '+@CR+
                     ''+@CR+
                     'declare @rowid int, @cnt int '+@CR+
                     'declare @f1 varchar(max), @f2 varchar(100), @vf2 varchar(100) '+@CR+
                     'declare @f12 varchar(10), @vf12 varchar(10)'+@CR+
                     'declare @f13 varchar(10), @vf13 varchar(10)'+@CR+
                     'set @cnt=1 '+@CR+
                     ''+@CR+
                     'open cur_book '+@CR+
                     'fetch next from cur_book into @rowid, @f1, @f2 '+@CR+
                     ''+@CR+
                     'set @cnt=1 '+@CR+
                     'while @@fetch_status =0 '+@CR+
                     'begin '+@CR+
                     ''+@CR+
                     -- �ѩ�C�@�a���ۤv�����ʳ渹�P���w��f��A�ҥH�B�~�B�z
             --'  print @f1  '+@CR+
                     '  set @f12 = substring(Convert(Text, @f1), 52, 8) '+@CR+
             --'  print @f12  '+@CR+
                     '  if @f12 = ''���ʳ渹'' '+@CR+
                     '     set @vf12 = Substring(Convert(Text, @f1), 64, 10) '+@CR+
             --'  print @vf12  '+@CR+
                     ''+@CR+
                     '  set @f13 = substring(Convert(Text, @f1), 52, 10) '+@CR+
             --'  print @f13  '+@CR+
                     '  if @f13 = ''���w��f��'' '+@CR+
                     '     set @vf13 = Substring(Convert(Text, @f1), 64, 10) '+@CR+
             --'  print @vf13  '+@CR+
                     ''+@CR+
                     '  if @f2 <> '''' '+@CR+
                     '     set @vf2 = @f2 '+@CR+
                     '  else '+@CR+
                     '  begin '+@CR+
                     '     update ['+@TB_tmp_Name+'] '+@CR+
                     '        set f2 = isnull(@vf2, ''''), '+@CR+
                     '            f12 = isnull(@vf12, ''''), '+@CR+
                     '            f13 = isnull(@vf13, '''') '+@CR+
                     '      where rowid = @rowid '+@CR+
                     '  end '+@CR+
                     '  fetch next from cur_book into @rowid, @f1, @f2 '+@CR+
                     'end '+@CR+
                     'Close cur_book '+@CR+
                     'Deallocate cur_book '
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '14'
       Set @Msg = '�B�z�U������� ['+@TB_tmp_Name+']�C'
       Set @strSQL = 'update ['+@TB_tmp_Name+']'+@CR+
                     '   set  F3=Replace(LTrim(RTrim(Substring(Convert(Text, F1),   1,  3))), '','', ''''), '+@CR+ --��
                     '        F4=Replace(LTrim(RTrim(Substring(Convert(Text, F1),   5, 14))), '','', ''''), '+@CR+ --�ӫ~�s��
                     '        F5=Replace(LTrim(RTrim(Substring(Convert(Text, F1),  20, 40))), '','', ''''), '+@CR+ --�ӫ~�W��
                     '        F6=Replace(LTrim(RTrim(Substring(Convert(Text, F1),  61,  4))), '','', ''''), '+@CR+ --�q�f�q
                     '        F7=Replace(LTrim(RTrim(Substring(Convert(Text, F1),  66,  4))), '','', ''''), '+@CR+ --�ث~�_  
                     '        F8=Replace(LTrim(RTrim(Substring(Convert(Text, F1),  72,  6))), '','', ''''), '+@CR+ --�z�f�q
                     '        F9=Replace(LTrim(RTrim(Substring(Convert(Text, F1),  78, 10))), '','', ''''), '+@CR+ --�i��
                     '       F10=Replace(LTrim(RTrim(Substring(Convert(Text, F1),  89, 12))), '','', ''''), '+@CR+ --�p�p
                     '       F11=LTrim(RTrim(Substring(Convert(Text, F1), 103, 10)))  '  --�Ƶ�
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '15'
       Set @Msg = '�R���{����J������ƪ��D���n��� ['+@TB_tmp_Name+']�C'
       Set @strSQL = 'delete ['+@TB_tmp_Name+']'+@CR+
                     ' where Isnumeric(Ltrim(Rtrim(substring(F1, 1, 3)))) <> 1'+@CR+
                     '    or Rtrim(Isnull(F4, '''')) = '''' '

       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
--=-[���ͳ�ک��Ӹ�ơA�ƶq���t�ƪ��n���Ͱh�f��]=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '16'
       IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_OR_Name+']') AND type in (N'U'))
       begin
print '17'
          Set @Msg = '�R����ƪ� ['+@TB_OR_Name+']'
          Set @strSQL = 'DROP TABLE [dbo].['+@TB_OR_Name+']'
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       end

       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
       --2013/04/25 ���@���v����J�N�n���Υ[�`�A����n��
       --�b���i����ӫ~�s���A�|�����U�s���H�ΰӫ~�򥻸���ɶi����@�~
print '18'
       Set @Msg = '�i����ӫ~�s���üg�J ['+@TB_OR_Name+']��ƪ�C'
       set @strSQL = 'select rowid, ct_no, ct_sname, ct_ssname, '+@CR+
                     '       isnull(d.co_skno, d1.sk_no) as sk_no, '+@CR+
                     '       isnull(d.sk_name, d1.sk_name) as sk_name, '+@CR+
                     -- ��ӫ~�W��
                     '       f5, '+@CR+ 
                     '       isnull(d.sk_bcode, d1.sk_bcode) as sk_bcode, '+@CR+
                     -- ��ӫ~�s��
                     '       f4, '+@CR+ 
                          -- sum(case when f13 < 0 then f13 * -1 else f13 end) as sale_qty, 
                          -- sum(case when f18 < 0 then f18 * -1 else f18 end) as sale_amt
                     --'       case when Convert(Numeric(20, 6), f13) < 0 then Convert(Numeric(20, 6), f13) * -1 else Convert(Numeric(20, 6), f13) end as sale_qty, '+@CR+
                     -- ��ӫ~�ƶq
                     '       Convert(Numeric(20, 6), F6) as Ori_sale_qty, '+@CR+ -- 2015/08/19 Rickliu �O�d�쥻���A�קK�᭱�ק�
                     '       Convert(Numeric(20, 6), F6) as sale_qty, '+@CR+
                     -- ���
                     '       (Convert(Numeric(20, 6), F10)) / (Convert(Numeric(20, 6), F6)) as Ori_sale_amt, '+@CR+ -- 2015/08/19 Rickliu �O�d�쥻���A�קK�᭱�ק�
                     '       (Convert(Numeric(20, 6), F10)) / (Convert(Numeric(20, 6), F6)) as sale_amt, '+@CR+
                     -- �p�p
                     '       (Convert(Numeric(20, 6), F10)) as Ori_sale_tot, '+@CR+ -- 2015/08/19 Rickliu �O�d�쥻���A�קK�᭱�ק�
                     '       (Convert(Numeric(20, 6), F10)) as sale_tot, '+@CR+ 
                     -- �Ƶ�
                     '       F11,  '+@CR+ 
                     -- ���ʳ渹
                     '       F12,  '+@CR+ 
                     -- ���w��f��
                     '       F13,  '+@CR+ 
                     -- �O�_����
                     '       case when isnull(d.co_skno, d1.sk_no) is null then ''N'' else ''Y'' end as isfound, '+@CR+
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
                     '       1 as isSend '+@CR+
                     '       into ['+@TB_OR_Name+'] '+@CR+
                     '  from (select * '+@CR+
                     '          from (select ct_no, rtrim(ct_sname) as ct_sname, rtrim(ct_name)+''#''+rtrim(ct_sname) as ct_ssname '+@CR+
                     '                  from SYNC_TA13.dbo.PCUST  '+@CR+
                     '                 where ct_class =''1''  '+@CR+
                     '                   and ct_name like '''+@CompanyLikeName+''' '+@CR+
                     '                   and substring(ct_no, 9, 1) ='''+@Kind+''' '+@CR+ -- -- 1:�ʳf, 2:3C, 3:�H��, 4:�N�uOEM
                     '                ) m, ['+@TB_tmp_Name+'] d '+@CR+
                     '         where m.ct_ssname collate Chinese_Taiwan_Stroke_CI_AS like ''%''+rtrim(d.f2)+''%'' collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                     --'           and Convert(Numeric(20, 6), f6) '+@oper+' 0 '+@CR+ -- �P�_�ƶq�O�_���t�ƥN��h�f
                     '       )m '+@CR+
                     '       left join '+@CR+ 
                     '         (select distinct co_ctno, co_skno, co_cono, sk_no, sk_name, sk_bcode '+@CR+
                     '            from SYNC_TA13.dbo.sauxf m '+@CR+ -- (�i�P�s)�ȤỲ�U�s��
                     '                 left join SYNC_TA13.dbo.sstock d '+@CR+
                     '                   on m.co_skno = d.sk_no '+@CR+
                     '           where co_class=''1'' '+@CR+
                     '         ) d '+@CR+
                     '          on 1=1 '+@CR+
                     '         and m.ct_no=d.co_ctno '+@CR+
                     -- F4 ��ӫ~�s��
                     '         and (ltrim(rtrim(m.f4)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.co_skno)) collate Chinese_Taiwan_Stroke_CI_AS  '+@CR+
                     '          or  ltrim(rtrim(m.f4)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.co_cono)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                     '          or  ltrim(rtrim(m.f4)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.sk_no)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                     '          or  ltrim(rtrim(m.f4)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.sk_bcode)) collate Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                     '        left join SYNC_TA13.dbo.sstock d1 '+@CR+
                     '          on 1=1 '+@CR+
                     '         and (ltrim(rtrim(m.f4)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d1.sk_no)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                     '          or  ltrim(rtrim(m.f4)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d1.sk_bcode)) collate Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                     ' order by 1 '
                     --group by ct_no, ct_sname, ct_ssname, sk_no, sk_name, sk_bcode, f3, f6
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

       -- 2015/08/19 Rickliu ���䴩�ƶq�ŶZ�����A�] SS_Mult_Text ����Ƥ���p�k�G@@19.000000;1:104.76@@1000.000000;1:61.9@@0.000000;:@@
       --                    @@19.000000;1:104.76 ==>    0 ~   19 ���椽�� = 104.76
       --                    @@1000.000000;1:61.9 ==>   20 ~ 1000 ���椽�� = 61.9
       --                    @@0.000000;:@@       ==> 1001 ~      ���椽�� = ""
       -- 2015/08/19 Rickliu �W�[�䴩 ��@�ƶq�϶�����
print '19'
       Set @Msg = '�i����ӫ~�s���üg�J ['+@TB_OR_Name+']��ƪ�C'
       set @strSQL = 'update ['+@TB_OR_Name+'] '+@CR+
                     '   set sale_amt = '+@CR+
                     '         case '+@CR+
                     '           when (sct_ss_sqty + sct_ss_eqty = 0) Or (sale_qty between sct_ss_sqty and sct_ss_eqty) '+@CR+
                     '           then Convert(Numeric(10,4), Convert(Varchar(10), sct_ss_form)) '+@CR+
                     '         end, '+@CR+
                     '       sale_tot =  '+@CR+
                     '         case '+@CR+
                     '           when (sct_ss_sqty + sct_ss_eqty = 0) Or (sale_qty between sct_ss_sqty and sct_ss_eqty) '+@CR+
                     '           then Convert(Numeric(10, 4), Convert(Varchar(10), sct_ss_form)) * sale_qty '+@CR+
                     '         end '+@CR+
                     ' where 1=1 '+@CR+
                     '   and Rtrim(sct_ss_csno + sct_ss_csname) <> '''' '+@CR+
                     '   and Rtrim(sct_ss_sendno) = '''' '+@CR+
                     '   and Substring(sct_ss_form, 1, 1) <> ''0'' '
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
   
print '20'
       -- 2015/08/19 Rickliu �W�[�䴩 ��@�ث~
       Set @Msg = '�i����ӫ~�s���üg�J ['+@TB_OR_Name+']��ƪ�C'
       set @strSQL = 'insert into ['+@TB_OR_Name+'] '+@CR+
                     'select rowid, ct_no, ct_sname, ct_ssname,'+@CR+
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
                     '       0 as isSend '+@CR+
                     '  from ['+@TB_OR_Name+'] m '+@CR+
                     '       left join fact_sstock d '+@CR+
                     '         on m.sct_ss_sendno = d.sk_no '+@CR+
                     ' where 1=1 '+@CR+
                     '   and Rtrim(sct_ss_csno + sct_ss_csname) <> '''' '+@CR+
                     '   and sct_ss_noqty <> 0 '+@CR+
                     '   and sct_ss_sendqty <> 0 '+@CR+
                     ' order by sct_ss_rec '
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
--       fetch next from Cur_Mall_Normal_Order_800yaoya_DataKind into @Kind, @KindName, @OR_Name, @OR_CName, @OR_Class
--       Continue
--goto proc_exit

print '20.5'

       Set @Msg = '�i���s���ƪ�rowid ['+@TB_OR_Name+']��ƪ�C'
       set @strSQL = 'update ['+@TB_OR_Name+'] '+@CR+
                     '   set rowid = rowid * 10  '+@CR+
                     ' where sk_no=sct_ss_sendno '+@CR+
                     '   and sk_no > ''''        '+@CR+
	                 ' and sct_ss_no > ''''      ' 
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL            
       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
       --���ͥX�f�h�^��
print '21'
       set @Cnt = 0
       Set @Msg = '�ˬd['+@TB_OR_Name+']�����ɸ�ƬO�_�s�b�C'
       set @strSQL = 'select count(1) from ['+@TB_OR_Name+']'
       delete @RowCount
       insert into @RowCount Exec (@strSQL)

       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
       -- �]���ϥ� Memory Variable Table, �ҥH�ϥ� >0 �P�_
       select @Cnt=cnt from @RowCount
print @Cnt
print '22'
       if @Cnt >0
       begin
       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '23'
          set @Msg = @Msg + '...�s�b�A�N�i����J�{�ǡC'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', @Cnt
          --���o�̷s�X�f�h�^�渹(�����q�w�֭�H�Υ��֭㪺������)
          --�ФŦA�[+1�A�]���s�W�ɷ|�۰ʥ[�J
          --2013/11/24 �אּ @od_date (�C��T�w�ϥ� 25�鰵�����ɤ�)
          set @New_Orno = (select convert(varchar(10), isnull(max(or_no), replace(substring(@Or_Date1, 3, 8), '/', '') +'0000')) 
                             from (select distinct or_no
                                     from SYNC_TA13.dbo.sorder m
                                    where or_class = @Or_Class
                                      and or_date1 = @Or_Date1
                                    union
                                   select distinct or_no
                                     from SYNC_TA13.dbo.sordertmp m
                                    where or_class = @Or_Class
                                      and or_date1 = @Or_Date1
                                  ) m
                          )
          set @Msg = '���o�̷s�q��渹(�����q�w�֭�H�Υ��֭㪺������),new_orno=['+@New_Orno+'], or_class=['+@Or_Class+'], or_date=['+@Or_Date1+'].'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', @Cnt
                         
          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '24'
          set @Msg = '�R�� ['+@Or_Date1+'] ����J���q���ک��ӡC'
          --2014/04/22 ���`���ܪ�����T�{�q��A�L���A��f�֡C

          set @strSQL ='delete TYEIPDBS2.lytdbta13.dbo.sorddt '+@CR+
                       ' where Convert(Varchar(10), od_date1, 111) = '''+@Or_Date1+''' '+@CR+
                       '   and od_class = '''+@Or_Class+''' '+@CR+
                       '   and od_lotno like ''%'+@Rm+'%'' '
           
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '25'
          set @Msg = '�R�� ['+@Or_Date1+'] ����J���q���ڥD�ɡC'
          --2014/04/22 ���`���ܪ�����T�{�q��A�L���A��f�֡C
          set @strSQL ='delete TYEIPDBS2.lytdbta13.dbo.sorder '+@CR+
                       ' where Convert(Varchar(10), or_date1, 111) = '''+@Or_Date1+''' '+@CR+
                       '   and or_class = '''+@Or_Class+''' '+@CR+
                       '   and or_rem like ''%'+@Rm+'%'' '
 
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '26'
          set @Msg = '�s�W ['+@Or_Date1+'] �q��D�ɡC'
          --2014/04/22 ���`���ܪ�����T�{�q��A�L���A��f�֡C
          set @strSQL ='insert into TYEIPDBS2.lytdbta13.dbo.sorder '+@CR+
                       '(  [OR_CLASS], [OR_NO], [OR_DATE1], [OR_DATE2], [OR_CTNO] '+@CR+
                       ' , [OR_CTNAME], [OR_SALES], [OR_DPNO], [OR_MAKER], [OR_TOT] '+@CR+
                       ' , [OR_TAX], [OR_RATE_NM], [OR_RATE], [OR_WKNO], [OR_REM] '+@CR+
                       ' , [or_ispack]  '+@CR+
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
                       '       Substring(F12+'':''+Convert(Varchar(10), rowid), 1, 10) as or_wkno, '+@CR+ -- �Ȥ�q��渹
                       '       '''+@Rm+'�� ''+Convert(Varchar(20), Getdate(), 120)+char(13)+char(10)+'''+@OR_CName+''' as or_rem, '+@CR+ -- �Ƶ�
                       '       ''0'' as od_is_pack  '+@CR+ -- ��f��
                       '  from ['+@TB_OR_Name+'] m '+@CR+
                       '       left join SYNC_TA13.dbo.sstock d '+@CR+
                       '         on m.sk_no = d.sk_no '+@CR+
                       '       left join SYNC_TA13.dbo.pcust d1 '+@CR+
                       '         on m.ct_no = d1.ct_no '+@CR+
                       '        and ct_class =''1'' '+@CR+
                       '       left join  '+@CR+
                       '        (select ct_no,  '+@CR+
                       '                convert(varchar(10), '+@New_Orno+'+row_number() over(order by m.ct_no)) as od_no '+@CR+ -- �f��s��
                       '           from (select distinct m.ct_no '+@CR+
                       '                   from ['+@TB_OR_Name+'] m  '+@CR+
                       '                        left join SYNC_TA13.dbo.pcust d1  '+@CR+
                       '                          on m.ct_no = d1.ct_no  '+@CR+
                       '                         and ct_class =''1'' '+@CR+
                       '                 )m '+@CR+
                       '         ) d3 on d1.ct_no = d3.ct_no '+@CR+
                       ' where isfound =''Y'' '+@CR+
                       --' where od_class = '''+@Or_Class+''' '+@CR+
                       --'   and od_date1 = '''+@Or_Date1+''' '+@CR+
                       --'   and od_lotno like '''+@RM+'%'+@OR_CName+'%'' '+@CR+
                       ' group by --od_class, '+@CR+
                       '       --od_date1, '+@CR+
                       '       --od_date2, '+@CR+
                       '       d3.od_no, '+@CR+
                       '       d1.ct_no, '+@CR+ -- �Ȥ�s��
                       '       d1.ct_name, '+@CR+ -- �Ȥ�W��
                       '       substring(d1.ct_addr3, 1, 255), '+@CR+ -- �e�f�a�}
                       '       d1.ct_sales, '+@CR+ -- �~�ȭ�
                       '       d1.ct_dept, '+@CR+ -- �����s��
                       '       d1.ct_porter, '+@CR+ -- �f�B���q
                       '       Substring(F12+'':''+Convert(Varchar(10), rowid), 1, 10), '+@CR+ -- �Ȥ�q��渹
                       '       d1.ct_payfg '
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL


          set @Msg = '�s�W ['+@Or_Date1+'] ��q����ӡC'
          --2014/04/22 ���`���ܪ�����T�{�q��A�L���A��f�֡C
          set @strSQL ='insert into TYEIPDBS2.lytdbta13.dbo.sorddt '+@CR+
                       '(od_class, or_no, od_date1, od_date2, od_ctno, '+@CR+
                       ' od_skno, od_name, od_price, od_unit, od_qty, '+@CR+
                       ' od_unit_fg, od_dis, od_stot, od_rate, od_ave_p, '+@CR+
                       ' od_lotno, od_seqfld, od_spec, od_sendfg, '+@CR+
                       ' od_csno, od_csrec '+@CR+
                       ' ,od_is_pack '+@CR+-- �w�槹�_
                       ') '+@CR+
                       'select '''+@Or_Class+''' as od_class, '+@CR+ -- ���O
                       '       d2.OR_NO, '+@CR+ -- �f��s��
                       '       d2.OR_DATE1 as od_date1, '+@CR+ -- �f����
					   -- 20160205 modify by Nan ��ޱ���X�A�i�N��f����P�f�����ۦP�A�Y�����ťաA����N�L�k�ק��f���A�C
                       '       d2.OR_DATE2 as od_date2, '+@CR+ -- ��f���
                       --'       --convert(datetime, f13) as od_date2, '+@CR+ -- ��f/���Ĥ��
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
                       '       rowid as od_seqfld, '+@CR+ -- ���ӧǸ�
                       '       F12+'':''+Convert(Varchar(10), rowid) as od_spec, '+@CR+ -- ���ӧǸ�, ���Ǹ��|�]����V�ק�ӥ[�H�ܧ�A�ҥH�t�s�@���� od_spec
                       '       Case '+@CR+
                       '         when m.sale_amt = 0 '+@CR+ --2013/12/28 �ȪA�۴f��ܡA�Y���B���s�h�N���ث~
                       '         then 1 '+@CR+
                       '         else 0 '+@CR+
                       '       end as od_sendfg, '+@CR+ -- �O�_���ث~
                       '       sct_ss_csno as od_csno,  '+@CR+ -- �P�P��ץN��
                       '       sct_ss_rec as od_csrec,  '+@CR+ -- �P�P��ש��ӥN�X
                       '       ''0'' as od_is_pack  '+@CR+ -- ��f��
                       '  from ['+@TB_OR_Name+'] m '+@CR+
                       '       left join SYNC_TA13.dbo.sstock d '+@CR+
                       '         on m.sk_no = d.sk_no '+@CR+
                       '       left join SYNC_TA13.dbo.pcust d1 '+@CR+
                       '         on m.ct_no = d1.ct_no '+@CR+
                       '        and ct_class =''1'' '+@CR+					   
                       '       left join TYEIPDBS2.LYTDBTA13.dbo.sorder d2 '+@CR+ -- 2017/03/28 Rickliu �]���q�\��ƨS�k�o�򨳳t�Y�ɡA�]���u��^�Y�h�����V��Ʈw
                       '         on m.ct_no = d2.OR_CTNO '+@CR+	
                       --'       left join  '+@CR+
                       --'        (select ct_no,  '+@CR+
                       ---- 2015/11/19 Rickliu, �ѩ�P�@�ѦP�@�a�����i����J�h�i�q�污�p�A�]���W�[ F12 ��q�檺���ʳ渹 �@�����渹�X�ϧO
                       ----'                convert(varchar(10), '+@New_Orno+'+row_number() over(order by m.ct_no)) as od_no '+@CR+ -- �f��s��
                       --'                convert(varchar(10), '+@New_Orno+'+row_number() over(order by m.ct_no, Ori_Buyno)) as od_no, '+@CR+ -- �f��s��
                       --'                Ori_Buyno '+@CR+ -- ��l�ĳ�s��
                       --'           from (select distinct m.ct_no, F12 as Ori_Buyno'+@CR+
                       --'                   from ['+@TB_OR_Name+'] m  '+@CR+
                       --'                        left join TYEIPDBS2.lytdbta13.dbo.pcust d1  '+@CR+
                       --'                          on m.ct_no = d1.ct_no  '+@CR+
                       --'                         and ct_class =''1'' '+@CR+
                       --'                 )m '+@CR+
                       --'         ) d2 on d1.ct_no = d2.ct_no '+@CR+
                       --'             And m.F12 = d2.Ori_Buyno'+@CR+
                       ' where isfound =''Y'' '+@CR+
					   '   and d2.or_class = '''+@Or_Class+''' ' +@CR+
					   '   and m.F12 COLLATE Chinese_Taiwan_Stroke_CI_AS = d2.or_wkno' +@CR+
					   '   and d2.or_rem like ''%'+@Rm+'%'' ' +@CR+
                       ' Order by 1, 2, 3, m.ct_no'
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL





        --  set @Msg = '�s�W ['+@Or_Date1+'] ��q����ӡC'
        --  --2014/04/22 ���`���ܪ�����T�{�q��A�L���A��f�֡C
        --  set @strSQL ='insert into TYEIPDBS2.lytdbta13.dbo.sorddt '+@CR+
        --               '(od_class, or_no, od_date1, od_date2, od_ctno, '+@CR+
        --               ' od_skno, od_name, od_price, od_unit, od_qty, '+@CR+
        --               ' od_unit_fg, od_dis, od_stot, od_rate, od_ave_p, '+@CR+
        --               ' od_lotno, od_seqfld, od_spec, od_sendfg, '+@CR+
        --               ' od_csno, od_csrec '+@CR+
        --               ') '+@CR+
        --               'select '''+@Or_Class+''' as od_class, '+@CR+ -- ���O
        --               '       d2.od_no, '+@CR+ -- �f��s��
        --               '       convert(datetime, '''+@Or_Date1+''') as od_date1, '+@CR+ -- �f����
					   ---- 20160205 modify by Nan ��ޱ���X�A�i�N��f����P�f�����ۦP�A�Y�����ťաA����N�L�k�ק��f���A�C
        --               '       convert(datetime, '''+@Or_Date1+''') as od_date2, '+@CR+ -- ��f���
        --               --'       --convert(datetime, f13) as od_date2, '+@CR+ -- ��f/���Ĥ��
        --               '       m.ct_no as od_ctno, '+@CR+ -- �Ȥ�s��

        --               '       d.sk_no as od_skno, '+@CR+ -- �f�~�s��
        --               '       d.sk_name as od_name, '+@CR+ -- �~�W�W��
        --               '       m.sale_amt as od_price, '+@CR+ -- ����
        --               '       d.sk_unit as od_unit, '+@CR+ -- ���
        --               '       m.sale_qty as od_qty, '+@CR+ -- ����ƶq
                       
        --               '       0 as od_unit_fg, '+@CR+ -- ���X��
        --               '       1 as od_dis, '+@CR+ -- ���
        --               '       m.sale_tot as od_stot, '+@CR+ -- �p�p
        --               '       1 as od_rate, '+@CR+ -- �ײv
        --               '       d.s_price4 as od_ave_p, '+@CR+ -- ��즨��

        --               '       '''+@Rm+@OR_CName +''' as od_lotno, '+@CR+ -- �ƥ����
        --               '       rowid as od_seqfld, '+@CR+ -- ���ӧǸ�
        --               '       F12+'':''+Convert(Varchar(10), rowid) as od_spec, '+@CR+ -- ���ӧǸ�, ���Ǹ��|�]����V�ק�ӥ[�H�ܧ�A�ҥH�t�s�@���� od_spec
        --               '       Case '+@CR+
        --               '         when m.sale_amt = 0 '+@CR+ --2013/12/28 �ȪA�۴f��ܡA�Y���B���s�h�N���ث~
        --               '         then 1 '+@CR+
        --               '         else 0 '+@CR+
        --               '       end as od_sendfg, '+@CR+ -- �O�_���ث~
        --               '       sct_ss_csno as od_csno,  '+@CR+ -- �P�P��ץN��
        --               '       sct_ss_rec as od_csrec  '+@CR+ -- �P�P��ש��ӥN�X
        --               '  from ['+@TB_OR_Name+'] m '+@CR+
        --               '       left join TYEIPDBS2.lytdbta13.dbo.sstock d '+@CR+
        --               '         on m.sk_no = d.sk_no '+@CR+
        --               '       left join TYEIPDBS2.lytdbta13.dbo.pcust d1 '+@CR+
        --               '         on m.ct_no = d1.ct_no '+@CR+
        --               '        and ct_class =''1'' '+@CR+
        --               '       left join  '+@CR+
        --               '        (select ct_no,  '+@CR+
        --               -- 2015/11/19 Rickliu, �ѩ�P�@�ѦP�@�a�����i����J�h�i�q�污�p�A�]���W�[ F12 ��q�檺���ʳ渹 �@�����渹�X�ϧO
        --               --'                convert(varchar(10), '+@New_Orno+'+row_number() over(order by m.ct_no)) as od_no '+@CR+ -- �f��s��
        --               '                convert(varchar(10), '+@New_Orno+'+row_number() over(order by m.ct_no, Ori_Buyno)) as od_no, '+@CR+ -- �f��s��
        --               '                Ori_Buyno '+@CR+ -- ��l�ĳ�s��
        --               '           from (select distinct m.ct_no, F12 as Ori_Buyno'+@CR+
        --               '                   from ['+@TB_OR_Name+'] m  '+@CR+
        --               '                        left join TYEIPDBS2.lytdbta13.dbo.pcust d1  '+@CR+
        --               '                          on m.ct_no = d1.ct_no  '+@CR+
        --               '                         and ct_class =''1'' '+@CR+
        --               '                 )m '+@CR+
        --               '         ) d2 on d1.ct_no = d2.ct_no '+@CR+
        --               '             And m.F12 = d2.Ori_Buyno'+@CR+
        --               ' where isfound =''Y'' '+@CR+
        --               ' Order by 1, 2, 3, m.ct_no'
        --  Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
          
          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '27'
          set @Msg = '�^�g�ӫ~��ﭫ�ƳƵ��C'
          --2014/04/22 Rickliu ���`���ܪ�����T�{�q��A�L���A��f�֡C
          --2017/06/21 Rickliu ���������H Join Update ���覡�^�g TYEIPDBS2�A�|�j�q���� Lock�A����K���ͥH�U���~�T���A�]����H Cursor �覡 Updated �^�h
          -- ���~�T��:�L�k�q�s�����A�� "TYEIPDBS2" �� OLE DB ���Ѫ� "SQLNCLI10" ���o��ƦC����ơC
          --set @strSQL ='update TYEIPDBS2.lytdbta13.dbo.sorddt '+@CR+
          --             '   set od_rem='''+@Rm+@OR_CName+'�ӫ~��ﭫ�� RowID:''+D.Rowid '+@CR+
          --             '  from TYEIPDBS2.lytdbta13.dbo.sorddt M '+@CR+
          --             '       inner join '+@CR+
          --             '         (select Convert(Varchar(100), rowid) as rowid, count(1) as cnt '+@CR+
          --             '            from ['+@TB_OR_Name+'] '+@CR+
          --             '           group by rowid '+@CR+
          --             '          having count(*) >1 '+@CR+
          --             '         ) D '+@CR+
          --             '          on od_seqfld=D.rowid '+@CR+
          --             ' where Convert(Varchar(10), m.od_date1, 111) = '''+@Or_Date1+''' '+@CR+
          --             '   and m.od_class = '''+@Or_Class+''' '+@CR+
          --             '   and m.od_lotno like ''%'+@Rm+'%'' '

          set @strSQL ='Declare @rowid as varchar(100)= '''' '+@CR+
                       'Declare @cnt as int = 0 '+@CR+
                       ''+@CR+
                       'Declare Cur_'+@TB_OR_Name+'_Update_RowID Cursor for '+@CR+
                       '  select Convert(Varchar(100), rowid) as rowid, count(1) as cnt '+@CR+
                       '    from ['+@TB_OR_Name+'] '+@CR+
                       '   group by rowid '+@CR+
                       '  having count(*) >1 '+@CR+
                       ''+@CR+
                       'Open Cur_'+@TB_OR_Name+'_Update_RowID '+@CR+
                       'Fetch Next From Cur_'+@TB_OR_Name+'_Update_RowID into @rowid, @cnt '+@CR+
                       ''+@CR+
                       'While @@Fetch_status = 0 '+@CR+
                       'begin '+@CR+
                       '  update TYEIPDBS2.lytdbta13.dbo.sorddt '+@CR+
                       '     set od_rem='''+@Rm+@OR_CName+'�ӫ~��ﭫ�� RowID:''+@Rowid '+@CR+
                       '   where Convert(Varchar(10), od_date1, 111) = '''+@Or_Date1+''' '+@CR+
                       '     and od_class = '''+@Or_Class+''' '+@CR+
                       '     and od_lotno like ''%'+@Rm+'%'+@OR_CName+'%'' '+@CR+
                       '     and od_seqfld=@rowid '+@CR+
                       ''+@CR+
                       '  Fetch Next From Cur_'+@TB_OR_Name+'_Update_RowID into @rowid, @cnt '+@CR+
                       'end '+@CR+
                       ''+@CR+
                       'Close Cur_'+@TB_OR_Name+'_Update_RowID '+@CR+
                       'Deallocate Cur_'+@TB_OR_Name+'_Update_RowID '

          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
                       
          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
          -- �s�W�X�f�h�^�D�� PKey: sp_no (ASC), sp_slip_fg (ASC)
print '28'
          --set @Msg = '�s�W ['+@Or_Date1+'] �q��D�ɡC'
          ----2014/04/22 ���`���ܪ�����T�{�q��A�L���A��f�֡C
          --set @strSQL ='insert into TYEIPDBS2.lytdbta13.dbo.sorder '+@CR+
          --             '(  [OR_CLASS], [OR_NO], [OR_DATE1], [OR_DATE2], [OR_CTNO] '+@CR+
          --             ' , [OR_CTNAME], [OR_SALES], [OR_DPNO], [OR_MAKER], [OR_TOT] '+@CR+
          --             ' , [OR_TAX], [OR_RATE_NM], [OR_RATE], [OR_WKNO], [OR_REM] '+@CR+
          --             ' , [or_ispack]  '+@CR+
          --             ') '+@CR+
          --             'select distinct od_class as or_class, '+@CR+
          --             '       or_no, '+@CR+
          --             '       od_date1 as or_date1, '+@CR+
          --             '       od_date2 as or_date2, '+@CR+
          --             '       d1.ct_no as or_ctno, '+@CR+ -- �Ȥ�s��
                       
          --             '       d1.ct_name as or_ctname, '+@CR+ -- �Ȥ�W��
          --             '       d1.ct_sales as or_sales, '+@CR+ -- �~�ȭ�
          --             '       d1.ct_dept as or_dpno, '+@CR+ -- �����s��
          --             '       '''+@Or_Maker+''' as or_maker, '+@CR+ -- �s��H��
          --             '       sum(od_stot) as or_tot, '+@CR+ -- �p�p
                       
          --             '       sum(od_stot)*0.05 as or_tax, '+@CR+ -- ��~�|(��)
          --             '       ''NT'' as or_rate_nm, '+@CR+ -- �ײv�W��
          --             '       1 as or_rate, '+@CR+ -- �ײv
          --             '       Substring(od_spec, 1, 10) as or_wkno, '+@CR+ -- �Ȥ�q��渹
          --             '       '''+@Rm+'�� ''+Convert(Varchar(20), Getdate(), 120)+char(13)+char(10)+'''+@OR_CName+''' as or_rem, '+@CR+ -- �Ƶ�
          --             '       ''0'' '+@CR+ -- �w�槹�_
          --             '  from TYEIPDBS2.lytdbta13.dbo.sorddt m '+@CR+
          --             '       left join TYEIPDBS2.lytdbta13.dbo.sstock d '+@CR+
          --             '         on m.od_skno = d.sk_no '+@CR+
          --             '       left join TYEIPDBS2.lytdbta13.dbo.pcust d1 '+@CR+
          --             '         on m.od_ctno = d1.ct_no '+@CR+
          --             '        and ct_class =''1'' '+@CR+
          --             ' where od_class = '''+@Or_Class+''' '+@CR+
          --             '   and od_date1 = '''+@Or_Date1+''' '+@CR+
          --             '   and od_lotno like '''+@RM+'%'+@OR_CName+'%'' '+@CR+
          --             ' group by od_class, '+@CR+
          --             '       od_date1, '+@CR+
          --             '       od_date2, '+@CR+
          --             '       or_no, '+@CR+
          --             '       d1.ct_no, '+@CR+ -- �Ȥ�s��
          --             '       d1.ct_name, '+@CR+ -- �Ȥ�W��
          --             '       substring(d1.ct_addr3, 1, 255), '+@CR+ -- �e�f�a�}
          --             '       d1.ct_sales, '+@CR+ -- �~�ȭ�
          --             '       d1.ct_dept, '+@CR+ -- �����s��
          --             '       d1.ct_porter, '+@CR+ -- �f�B���q
          --             '       Substring(od_spec, 1, 10), '+@CR+ -- �Ȥ�q��渹
          --             '       d1.ct_payfg '
          --Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
          -- 2015/11/25 Rickliu �W�[�R���L�Y�����Ӹ�ơA�קK�ȪA�H���H�u���q����P�f�ɡA���ͦ]���ɥ��Ѫ���ƥX�{�C
          -- 2017/07/17 Rickliu �۱q��ƦP�B�אּ�q�\�覡��A������G�A��DB�i�� Update �|�`�A�ʥ��ѨåB���ͥH�U���~�T���A��L�䤣����D�Ҧb�A�u�n��g�� Cursor �覡
          -- ���~�T��:�L�k�q�s�����A�� "TYEIPDBS2" �� OLE DB ���Ѫ� "SQLNCLI10" ���o��ƦC����ơC
                       --'delete TYEIPDBS2.lytdbta13.dbo.sorddt '+@CR+
                       --'  from TYEIPDBS2.lytdbta13.dbo.sorddt m '+@CR+
                       --' where not exists '+@CR+
                       --'       (select * '+@CR+
                       --'          from TYEIPDBS2.lytdbta13.dbo.sorder d '+@CR+
                       --'         where m.od_class = d.or_class '+@CR+
                       --'           and m.or_no = d.or_no '+@CR+
                       --'           and m.od_date1 = d.or_date1 '+@CR+
                       --'       ) '+@CR+
                       --'   and m.od_class = '''+@Or_Class+''' '+@CR+
                       --'   and Convert(Varchar(10), m.od_date1, 111) = '''+@Or_Date1+''' '+@CR+
                       --'   and m.od_lotno like '''+@RM+'%'+@OR_CName+'%'' '

          set @Msg = '�R�� ['+@Or_Date1+'] �L�Y���q������ɡC'
          set @strSQL ='Declare @od_class as varchar(100) '+@CR+
                       'Declare @or_no as varchar(100) '+@CR+
                       'Declare @od_date1 as DateTime '+@CR+
                       'Declare @cnt as int = 0 '+@CR+
                       ''+@CR+
                       'Declare Cur_'+@TB_OR_Name+'_Del_No_Header Cursor for '+@CR+
                       '  select m.od_class, m.or_no, m.od_date1, count(distinct d.or_no) as cnt '+@CR+
                       '    from TYEIPDBS2.lytdbta13.dbo.sorddt m '+@CR+
                       '         left join TYEIPDBS2.lytdbta13.dbo.sorder d '+@CR+
                       '           on m.od_class = d.or_class '+@CR+
                       '          and m.or_no = d.or_no '+@CR+
                       '          and m.od_date1 = d.or_date1 '+@CR+
                       '   where 1=1 '+@CR+
                       '     and Convert(Varchar(10), m.od_date1, 111) = '''+@Or_Date1+''' '+@CR+
                       '     and m.od_lotno like ''%'+@Rm+'%'+@OR_CName+'%'' '+@CR+
                       '   group by m.od_class, m.or_no, m.od_date1 '+@CR+
                       '  having count(distinct d.or_no) > 1 '+@CR+
                       ''+@CR+
                       'Open Cur_'+@TB_OR_Name+'_Del_No_Header '+@CR+
                       'Fetch Next From Cur_'+@TB_OR_Name+'_Del_No_Header into @od_class, @or_no, @od_date1, @cnt '+@CR+
                       ''+@CR+
                       'While @@Fetch_status = 0 '+@CR+
                       'begin '+@CR+
                       '  Delete TYEIPDBS2.lytdbta13.dbo.sorddt '+@CR+
                       '   where 1=1 '+@CR+
                       '     and od_date1 = Convert(Varchar(10), @od_date1, 111) '+@CR+
                       '     and od_class = ''@od_Class'' '+@CR+
                       '     and or_no = ''@or_no'' '+@CR+
                       ''+@CR+
                       '  Fetch Next From Cur_'+@TB_OR_Name+'_Del_No_Header into @od_class, @or_no, @od_date1, @cnt '+@CR+
                       'end '+@CR+
                       ''+@CR+
                       'Close Cur_'+@TB_OR_Name+'_Del_No_Header '+@CR+
                       'Deallocate Cur_'+@TB_OR_Name+'_Del_No_Header '

          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       
       end
       else
       begin
       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '29'
          set @Msg = @Msg + '...���s�b�A�פ���J�{�ǡC'
          Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
       end

print '30'
    end
    fetch next from Cur_Mall_Normal_Order_800yaoya_DataKind into @Kind, @KindName, @OR_Name, @OR_CName, @OR_Class

print '31'
  end

proc_exit:
print '32'
  close Cur_Mall_Normal_Order_800yaoya_DataKind
  deallocate Cur_Mall_Normal_Order_800yaoya_DataKind
  Return(0)
print '33'
end
GO
