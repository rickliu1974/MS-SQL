USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_SCM_EC_Consign_Order_PChome]    Script Date: 07/24/2017 14:43:59 ******/
DROP PROCEDURE [dbo].[uSP_Imp_SCM_EC_Consign_Order_PChome]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_Imp_SCM_EC_Consign_Order_PChome]
as
begin
  /***********************************************************************************************************
     PCHome ��x�Գ�ɡA�п�ܨ̤��������ơA�p�� XLS �榡�~�|���T
     -- 2013/11/28 �W�[���Ѧ^�ǭ�
  ***********************************************************************************************************/
  Declare @Proc Varchar(50) = 'uSP_Imp_SCM_EC_Consign_Order_PChome'
  Declare @Cnt Int =0
  Declare @RowCnt Int =0
  Declare @Msg Varchar(4000) =''
  Declare @Errcode int = -1
  
  Declare @Kind Varchar(1), @KindName Varchar(10), @sd_class Varchar(1), @sd_slip_fg Varchar(1), @OD_Name Varchar(20), @OD_CName Varchar(20), @oper Varchar(1), @F1_Tag Varchar(100)
  Declare @xls_head Varchar(100), @TB_head Varchar(200), @TB_xls_Name Varchar(200), @TB_OD_Name Varchar(200), @TB_tmp_name Varchar(200)

  Declare @strSQL Varchar(Max)
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @RowCount Table (cnt int)
  Declare @RowData Table (aData Varchar(255))
  
  Declare @new_sdno varchar(10)
  Declare @sd_date varchar(10)
  Declare @sp_date varchar(10)
  Declare @Print_Date varchar(10)
  Declare @Last_Date Varchar(2)
  Declare @sp_cnt int
       
  --���ͳf�����A2013/8/27 �L�ҽT�w�C��T�w�H25�������ɤ�
  set @Last_Date = '25'
  set @sp_date = Substring(Convert(Varchar(10), getdate(), 111), 1, 8)+@Last_Date
  set @sd_date = @sp_date

  Declare @CompanyName Varchar(100), @CompanyLikeName Varchar(100), @Rm Varchar(100), @sp_maker Varchar(20), @str Varchar(200)
  Declare @Pos Int
  
  Set @CompanyName = '�����a�x���' --> �ФŶ��ܰ�
  Set @CompanyLikeName = '%'+@CompanyName+'%'
  Set @Rm = '�t�ζפJ'+@CompanyName
  Set @sp_maker = 'Admin'

print '1'
  Declare Cur_Mall_Consign_Order_PCHome_DataKind cursor for
    select *
      from (select 'Reject' as OD_Name, '�h�^��U��' as OD_CName, '1' as sd_class, '3' as sd_slip_fg, '<' as oper, '�h�f�f�ک���(�ӫ~�J�`) - �H�ܭq��' as F1_Tag
             union
            select 'Return' as OD_Name, '�U�^��P��' as OD_CName, '3' as sd_class, '7' as sd_slip_fg, '>' as oper, '�q��f�ک���(�ӫ~�J�`) - �H�ܭq��' as F1_Tag
           )d
     
  open Cur_Mall_Consign_Order_PCHome_DataKind
  fetch next from Cur_Mall_Consign_Order_PCHome_DataKind into @OD_Name, @OD_CName, @sd_class, @sd_slip_fg, @oper, @F1_Tag

print '2'
  while @@fetch_status =0
  begin
print '3'
    Set @xls_head = 'Ori_Xls#'
    Set @TB_head = 'EC_Consign_Order_PChome'
    Set @TB_xls_Name = @xls_head+@TB_Head -- Ex: Ori_Xls#EC_Consign_Order_PChome -- Excel ���ɸ��
    Set @TB_tmp_Name = @TB_Head+'_'+'tmp' -- Ex: EC_Consign_Order_PChome_tmp -- �{����J����
    Set @TB_OD_Name = @TB_Head+'_'+@OD_Name -- Ex: EC_Consign_Order_PChome_Reject  

    set @strSQL = ''
    set @Cnt = 0
    set @Msg = '�i�� '+@F1_Tag+' ==> '+@OD_CName+ ' ���ɳB�z�C'
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

    -- Check �פJ�ɮ׬O�_�s�b
    IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_xls_Name+']') AND type in (N'U'))
    begin
print '4'
       --Send Message
       set @Cnt = @Errcode
       set @strSQL = ''
       Set @Msg = '�~�� Excel �פJ��ƪ� ['+@TB_xls_Name+']���s�b�A�פ�i�����ɧ@�~�C'
       Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

       fetch next from Cur_Mall_Consign_Order_PCHome_DataKind into @OD_Name, @OD_CName, @sd_class, @sd_slip_fg, @oper, @F1_Tag
       Continue
    end
    
    -- �P�O Excel �C�L��� �O�_�s�b
print '5'
    set @Cnt = 0
    set @Print_Date  = ''

    set @strSQL = 'Select Distinct SUBSTRING(F2, CHARINDEX(''~'', F2)+1, 10) as Print_date from [dbo].['+@TB_xls_Name+'] where F2 like ''�C�b���%'' '
    set @Msg = '�P�O Excel �C�b��� �O�_�s�b'
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

    delete @RowData
    insert into @RowData exec (@strSQL)
    select @Print_Date=Rtrim(isnull(aData, '')) from @RowData
    set @Print_Date = Convert(Varchar(10), CONVERT(Date, @Print_Date) , 111)
    
    if @Print_Date = ''
    begin
print '6'
       set @Cnt = @Errcode
       set @strSQL = ''
       Set @Msg = '�䤣�� Excel ��Ƥ����C�L����A�פ�i�����ɧ@�~�C'
       Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
       -- 2013/11/28 �W�[���Ѧ^�ǭ�
       close Cur_Mall_Consign_Order_PCHome_DataKind
       deallocate Cur_Mall_Consign_Order_PCHome_DataKind
       Return(@Errcode)

--       fetch next from Cur_Mall_Consign_Order_PCHome_DataKind into @OD_Name, @OD_CName, @sd_class, @sd_slip_fg, @oper, @F1_Tag
--       Continue
    end
   
    -- �W�[�P�O �T�{ �� ���T�{ ��� �Y�O�s�b�h���o�i�歫��@�~�C
    -- 2013/10/28 �W�[�P�_�Y���24��w�ѥ��T�{��T�{��h���A����A�קK��V���T�{��ڤ��_�R���s�W�y���渹�ּW���p
print '7'
    set @Print_Date = Substring(Convert(Varchar(10), @Print_Date, 111), 1, 8)+@Last_Date
    Set @sp_date = @Print_Date
    Set @sd_date = @Print_Date

print '���:'+@Print_Date
    select @sp_cnt = sum(cnt)
      from (select count(1) as cnt
              from SYNC_TA13.dbo.sslpdt m
             where sd_class = @sd_class
               and sd_slip_fg = @sd_slip_fg
               and sd_date = @Print_Date
               and sd_lotno like '%'+@Rm+'%'
             union
            select count(1) as cnt
              from SYNC_TA13.dbo.sslpdttmp m
             where sd_class = @sd_class
               and sd_slip_fg = @sd_slip_fg
               and sd_date = @Print_Date
               and sd_lotno like '%'+@Rm+'%') m

    if @sp_cnt > 0
    begin
print '8'
       set @Cnt = @Errcode
       set @strSQL = ''
       set @Msg ='[ �w�T�{ �� ���T�{ �� ] �w�� ['+@Print_Date+' '+@OD_CName+'] ��ơA�Y�ݭn����h����M���T�{�Υ��T�{��ګ᭫��Y�i�Csslpdt, @sd_class=['+@sd_class+'], @sd_slip_fg=['+@sd_slip_fg+'], @sd_date=['+@Print_Date+'], sd_lotno=['+@RM+']�C'
       Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

       fetch next from Cur_Mall_Consign_Order_PCHome_DataKind into @OD_Name, @OD_CName, @sd_class, @sd_slip_fg, @oper, @F1_Tag
       Continue
    end
    else
    begin
print '9'
       -- 
       IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_tmp_Name+']') AND type in (N'U'))
       begin
print '10'
          Set @Msg = '�M�� '+@CompanyName+' �{����J������ƪ� ['+@TB_tmp_Name+']�C'

          Set @strSQL = 'DROP TABLE [dbo].['+@TB_tmp_Name+']'
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       end

       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
       --�N�C����ƥ[�J�ߤ@���(ps.����ƪ�ФűƧǡA�O�d��l XLS �˻��A�H�Q��Უ�X��b��)
       --�إ߼Ȧs�ɥB�إߧǸ��A�إߧǸ��B�J�ܭ��n�A�]������n�̧ǲ��ͩ��W
print '11'
       Set @Msg = '�إ��{����J������ƪ�C���ߤ@�Ǹ� ['+@TB_tmp_Name+']�C'
       Set @strSQL = 'select *, print_date='''+@Print_Date+''', '+@CR+
                     '       SP_Exec_date=convert(date, getdate()) '+@CR+ -- ���楻�{�Ǥ��
                     '  into '+@TB_tmp_Name+
                     '  from ['+@TB_xls_Name+'] '
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '12-1'
       Set @Msg = '�R���{����J������ƪ��D���n��� ['+@TB_tmp_Name+']�C'
       Set @strSQL = 'update ['+@TB_tmp_Name+'] '+
                     '   set F1=F2 '+
                     ' where F2 like ''%��b����%'' '
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       
print '12'
       Set @Msg = '�R���{����J������ƪ��D���n��� ['+@TB_tmp_Name+']�C'
       Set @strSQL = 'delete ['+@TB_tmp_Name+'] '+
                     ' where (Rtrim(F2) like ''%��b���Ӫ�%'') '+
                     '    or (Rtrim(F2) like ''%��b��%'') '+
                     '    or (Rtrim(F2) like ''%�`�B�G%'') '+
                     '    or (Rtrim(F2) like ''�ӫ~�W��'') '+
                     '    or (Rtrim(F2) like ''END'') '
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
print '13'
       -- ���B�|�N�C�@�������W���g�J
       Set @Msg = '�B�z�{����J������ƪ��s�դp�p���W��� ['+@TB_tmp_Name+']�C'
       Set @strSQL = 'declare cur_book cursor for '+@CR+
                     ' select rowid, f1 '+@CR+
                     '   from ['+@TB_tmp_Name+'] '+@CR+
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
                     '     update ['+@TB_tmp_Name+'] '+@CR+
                     '        set f1 = @vf1 '+@CR+
                     '      where rowid = @rowid '+@CR+
                     '  end '+@CR+
                     '  fetch next from cur_book into @rowid, @f1 '+@CR+
                     'end '+@CR+
                     'Close cur_book '+@CR+
                     'Deallocate cur_book '
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
       --Set @Msg = '�R�� F1 ��쬰�ťո�� ['+@TB_tmp_Name+']'

       --Set @strSQL = 'delete ['+@TB_tmp_Name+'] where isnull(f3, '''') ='''' and isnull(f6, '''') ='''' and isnull(f13, '''') ='''' '
       --Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
--=-[���ͳ�ک��Ӹ�ơA�ƶq���t�ƪ��n���Ͱh�f��]=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
print '19'
       IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_OD_Name+']') AND type in (N'U'))
       begin
print '20'
          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
          Set @Msg = '�R����ƪ� ['+@TB_OD_Name+']'
          Set @strSQL = 'DROP TABLE [dbo].['+@TB_OD_Name+']'
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       end

       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
       --2013/04/25 ���@���v����J�N�n���Υ[�`�A����n��
       --�b���i����ӫ~�s���A�|�����U�s���H�ΰӫ~�򥻸���ɶi����@�~
print '21'
       Set @Msg = '�i����ӫ~�s���üg�J ['+@TB_OD_Name+']��ƪ�C'
       set @strSQL = 'select rowid, ct_no, ct_sname, ct_ssname, '+@CR+
                     '       isnull(d.co_skno, d1.sk_no) as sk_no, '+@CR+
                     '       isnull(d.sk_name, d1.sk_name) as sk_name, '+@CR+
                     '       f6, '+@CR+
                     '       isnull(d.sk_bcode, d1.sk_bcode) as sk_bcode, '+@CR+
                     '       f3, '+@CR+
                          -- sum(case when f13 < 0 then f13 * -1 else f13 end) as sale_qty, 
                          -- sum(case when f18 < 0 then f18 * -1 else f18 end) as sale_amt
                     --'       case when Convert(Numeric(20, 6), f13) < 0 then Convert(Numeric(20, 6), f13) * -1 else Convert(Numeric(20, 6), f13) end as sale_qty, '+@CR+
                     '       Convert(Numeric(20, 6), F4) as sale_qty, '+@CR+
                     -- 2013/11/25 �p�_�����B�n�אּ�˺�^���|���B
                     '       (Convert(Numeric(20, 6), F5)) / (Convert(Numeric(20, 6), F4)) / 1.05 as sale_amt, '+@CR+ -- ���
                     '       (Convert(Numeric(20, 6), F5)) / 1.05 as sale_tot, '+@CR+ -- �p�p
                     '       case when isnull(d.co_skno, d1.sk_no) is null then ''N'' else ''Y'' end as isfound '+@CR+
                     '       into '+@TB_OD_Name+' '+@CR+
                     '  from (select * '+@CR+
                     '          from (select ct_no, rtrim(ct_sname) as ct_sname, rtrim(ct_name)+''#''+rtrim(ct_sname) as ct_ssname '+@CR+
                     '                  from SYNC_TA13.dbo.PCUST  '+@CR+
                     '                 where ct_class =''1''  '+@CR+
                     '                   and ct_name like '''+@CompanyLikeName+''' '+@CR+
                     '                   and substring(ct_no, 9, 1) =''3'' '+@CR+ -- �@�묰 1, 3C �� 2�A �ʳf�� 3
                     '                ) m, '+@CR+
                     '               (select * '+@CR+
                     '                  from ['+@TB_tmp_Name+'] '+@CR+
                     '                 where F1 like ''%'+@F1_Tag+'%'' '+@CR+
                     '                   and F1 <> F2 '+@CR+
                     '                   and F4+F5+F6+F7 <> '''' '+@CR+
                     '               ) d '+@CR+
                     '       )m '+@CR+
                     '       left join '+@CR+ 
                     '         (select distinct co_ctno, co_skno, co_cono, sk_no, sk_name, sk_bcode '+@CR+
                     '            from SYNC_TA13.dbo.sauxf m '+@CR+
                     '                 left join SYNC_TA13.dbo.sstock d '+@CR+
                     '                   on m.co_skno = d.sk_no '+@CR+
                     '           where co_class=''1'' '+@CR+
                     '         ) d '+@CR+
                     '          on 1=1 '+@CR+
                     '         and m.ct_no=d.co_ctno '+@CR+
                     '         and (ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.co_skno)) collate Chinese_Taiwan_Stroke_CI_AS  '+@CR+
                     '          or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.co_cono)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                     '          or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.sk_no)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                     '          or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.sk_bcode)) collate Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                     '        left join SYNC_TA13.dbo.sstock d1 '+@CR+
                     '          on 1=1 '+@CR+
                     '         and (ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d1.sk_no)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                     '          or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d1.sk_bcode)) collate Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                     ' order by 1 '
                     --group by ct_no, ct_sname, ct_ssname, sk_no, sk_name, sk_bcode, f3, f6
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

       --���ͥX�f�h�^��
print '22'
       set @Cnt = 0
       Set @Msg = '�ˬd['+@TB_OD_Name+']����J��ƬO�_�s�b�C'
       set @strSQL = 'select count(1) from ['+@TB_OD_Name+'] '
       delete @RowCount
       print @strSQL
       insert into @RowCount Exec (@strSQL)

       -- �]���ϥ� Memory Variable Table, �ҥH�ϥ� >0 �P�_
       select @Cnt=cnt from @RowCount
print '23'
       if @Cnt >0
       begin
print '24'
          set @Msg = @Msg + '...�s�b�A�N�i����J�{�ǡC'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', @Cnt
          --���o�̷s�X�f�h�^�渹(�����q�w�֭�H�Υ��֭㪺������)
          --�ФŦA�[+1�A�]���s�W�ɷ|�۰ʥ[�J
          --2013/11/24 �אּ @sd_date (�C��T�w�ϥ� 25�鰵�����ɤ�)
          set @new_sdno = (select convert(varchar(10), isnull(max(sd_no), replace(substring(@sd_date, 3, 8), '/', '') +'0000')) 
                             from (select distinct sd_no
                                     from SYNC_TA13.dbo.sslpdt m
                                    where sd_class = @sd_class
                                      and sd_slip_fg = @sd_slip_fg
                                      and sd_date = @sd_date
                                    union
                                   select distinct sd_no
                                     from SYNC_TA13.dbo.sslpdttmp m
                                    where sd_class = @sd_class
                                      and sd_slip_fg = @sd_slip_fg
                                      and sd_date = @sd_date
                                  ) m
                          )
          set @Msg = '���o�̷s�X�f�h�^�渹(�����q�w�֭�H�Υ��֭㪺������),new_sdno=['+@new_sdno+'], sd_class=['+@sd_class+'], sd_slip_fg=['+@sd_slip_fg+'], sd_date=['+@sd_date+'].'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', @Cnt
           
                         
          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '25'
          set @Msg = '�R��['+@sd_date+'] ����J����ک��ӡC'
          
          set @strSQL ='delete TYEIPDBS2.lytdbta13.dbo.sslpdttmp '+@CR+
                       ' where Convert(Varchar(10), sd_date, 111) = '''+@sd_date+''' '+@CR+
                       '   and sd_class = '''+@sd_class+''' '+@CR+
                       '   and sd_slip_fg = '''+@sd_slip_fg+''' '+@CR+
                       '   and sd_lotno like ''%'+@Rm+'%'' '
           
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '26'
          set @Msg = '�R��['+@sd_date+'] ����J����ڥD�ɡC'
          set @strSQL ='delete TYEIPDBS2.lytdbta13.dbo.ssliptmp '+@CR+
                       ' where Convert(Varchar(10), sp_date, 111) = '''+@sd_date+''' '+@CR+
                       '   and sp_class = '''+@sd_class+''' '+@CR+
                       '   and sp_slip_fg = '''+@sd_slip_fg+''' '+@CR+
                       '   and sp_rem like ''%'+@Rm+'%'' '
 
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '27'
          set @Msg = '�s�W['+@sd_date+'] ��X�f�h�^���ӡC'
          set @strSQL ='insert into TYEIPDBS2.lytdbta13.dbo.sslpdttmp '+@CR+
                       '(sd_class, sd_slip_fg, sd_date, sd_no, sd_ctno, '+@CR+
                       ' sd_skno, sd_name, sd_whno, sd_whno2, sd_qty, '+@CR+
                       ' sd_price, sd_dis, sd_stot, sd_lotno, sd_unit, '+@CR+
                       ' sd_unit_fg, sd_ave_p, sd_rate, sd_seqfld, sd_ordno '+@CR+
                       ') '+@CR+
                       'select '''+@sd_class+''' as sd_class, '+@CR+ -- ���O
                       '       '''+@sd_slip_fg+''' sd_slip_fg, '+@CR+ -- ��ں���
                       '       convert(datetime, '''+@sd_date+''') as sd_date, '+@CR+ -- �f����
                       '       d2.sd_no, '+@CR+ -- �f��s��
                       '       m.ct_no as sd_ctno, '+@CR+ -- �Ȥ�s��
                       ''+@CR+
                       '       d.sk_no as sd_skno, '+@CR+ -- �f�~�s��
                       '       d.sk_name as sd_name, '+@CR+ -- �~�W�W��
                       '       ''LB'' as sd_whno, '+@CR+ -- �ܮw(�J)
                       '       '''' as sd_whno2, '+@CR+ -- �ܮw(�X)
                       '       m.sale_qty as sd_qty, '+@CR+ -- ����ƶq
                       ''+@CR+
                       '       m.sale_amt as sd_price, '+@CR+ -- ����
                       '       1 as sd_dis, '+@CR+ -- ���
                       '       m.sale_tot as sd_stot, '+@CR+ -- �p�p
                       --2013/5/3 ���F������Ƹ�ƨæ^�g��Ƶ����A�ҥH��g�� sd_lotno �����ثe���G���ϥΡA�u�n��@�{�ɳƵ��ϥ�
                       --'       '''+@Rm+'�� ''+Convert(Varchar(20), Getdate(), 120)+char(13)+char(10)+'''+@OD_CName +''' as sd_rem, '+@CR+ -- �ƥ����
                       '       '''+@Rm+@OD_CName +''' as sd_lotno, '+@CR+ -- �ƥ����
                       '       d.sk_unit as sd_unit, '+@CR+ -- ���
                       ' '+@CR+
                       '       0 as sd_unit_fg, '+@CR+ -- ���X��
                       '       d.s_price4 as sd_ave_p, '+@CR+ -- ��즨��
                       '       1 as sd_rate, '+@CR+ -- �ײv
                       '       rowid as sd_seqfld, '+@CR+ -- ���ӧǸ�, ���Ǹ��|�]����V�ק�ӥ[�H�ܧ�A�ҥH�t�s�@���� sd_ordno
                       '       rowid as sd_ordno'+@CR+ -- XLS ���ӧǸ�
                       --'       row_number() over(order by m.ct_no, d.sk_no) as sd_seqfld '+@CR+
                       '  from ['+@TB_OD_Name+'] m '+@CR+
                       '       left join SYNC_TA13.dbo.sstock d '+@CR+
                       '         on m.sk_no = d.sk_no '+@CR+
                       '       left join SYNC_TA13.dbo.pcust d1 '+@CR+
                       '         on m.ct_no = d1.ct_no '+@CR+
                       '        and ct_class =''1'' '+@CR+
                       '       left join  '+@CR+
                       '        (select ct_no,  '+@CR+
                       '                convert(varchar(10), '+@new_sdno+'+row_number() over(order by m.ct_no)) as sd_no '+@CR+ -- �f��s��
                       '           from (select distinct m.ct_no '+@CR+
                       '                   from ['+@TB_OD_Name+'] m  '+@CR+
                       '                        left join SYNC_TA13.dbo.pcust d1  '+@CR+
                       '                          on m.ct_no = d1.ct_no  '+@CR+
                       '                         and ct_class =''1'' '+@CR+
                       '                 )m '+@CR+
                       '         ) d2 on d1.ct_no = d2.ct_no '+@CR+
                       ' where isfound =''Y'' '+@CR+
                       ' Order by 1, 2, 3, m.ct_no'
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
          
          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '28'
          set @Msg = '�^�g�ӫ~��ﭫ�ƳƵ��C'
          set @strSQL ='update TYEIPDBS2.lytdbta13.dbo.sslpdttmp '+@CR+
                       '   set sd_rem='''+@Rm+@OD_CName+'�ӫ~��ﭫ�� RowID:''+D.Rowid '+@CR+
                       '  from TYEIPDBS2.lytdbta13.dbo.sslpdttmp M '+@CR+
                       '       inner join '+@CR+
                       '         (select Convert(Varchar(100), rowid) as rowid, count(1) as cnt '+@CR+
                       '            from ['+@TB_OD_Name+'] '+@CR+
                       '           group by rowid '+@CR+
                       '          having count(*) >1 '+@CR+
                       '         ) D '+@CR+
                       '          on sd_seqfld=D.rowid '+@CR+
                       ' where Convert(Varchar(10), m.sd_date, 111) = '''+@sd_date+''' '+@CR+
                       '   and m.sd_class = '''+@sd_class+''' '+@CR+
                       '   and m.sd_slip_fg = '''+@sd_slip_fg+''' '+@CR+
                       '   and m.sd_lotno like ''%'+@Rm+'%'' '
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
                       
          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
          -- �s�W�X�f�h�^�D�� PKey: sp_no (ASC), sp_slip_fg (ASC)
print '29'
          set @Msg = '�s�W�X�f�h�^�D�ɡC'
          set @strSQL ='insert into TYEIPDBS2.lytdbta13.dbo.ssliptmp '+@CR+
                       '(sp_class, sp_slip_fg, sp_date, sp_pdate, sp_no, '+@CR+
                       ' sp_ctno,  sp_ctname, sp_ctadd2, sp_sales, sp_dpno, '+@CR+
                       ' sp_maker, sp_conv, sp_tot, sp_tax, sp_dis, '+@CR+
                       ' sp_pay_kd, sp_rate_nm, sp_rate, sp_itot, sp_inv_kd, '+@CR+
                       ' sp_tax_kd, sp_invtype, sp_rem '+@CR+
                       ') '+@CR+
                       'select distinct sd_class as sp_class, '+@CR+
                       '       sd_slip_fg as sp_slip_fg, '+@CR+
                       '       sd_date as sp_date, '+@CR+
                       '       dbo.uFn_Getdate(ct_pmode, sd_date, ct_pdate) as sp_pdate, '+@CR+
                       /*
                       '       case '+@CR+
                       '         when day(sd_date) <=  ct_pdate then dateadd(day, ct_pdate - day(sd_date), sd_date) '+@CR+
                       '         else Convert(DateTime, '+@CR+
                       '                Convert(Varchar(4), year(sd_date))+''/''+ '+@CR+
                       '                Convert(varchar(2), month(dateadd(mm, 1, sd_date)))+''/''+ '+@CR+
                       '                Convert(varchar(2), ct_pdate)) '+@CR+
                       '       end as sp_pdate, '+@CR+
                       */
                       '       sd_no as sp_no, '+@CR+
                       ' '+@CR+
                       '       d1.ct_no as sp_ctno, '+@CR+ -- �Ȥ�s��
                       '       d1.ct_name as sp_ctname, '+@CR+ -- �Ȥ�W��
                       '       substring(d1.ct_addr3, 1,255) as sp_ctadd2, '+@CR+ -- �e�f�a�}
                       '       d1.ct_sales as sp_sales, '+@CR+ -- �~�ȭ�
                       '       d1.ct_dept as sp_dpno, '+@CR+ -- �����s��
                       ' '+@CR+
                       '       '''+@sp_maker+''' as sp_maker, '+@CR+ -- �s��H��
                       '       d1.ct_porter as sp_conv, '+@CR+ -- �f�B���q
                       '       sum(sd_stot) as sp_tot, '+@CR+ -- �p�p
                       
                       -- 2013/6/10 ����^�f���ݭn�|��
                       '       case when sd_class =''3'' and sd_slip_fg=''7'' then 0 else sum(sd_stot)*0.05 end as sp_tax, '+@CR+ -- ��~�|(��)
                       
                       '       0 as sp_dis, '+@CR+ -- �������B(��)
                       ' '+@CR+
                       '       d1.ct_payfg as sp_pay_kd, '+@CR+ -- �������
                       '       ''NT'' as sp_rate_nm, '+@CR+ -- �ײv�W��
                       '       1 as sp_rate, '+@CR+ -- �ײv
                       '       sum(sd_stot) as sp_itot, '+@CR+ -- �o�����B
                       '       1 as sp_inv_kd, '+@CR+ -- �o�����O(=1  �T�p��,=2  �G�p��,=3  ���Ⱦ�)
                       ' '+@CR+
                       '       1 as sp_tax_kd, '+@CR+ -- �|�O(=1���|,=2�s�| )
                       '       1 as sp_invtype, '+@CR+ -- �}�ߤ覡(=1���}, =2�H��}��, =3�妸�}��)
                       '       '''+@Rm+'�� ''+Convert(Varchar(20), Getdate(), 120)+char(13)+char(10)+'''+@OD_CName+''' as sp_rem '+@CR+ -- �Ƶ�
                       '  from TYEIPDBS2.lytdbta13.dbo.sslpdttmp m '+@CR+
                       '       left join SYNC_TA13.dbo.sstock d '+@CR+
                       '         on m.sd_skno = d.sk_no '+@CR+
                       '       left join SYNC_TA13.dbo.pcust d1 '+@CR+
                       '         on m.sd_ctno = d1.ct_no '+@CR+
                       '        and ct_class =''1'' '+@CR+
                       ' where sd_class = '''+@sd_class+''' '+@CR+
                       '   and sd_slip_fg = '''+@sd_slip_fg+''' '+@CR+
                       '   and sd_date = '''+@sd_date+''' '+@CR+
                       '   and sd_lotno like '''+@RM+'%'+@OD_CName+'%'' '+@CR+
                       ' group by sd_class, '+@CR+
                       '       sd_slip_fg, '+@CR+
                       '       sd_date, '+@CR+
                       '       ct_pmode, '+@CR+
                       '       ct_pdate, '+@CR+
                       '       sd_no, '+@CR+
                       '       d1.ct_no, '+@CR+ -- �Ȥ�s��
                       '       d1.ct_name, '+@CR+ -- �Ȥ�W��
                       '       substring(d1.ct_addr3, 1, 255), '+@CR+ -- �e�f�a�}
                       '       d1.ct_sales, '+@CR+ -- �~�ȭ�
                       '       d1.ct_dept, '+@CR+ -- �����s��
                       '       d1.ct_porter, '+@CR+ -- �f�B���q
                       '       d1.ct_payfg '

          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
          -- 2015/11/25 Rickliu �W�[�R���L�Y�����Ӹ�ơA�קK�ȪA�H���H�u���X�f�h�^�����ɮɡA���ͦ]���ɥ��Ѫ���ƥX�{�C
          -- 2017/07/17 Rickliu �۱q��ƦP�B�אּ�q�\�覡��A������G�A��DB�i�� Update �|�`�A�ʥ��ѨåB���ͥH�U���~�T���A��L�䤣����D�Ҧb�A�u�n��g�� Cursor �覡
          -- ���~�T��:�L�k�q�s�����A�� "TYEIPDBS2" �� OLE DB ���Ѫ� "SQLNCLI10" ���o��ƦC����ơC
          set @Msg = '�R�� ['+@sd_date+'] �L�Y���X�f�h�^�����ɡC'
          --set @strSQL ='delete TYEIPDBS2.LYTDBTA13.dbo.sslpdttmp '+@CR+
          --             '  from TYEIPDBS2.LYTDBTA13.dbo.sslpdttmp m '+@CR+
          --             ' where not exists '+@CR+
          --             '       (select * '+@CR+
          --             '          from TYEIPDBS2.LYTDBTA13.dbo.ssliptmp d '+@CR+
          --             '         where m.sd_class = d.sp_class '+@CR+
          --             '           and m.sd_slip_fg = d.sp_slip_fg '+@CR+
          --             '           and m.sd_no = d.sp_no '+@CR+
          --             '       ) '+@CR+
          --             '   and m.sd_class = '''+@sd_class+''' '+@CR+
          --             '   and m.sd_slip_fg = '''+@sd_slip_fg+''' '+@CR+
          --             '   and m.sd_date = '''+@sd_date+''' '+@CR+
          --             '   and m.sd_lotno like '''+@RM+'%'+@OD_CName+'%'' '


          set @strSQL ='Declare @sd_class as varchar(100) '+@CR+
                       'Declare @sd_slip_fg as varchar(100) '+@CR+
                       'Declare @sd_no as varchar(100) '+@CR+
                       'Declare @cnt as int = 0 '+@CR+
                       ''+@CR+
                       'Declare Cur_'+@TB_OD_Name+'_Del_No_Header Cursor for '+@CR+
                       '  select m.sd_class, m.sd_slip_fg, m.sd_no, count(distinct d.sp_no) as cnt '+@CR+
                       '    from TYEIPDBS2.lytdbta13.dbo.sslpdttmp m '+@CR+
                       '         left join TYEIPDBS2.lytdbta13.dbo.ssliptmp d '+@CR+
                       '           on m.sd_class = d.sp_class '+@CR+
                       '          and m.sd_slip_fg = d.sp_slip_fg '+@CR+
                       '          and m.sd_no = d.sp_no '+@CR+
                       '   where 1=1 '+@CR+
                       '     and m.sd_class = '''+@sd_class+''' '+@CR+
                       '     and m.sd_slip_fg = '''+@sd_slip_fg+''' '+@CR+
                       '     and m.sd_date = '''+@sd_date+''' '+@CR+
                       '     and m.sd_lotno like ''%'+@RM+'%'+@OD_CName+'%'' '+@CR+
                       '   group by m.sd_class, m.sd_slip_fg, m.sd_no '+@CR+
                       '  having count(distinct d.sp_no) > 1 '+@CR+
                       ''+@CR+
                       'Open Cur_'+@TB_OD_Name+'_Del_No_Header '+@CR+
                       'Fetch Next From Cur_'+@TB_OD_Name+'_Del_No_Header into @sd_class, @sd_slip_fg, @sd_no, @cnt '+@CR+
                       ''+@CR+
                       'While @@Fetch_status = 0 '+@CR+
                       'begin '+@CR+
                       '  Delete TYEIPDBS2.lytdbta13.dbo.sslpdttmp '+@CR+
                       '   where 1=1 '+@CR+
                       '     and sd_class = ''@sd_class'' '+@CR+
                       '     and sd_slip_fg = ''@sd_slip_fg'' '+@CR+
                       '     and sd_no = ''@sd_no'' '+@CR+
                       ''+@CR+
                       '  Fetch Next From Cur_'+@TB_OD_Name+'_Del_No_Header into @sd_class, @sd_slip_fg, @sd_no, @cnt '+@CR+
                       'end '+@CR+
                       ''+@CR+
                       'Close Cur_'+@TB_OD_Name+'_Del_No_Header '+@CR+
                       'Deallocate Cur_'+@TB_OD_Name+'_Del_No_Header '


          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       end
       else
       begin
print '30'
          set @Msg = @Msg + '...���s�b�A�פ���J�{�ǡC'
          Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

          -- 2013/11/28 �W�[���Ѧ^�ǭ�
          --fetch next from Cur_Mall_Consign_Order_PCHome_DataKind into @OD_Name, @OD_CName, @sd_class, @sd_slip_fg, @oper, @F1_Tag
       end

print '31'
    end
    fetch next from Cur_Mall_Consign_Order_PCHome_DataKind into @OD_Name, @OD_CName, @sd_class, @sd_slip_fg, @oper, @F1_Tag
  
print '32'
  end

print '33'
  close Cur_Mall_Consign_Order_PCHome_DataKind
  deallocate Cur_Mall_Consign_Order_PCHome_DataKind
  -- 2013/11/28 �W�[���Ѧ^�ǭ�
  Return(0)

print '34'
end
GO
