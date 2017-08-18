USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_Stock_taking_TA13]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_Imp_Stock_taking_TA13]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_Imp_Stock_taking_TA13]
as
begin
  /***********************************************************************************************************
     2013/12/06 -- Rickliu
     �L�I����J�A�Ш̽L�I Excel �ɮפγW�d�i����J�C     
  ***********************************************************************************************************/
  Declare @Proc Varchar(50) = 'uSP_Imp_Stock_taking_TA13'
  Declare @Cnt Int =0
  Declare @RowCnt Int =0
  Declare @Msg Varchar(4000) =''
  Declare @Errcode int = -1

  Declare @Kind Varchar(1), @KindName Varchar(10), @SP_Class Varchar(1), @sp_slip_fg Varchar(1)
  Declare @xls_head Varchar(100), @TB_head Varchar(200), @TB_xls_Name Varchar(200), @TB_SP_Name Varchar(200), @TB_tmp_name Varchar(200)

  Declare @strSQL Varchar(Max)
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @RowCount Table (cnt int)
  Declare @RowData Table (aData Varchar(255))
  
  Declare @new_orno varchar(10)

  Declare @SP_date varchar(10) -- �L�I���
  Declare @SP_Sales varchar(10) -- �s���
  Declare @SP_dpno varchar(10) -- �����s��
  Declare @SP_whno varchar(10) --�ܧO

  Declare @SP_cnt int
       
  Declare @CompanyName Varchar(100), @CompanyLikeName Varchar(100), @Rm Varchar(100), @SP_maker Varchar(20), @str Varchar(200)
  Declare @Pos Int
  
  Set @CompanyName = '���T�{�L�I��' --> �ФŶ��ܰ�
  Set @CompanyLikeName = '%'+@CompanyName+'%'
  Set @Rm = '�t�ζפJ'+@CompanyName
  Set @SP_maker = 'Admin'

print '1'
  Set @xls_head = 'Ori_xls#'
  Set @TB_head = 'Stock_taking_TA13'
  Set @TB_xls_Name = @xls_head+@TB_Head -- Ex: Ori_xls#Stock_taking_TA13 -- Excel ���ɸ��
  Set @TB_tmp_Name = @TB_Head+'_'+'tmp' -- Ex: Stock_taking_TA13_tmp -- �{����J����
  Set @TB_SP_Name = @TB_Head -- Ex: Stock_taking_TA13
  set @SP_Class = '7' -- �L�I��
  set @sp_slip_fg = 'T' -- �L�I��

  -----------------------------------------------------------------------------------------------------------------------------
  -- Check �פJ�ɮ׬O�_�s�b
  IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_xls_Name+']') AND type in (N'U'))
  begin
print '4'
     set @Cnt = @Errcode
     set @strSQL = ''
     Set @Msg = '�~�� Excel �פJ��ƪ� ['+@TB_xls_Name+'] ���s�b�A�פ�i�����ɧ@�~�C'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

     Return(@Errcode)
  end
    
  -----------------------------------------------------------------------------------------------------------------------------
  -- �P�O Excel �O�_�����
  set @Cnt = 0
  set @strSQL ='Select Count(1) as cnt from [dbo].['+@TB_xls_Name+']'
  set @Msg = '�P�O Excel �O�_�����...'
  delete @RowData
  insert into @RowCount exec (@strSQL)
  if (Select cnt from @RowCount) > 0
  begin
     set @Msg = @Msg + '�s�b'
  end
  else
  begin
     set @Cnt = @Errcode
     set @Msg = @Msg + '���s�b'
  end
  
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

  if @Cnt = @Errcode  Return(@Errcode)
    
  -----------------------------------------------------------------------------------------------------------------------------
  -- �P�O Excel �L�I��� ���n���O�_�s�b 
print '5'
  set @Cnt = 0
  set @SP_Date  = ''

  set @strSQL = 'Select F2 as sp_Date from [dbo].['+@TB_xls_Name+'] where Rowid = 1 '
  set @Msg = '�P�O Excel [�L�I���] �O�_�s�b'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

  delete @RowData
  insert into @RowData exec (@strSQL)
  select @SP_Date=Rtrim(isnull(aData, '')) from @RowData

  if @SP_Date = ''
  begin
print '6'
     set @Cnt = @Errcode
     set @strSQL = ''
     Set @Msg = '�䤣�� Excel ��Ƥ��� [�L�I���]�A�פ�i�����ɧ@�~�C'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

     Return(@Errcode)
  end
  -----------------------------------------------------------------------------------------------------------------------------
  -- �P�O Excel �L�I�� ���n���O�_�s�b 
print '5'
  set @Cnt = 0
  set @SP_Sales  = ''

  set @strSQL = 'Select F4 as sp_sales from [dbo].['+@TB_xls_Name+'] where Rowid = 1 '
  set @Msg = '�P�O Excel [�L�I��] �O�_�s�b'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

  delete @RowData
  insert into @RowData exec (@strSQL)
  select @SP_Sales=Rtrim(isnull(aData, '')) from @RowData
    
  if @SP_Sales = ''
  begin
print '6'
     set @Cnt = @Errcode
     set @strSQL = ''
     Set @Msg = '�䤣�� Excel ��Ƥ��� [�L�I��]�A�פ�i�����ɧ@�~�C'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

     Return(@Errcode)
  end
  -----------------------------------------------------------------------------------------------------------------------------
  -- �P�O Excel �����s�� ���n���O�_�s�b 
print '5'
  set @Cnt = 0
  set @SP_dpno  = ''

  set @strSQL = 'Select F6 as sp_dpno from [dbo].['+@TB_xls_Name+'] where Rowid = 1 '
  set @Msg = '�P�O Excel [�����s��] �O�_�s�b'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

  delete @RowData
  insert into @RowData exec (@strSQL)
  select @SP_dpno=Rtrim(isnull(aData, '')) from @RowData

  if @SP_dpno = ''
  begin
print '6'
     set @Cnt = @Errcode
     set @strSQL = ''
     Set @Msg = '�䤣�� Excel ��Ƥ��� [�����s��]�A�פ�i�����ɧ@�~�C'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

     Return(@Errcode)
  end
print '7'
     IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_tmp_Name+']') AND type in (N'U'))
     begin
print '8'
        Set @Msg = '���� '+@CompanyName+' �{����J������ƪ� ['+@TB_tmp_Name+']�C'

        Set @strSQL = 'DROP TABLE [dbo].['+@TB_tmp_Name+']'
        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
     end
     --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
     --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
     -- �R���e�ⵧ���Y���
print '9'
     Set @Msg = '�R���e�ⵧ���Y��ƥH�ΧR���ť���� ['+@TB_tmp_Name+']�C'
     Set @strSQL = 'select Isnull(LTrim(Rtrim(F1)), '''') as F1, '+@CR+ -- �ӫ~�s��
                   '       Isnull(LTrim(Rtrim(F2)), '''') as F2, '+@CR+ -- �ӫ~�W��
                   '       Isnull(LTrim(Rtrim(F3)), '''') as F3, '+@CR+ -- �ܧO
                   '       Case '+@CR+
                   '         when LTrim(RTrim(F4)) = '''' then ''0'' '+@CR+
                   '         else Isnull(LTrim(Rtrim(F4)), ''0'') '+@CR+
                   '       end as F4, '+@CR+ -- �L�I�ƶq
                   '       sp_date='''+@SP_date+''', '+@CR+ -- �L�I���
                   '       sp_sales='''+@SP_Sales+''', '+@CR+ -- �L�I��
                   '       sp_dpno='''+@SP_dpno+''', '+@CR+ -- �����s��
                   '       Isnull(LTrim(Rtrim(xlsFileName)), '''') as xlsFileName, '+@CR+ -- �ɮץN�� ����ƥ�
                   '       import_date=convert(date, getdate()), '+@CR+
                   '       rowid '+@CR+ -- Xls �C��
                   '       into '+@TB_tmp_Name+@CR+
                   '  from ['+@TB_xls_Name+']'+@CR+
				   ' where f3 <> ''�L�I��'' and f3 <> ''�ܧO'' '
     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

print '10'

  Declare Cur_Stock_taking_TA13 cursor for
    select distinct f3 as sp_whno
      from [DW].[dbo].[Ori_xls#Stock_taking_TA13]
	 where f3 <> '�L�I��' and f3 <> '�ܧO'

  open Cur_Stock_taking_TA13
  fetch next from Cur_Stock_taking_TA13 into @SP_whno
  
  Set @Rm = '�t�ζפJ'+@CompanyName + '_' + @SP_whno

print '11'

  while @@fetch_status =0
  begin   
  -- �W�[�P�O �T�{ �� ���T�{ ��� �Y�O�s�b�h���o�i�歫��@�~�C
  -- 2013/10/28 �W�[�P�_�Y���24��w�ѥ��T�{��T�{��h���A����A�קK��V���T�{��ڤ��_�R���s�W�y���渹�ּW���p
print '12'
  select @SP_cnt = sum(cnt)
    from (select count(1) as cnt
           from SYNC_TA13.dbo.sslip m
          where sp_class = @SP_class
            and sp_slip_fg = @sp_slip_fg
            and sp_date = @SP_date
            and sp_rem like '%'+@Rm+'%'
          union
         select count(1) as cnt
           from SYNC_TA13.dbo.ssliptmp m
          where sp_class = @SP_class
            and sp_slip_fg = @sp_slip_fg
            and sp_date = @SP_date
            and sp_rem like '%'+@Rm+'%') m

  if @SP_cnt > 0
  begin
print '13'
     set @Cnt = @Errcode
     set @strSQL = ''
     set @Msg ='[ �w�T�{ �� ���T�{ �� ] �w�� ['+@SP_Date+'] ��ơA�Y�ݭn����h����M���T�{�Υ��T�{��ګ᭫��Y�i�Csslip, @SP_class=['+@SP_class+'], @SP_slip_fg=['+@SP_slip_fg+'], @SP_date=['+@SP_date+'], @SP_lotno=['+@RM+']�C'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
     
     Return(@Errcode)
  end
  else
  begin



     --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
     --���ͥX�f�h�^��
print '22'
     set @Cnt = 0
     Set @Msg = '�ˬd['+@TB_tmp_Name+']��ڸ�ƬO�_�s�b�C'
     set @strSQL = 'select count(1) from ['+@TB_tmp_Name+']'
     delete @RowCount
     print @strSQL
     insert into @RowCount Exec (@strSQL)

     -- �]���ϥ� Memory Variable Table, �ҥH�ϥ� >0 �P�_
print '23'
     if (select cnt from @RowCount) > 0
     begin
print '24'
        set @Msg = @Msg + '...���s�b�A�N�i����J�{�ǡC'
        Exec uSP_Sys_Write_Log @Proc, @Msg, '', @Cnt
        --���o�̷s�X�f�h�^�渹(�����q�w�֭�H�Υ��֭㪺������)
        --�ФŦA�[+1�A�]���s�W�ɷ|�۰ʥ[�J
        --2013/11/24 �אּ @SP_date (�C��T�w�ϥ� 25�鰵�����ɤ�)
        set @new_orno = (select convert(varchar(10), isnull(max(SP_no)+1, replace(substring(@SP_date, 3, 8), '/', '') +'0000')) 
                           from (select distinct sp_no
                                   from SYNC_TA13.dbo.sslip m
                                  where SP_class = @SP_class
                                    and SP_slip_fg = @sp_slip_fg
                                    and SP_date = @sp_date
                                  union
                                 select distinct sp_no
                                   from SYNC_TA13.dbo.ssliptmp m
                                  where SP_class = @SP_class
                                    and SP_slip_fg = @sp_slip_fg
                                    and SP_date = @sp_date
                                ) m
                        )
        set @Msg = '���o�̷s�L�I�渹(�����q�w�֭�H�Υ��֭㪺������),new_orno=['+@new_orno+'], sp_class=['+@SP_class+'], sp_slip_fg=['+@SP_slip_fg+'], sp_date=['+@SP_date+'].'
        Exec uSP_Sys_Write_Log @Proc, @Msg, '', @Cnt

        --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '25'
        set @Msg = '�R�� ['+@SP_date+'] ����J���L�I��ک��ӡC'
        
        set @strSQL ='delete TYEIPDBS2.lytdbta13.dbo.sslpdttmp '+@CR+
                     ' where Convert(Varchar(10), sd_date, 111) = '''+@SP_date+''' '+@CR+
                     '   and sd_class = '''+@SP_class+''' '+@CR+
                     '   and sd_slip_fg = '''+@sp_slip_fg+''' '+@CR+
                     '   and sd_lotno like ''%'+@Rm+'%'' '
           
        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '26'
        set @Msg = '�R�� ['+@SP_date+'] ����J���q���ڥD�ɡC'
        set @strSQL ='delete TYEIPDBS2.lytdbta13.dbo.ssliptmp '+@CR+
                     ' where Convert(Varchar(10), sp_date, 111) = '''+@SP_date+''' '+@CR+
                     '   and sp_class = '''+@SP_class+''' '+@CR+
                     '   and sp_slip_fg = '''+@sp_slip_fg+''' '+@CR+
                     '   and sp_rem like ''%'+@Rm+'%'' '
 
        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

        --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '27'
        set @Msg = '�s�W ['+@SP_date+'] �L�I��ک��ӡC'
        set @strSQL ='insert into TYEIPDBS2.lytdbta13.dbo.sslpdttmp '+@CR+
                     '(sd_class, sd_slip_fg, sd_date, sd_no, sd_ctno, '+@CR+
                     ' sd_skno, sd_name, sd_whno, sd_whno2, sd_qty, '+@CR+
                     ' sd_price, sd_dis, sd_stot, sd_lotno, sd_rem , sd_unit, '+@CR+
                     ' sd_unit_fg, sd_ave_p, sd_rate, sd_seqfld, sd_ordno '+@CR+
                     ') '+@CR+
                     'select '''+@sp_class+''' as sd_class, '+@CR+ -- ���O
                     '       '''+@sp_slip_fg+''' sd_slip_fg, '+@CR+ -- ��ں���
                     '       convert(datetime, '''+@sp_date+''') as sd_date, '+@CR+ -- �f����
                     '       '''+@new_orno+''' as sd_no, '+@CR+ -- �f��s��
                     '       '''' as sd_ctno, '+@CR+ -- �Ȥ�s��
                     ''+@CR+
                     '       sk_no as sd_skno, '+@CR+ -- �f�~�s��
                     '       sk_name as sd_name, '+@CR+ -- �~�W�W��
                     '       isnull(F3, '''') as sd_whno, '+@CR+ -- �ܮw(�J)
                     '       '''' as sd_whno2, '+@CR+ -- �ܮw(�X)
                     '       isnull(F4, 0)-isnull(wd_qty, 0) as SD_qty, '+@CR+ -- �t���ƶq
                     ''+@CR+
                     '       isnull(F4, 0) as SD_PRICE, '+@CR+ -- �L�I�ƶq
                     '       isnull(wd_qty, 0) as SD_DIS, '+@CR+ -- �ثe�w�s
                     '       0 as sd_stot, '+@CR+ -- �p�p
                     '       '''+@Rm+''' as sd_lotno, '+@CR+ -- �ƥ����
                     '       m.xlsFileName  as sd_rem , '+@CR+ -- �ƥ����
                     '       sk_unit as SD_UNIT, '+@CR+ -- ���
                     ' '+@CR+
                     '       0 as sd_unit_fg, '+@CR+ -- ���X��
                     '       d.sk_save as sd_ave_p, '+@CR+ -- ��즨��(�зǦ���)
                     '       1 as sd_rate, '+@CR+ -- �ײv
                     '       row_number() over(partition by '''+@new_orno+''' order by d1.wd_skno) as sd_seqfld, '+@CR+ -- ���ӧǸ�, ���Ǹ��|�]����V�ק�ӥ[�H�ܧ�A�ҥH�t�s�@���� sd_ordno
                     '       rowid as sd_ordno'+@CR+ -- XLS ���ӧǸ�
                     
                     '  from '+@TB_tmp_Name+' m '+@CR+
                     '       inner join SYNC_TA13.dbo.sstock d '+@CR+
                     '          on m.F1 collate Chinese_PRC_BIN = d.sk_no collate Chinese_PRC_BIN '+@CR+
                     '        left join '+@CR+
                     '         (select m.wd_no, m.wd_skno, '+@CR+
                     '                 wd_qty=sum(wd_amt0+wd_amt1+wd_amt2+wd_amt3+wd_amt4+wd_amt5+wd_amt6+wd_amt7+wd_amt8+wd_amt9+wd_amt10+wd_amt11+wd_amt12) '+@CR+
                     '            from SYNC_TA13.dbo.swaredt m'+@CR+
                     '                 inner join '+@CR+
                     '                   (select wd_no, wd_skno, max(wd_yr) as wd_yr '+@CR+
                     '                      from SYNC_TA13.dbo.swaredt '+@CR+
                     '                     group by wd_no, wd_skno '+@CR+
                     '                   ) d on m.wd_no=d.wd_no '+@CR+
                     '                      and m.wd_skno=d.wd_skno '+@CR+
                     '                      and m.wd_yr=d.wd_yr '+@CR+
                     '           where m.wd_class = ''0'' '+@CR+
                     --'             and wd_yr = '''+SUBSTRING(@SP_date, 1, 4)+''' '+@CR+
                     '           group by m.wd_no, m.wd_skno '+@CR+
                     '         ) d1 '+@CR+
                     '       on Rtrim(m.F1) collate Chinese_PRC_BIN = Rtrim(d1.wd_skno) collate Chinese_PRC_BIN '+@CR+
                     '      and upper(ltrim(rtrim(m.F3))) collate Chinese_PRC_BIN = upper(ltrim(rtrim(wd_no))) collate Chinese_PRC_BIN '+@CR+
                     '      and F1 <> '''' '+@CR+
                     '      and F3 <> '''' '+@CR+
                     --' where m.rowid > 2 '
                     ' where 1=1 '+@CR+
					 '   and m.F3 = ''' + @SP_whno + ''' '

        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
        

        --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
        -- �s�W�X�f�h�^�D�� PKey: sp_no (ASC), sp_slip_fg (ASC)
print '29'
        set @Msg = '�s�W ['+@SP_date+'] �L�I��ڥD�ɡC'
        set @strSQL ='insert into TYEIPDBS2.lytdbta13.dbo.ssliptmp '+@CR+
                     '(sp_class, sp_slip_fg, sp_date, sp_pdate, sp_no, '+@CR+
                     ' sp_ctno,  sp_ctname, sp_ctadd2, sp_sales, sp_dpno, '+@CR+
                     ' sp_maker, sp_conv, sp_tot, sp_tax, sp_dis, '+@CR+
                     ' sp_pay_kd, sp_rate_nm, sp_rate, sp_itot, sp_inv_kd, '+@CR+
                     ' sp_tax_kd, sp_invtype, sp_rem, sp_tal_rec '+@CR+
                     ') '+@CR+
                     'select distinct sd_class as sp_class, '+@CR+
                     '       sd_slip_fg as sp_slip_fg, '+@CR+
                     '       sd_date as sp_date, '+@CR+
                     '       '''' as sp_pdate, '+@CR+
                     '       sd_no as sp_no, '+@CR+
                     ' '+@CR+
                     '       '''' as sp_ctno, '+@CR+ -- �Ȥ�s��
                     '       '''' as sp_ctname, '+@CR+ -- �Ȥ�W��
                     '       '''' as sp_ctadd2, '+@CR+ -- �e�f�a�}
                     '       '''+@SP_Sales+''' as sp_sales, '+@CR+ -- �~�ȭ�
                     '       '''+@SP_dpno+''' as sp_dpno, '+@CR+ -- �����s��
                     ' '+@CR+
                     '       '''+@sp_maker+''' as sp_maker, '+@CR+ -- �s��H��
                     '       '''' as sp_conv, '+@CR+ -- �f�B���q
                     '       sum(sd_stot) as sp_tot, '+@CR+ -- �p�p
                     
                     '       0 as sp_tax, '+@CR+ -- ��~�|(��)
                     
                     '       1 as sp_dis, '+@CR+ -- �������B(��)
                     ' '+@CR+
                     '       '''' as sp_pay_kd, '+@CR+ -- �������
                     '       ''NT'' as sp_rate_nm, '+@CR+ -- �ײv�W��
                     '       1 as sp_rate, '+@CR+ -- �ײv
                     '       sum(sd_stot) as sp_itot, '+@CR+ -- �o�����B
                     '       1 as sp_inv_kd, '+@CR+ -- �o�����O(=1  �T�p��,=2  �G�p��,=3  ���Ⱦ�)
                     ' '+@CR+
                     
                     '       1 as sp_tax_kd, '+@CR+ -- �|�O(=1���|,=2�s�| )
                     '       1 as sp_invtype, '+@CR+ -- �}�ߤ覡(=1���}, =2�H��}��, =3�妸�}��)
                     '       '''+@Rm+'�� ''+Convert(Varchar(20), Getdate(), 120)+char(13)+char(10) as sp_rem, '+@CR+ -- �Ƶ�
                     '       count(1) as sp_tal_rec '+@CR+
                     '  from TYEIPDBS2.lytdbta13.dbo.sslpdttmp m '+@CR+
                     '       left join SYNC_TA13.dbo.sstock d '+@CR+
                     '         on m.sd_skno = d.sk_no '+@CR+
                     ' where sd_class = '''+@sp_class+''' '+@CR+
                     '   and sd_slip_fg = '''+@sp_slip_fg+''' '+@CR+
                     '   and sd_date = '''+@sp_date+''' '+@CR+
                     '   and sd_lotno like ''%'+@RM+'%'' '+@CR+
                     ' group by sd_class, '+@CR+
                     '       sd_slip_fg, '+@CR+
                     '       sd_date, '+@CR+
                     '       sd_no '

        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

        --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
        -- �@�~�����A�M���L�I���l��J���
print '30'
     end
     else
     begin
print '31'
        set @Msg = @Msg + '...�w�s�b�A�פ���J�{�ǡC'
        Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
        set @Cnt = @Errcode
     end

	 
	 end


        fetch next from Cur_Stock_taking_TA13 into @SP_whno
		
        Set @Rm = '�t�ζפJ'+@CompanyName + '_' + @SP_whno
print '31'
  end
proc_exit:
print '32'
  close Cur_Stock_taking_TA13
  deallocate Cur_Stock_taking_TA13
  set @Cnt = 0
  set @SP_Date  = ''

  set @strSQL = 'Delete [dbo].['+@TB_xls_Name+'] '
  set @Msg = '�@�~�����A�M���L�I���l��J���!!'
  Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  --Return(0)
print '33'
  Return(@Cnt)

end
GO
