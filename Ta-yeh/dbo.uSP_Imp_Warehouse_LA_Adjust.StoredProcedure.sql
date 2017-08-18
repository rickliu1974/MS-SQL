USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_Warehouse_LA_Adjust]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_Imp_Warehouse_LA_Adjust]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_Imp_Warehouse_LA_Adjust]
as
begin
  /***********************************************************************************************************
    2014/10/14 Rickliu 
    �C��T�w 10 �鲣�ͽվ��ڦ� TA13�A�ӽվ��ڶȰw�� LA �ܡA�ñN�Ҧ��~���վ㬰�s�A�վ��ڶȲ��ͥ��T�{��     
  ***********************************************************************************************************/
  Declare @Proc Varchar(50) = 'uSP_Imp_Warehouse_LA_Adjust'
  Declare @Cnt Int = 0
  Declare @RowCnt Int = 0
  Declare @Msg Varchar(4000) =''
  Declare @Errcode int = -1
  Declare @subject Varchar(100)

  Declare @Kind Varchar(1), @KindName Varchar(10), @sd_class Varchar(1), @sd_slip_fg Varchar(1), @OD_Name Varchar(20), @OD_CName Varchar(20), @oper Varchar(1)
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
       
  --���ͽվ����A2014/10/14 �L�Ҫ�ܨC��T�w�H 10 �����ͽվ���
  set @Last_Date = '10'
  set @sp_date = Substring(Convert(Varchar(10), getdate(), 111), 1, 8)+@Last_Date
  set @sd_date = @sp_date
  Set @sd_class = '5' -- �վ�.�ռ���
  Set @sd_slip_fg = 'A' -- �վ��
  
  Declare @Rm Varchar(100), @sp_maker Varchar(20), @str Varchar(200)
  Declare @Pos Int

  Set @Rm = '�t�ζפJ�վ��('+Convert(Varchar(10), getdate(), 111)+')'
  Set @sp_maker = 'Admin'

  -- �P�O��ڬO�_�s�b
  select @sp_cnt = sum(cnt)
    from (select count(1) as cnt
            from SYNC_TA13.dbo.sslpdt m
           where sd_class = @sd_class
             and sd_slip_fg = @sd_slip_fg
             and sd_date = @sp_date
             and sd_lotno like '%'+@Rm+'%'
           union
          select count(1) as cnt
            from SYNC_TA13.dbo.sslpdttmp m
           where sd_class = @sd_class
             and sd_slip_fg = @sd_slip_fg
             and sd_date = @sp_date
             and sd_lotno like '%'+@Rm+'%') m

  if @sp_cnt > 0
  begin
     set @Cnt = @Errcode
     set @strSQL = ''
     Set @subject = 'Exec '+@Proc+' LA�ܲ��ͽվ��ڥ���!!'
     set @Msg ='�w�� ['+@sp_Date+'] �վ���(�w/���T�{)��ơA�Y�ݭn����h����M���T�{�Υ��T�{��ګ᭫��Y�i�Csslpdt, @sd_class=['+@sd_class+'], @sd_slip_fg=['+@sd_slip_fg+'], @sd_date=['+@sp_Date+']�C'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
     Execute uSP_Sys_Send_Mail @Proc, @Subject, 'fa@ta-yeh.com.tw', @Msg, ''

     Return(@Errcode)
  end
  else
  begin
     --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
     -- ���o�̷s�渹
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
     set @New_sdno = Convert(Varchar(10), Convert(Int, @New_sdno) +1)
     set @Msg = '���o�̷s�վ�渹(�����q�w�֭�H�Υ��֭㪺������),new_sdno=['+@new_sdno+'], sd_class=['+@sd_class+'], sd_slip_fg=['+@sd_slip_fg+'], sd_date=['+@sd_date+'].'
     Exec uSP_Sys_Write_Log @Proc, @Msg, '', @Cnt
     --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
     set @Msg = '�s�W['+@sd_date+'] ��վ���ӡC'
     set @strSQL ='insert into TYEIPDBS2.lytdbta13.dbo.sslpdttmp '+@CR+
                  '(sd_class, sd_slip_fg, sd_date, sd_no, '+@CR+
                  ' sd_skno, sd_name, sd_whno, sd_whno2, sd_qty, '+@CR+
                  ' sd_price, sd_dis, sd_stot, sd_lotno, sd_unit, '+@CR+
                  ' sd_unit_fg, sd_ave_p, sd_rate, sd_seqfld, sd_ordno '+@CR+
                  ') '+@CR+
                  'select '''+@sd_class+''' as sd_class, '+@CR+ -- ���O
                  '       '''+@sd_slip_fg+''' sd_slip_fg, '+@CR+ -- ��ں���
                  '       convert(datetime, '''+@sd_date+''') as sd_date, '+@CR+ -- �f����
                  '       '''+@New_sdno+''' as sd_no, '+@CR+ -- �f��s��
                  ''+@CR+
                  '       m.wd_skno as sd_skno, '+@CR+ -- �f�~�s��
                  '       m.sk_name as sd_name, '+@CR+ -- �~�W�W��
                  '       ''LA'' as sd_whno, '+@CR+ -- �ܮw(�J)
                  '       '''' as sd_whno2, '+@CR+ -- �ܮw(�X)
                  '       m.wd_last_qty * -1 as sd_qty, '+@CR+ -- ����ƶq
                  ''+@CR+
                  '       m.sk_save as sd_price, '+@CR+ -- ����
                  '       0 as sd_dis, '+@CR+ -- ���
                  '       m.wd_last_qty * -1 * m.sk_save as sd_stot, '+@CR+ -- �p�p
                  '       '''+@Rm+''' as sd_lotno, '+@CR+ -- �ƥ����
                  '       m.sk_unit as sd_unit, '+@CR+ -- ���
                  ' '+@CR+
                  '       0 as sd_unit_fg, '+@CR+ -- ���X��
                  '       m.sk_save as sd_ave_p, '+@CR+ -- ��즨��
                  '       1 as sd_rate, '+@CR+ -- �ײv
                  '       rowid as sd_seqfld, '+@CR+ -- ���ӧǸ�, ���Ǹ��|�]����V�ק�ӥ[�H�ܧ�A�ҥH�t�s�@���� sd_ordno
                  '       rowid as sd_ordno'+@CR+ -- XLS ���ӧǸ�
                  '  from (select m.wd_skno, d.sk_name, sk_unit, sk_save, '+@CR+
                  '               sum(isnull(wd_amt0+wd_amt1+wd_amt2+wd_amt3+wd_amt4+wd_amt5+wd_amt6+ '+@CR+
                  '                          wd_amt7+wd_amt8+wd_amt9+wd_amt10+wd_amt11+wd_amt12, 0)) as wd_last_qty, '+@CR+
                  '               row_number() over(order by m.wd_skno) as rowid  '+@CR+
                  '          from SYNC_TA13.dbo.swaredt m '+@CR+
                  '               left join SYNC_TA13.dbo.sstock d '+@CR+
                  '                 on m.wd_skno = d.sk_no '+@CR+
                  -- 2014/10/15 Rickliu �쥻�ϥΨ���~�ת��w�s��ơA���o�{��V���w�s��ƬO�����ʤ~�|�i��L�b�A
                  -- �]���Y�h�~�ת��ӫ~�Y�S���ʫh�w�s�Ƥ����d�b�h�~�A�ҥH��ħ���C�Ӱӫ~�̫�@�~���w�s��ơC
                  '               inner join '+@CR+
                  '               (select wd_no, max(wd_yr) as wd_yr, wd_skno '+@CR+
                  '                  from SYNC_TA13.dbo.swaredt '+@CR+
                  '                 group by wd_no, wd_skno '+@CR+
                  '               ) d1 '+@CR+
                  '                 on m.wd_no = d1.wd_no '+@CR+
                  '                and m.wd_yr = d1.wd_yr '+@CR+
                  '                and m.wd_skno = d1.wd_skno '+@CR+
                  '         where 1=1 '+@CR+
                  '           and m.wd_no = ''LA'' '+@CR+
                  '           and m.wd_class=''0'' '+@CR+
                  '           and m.wd_skno like ''A%'' '+@CR+
                  '         group by m.wd_skno, d.sk_name, sk_unit, sk_save '+@CR+
                  '         having sum(isnull(wd_amt0+wd_amt1+wd_amt2+wd_amt3+wd_amt4+wd_amt5+wd_amt6+ '+@CR+
                  '                           wd_amt7+wd_amt8+wd_amt9+wd_amt10+wd_amt11+wd_amt12, 0)) <> 0 '+@CR+
                  '       ) m '+@CR+
                  ' Order by 1, 2, 3'
     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
     
     --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
     -- �s�W�վ�D�� PKey: sp_no (ASC), sp_slip_fg (ASC)
     set @Msg = '�s�W�վ�D�ɡC'
     set @strSQL ='insert into TYEIPDBS2.lytdbta13.dbo.ssliptmp '+@CR+
                  '(sp_class, sp_slip_fg, sp_date, sp_pdate, sp_no, '+@CR+
                  ' sp_maker, sp_tot, sp_tax, sp_dis, '+@CR+
                  ' sp_pay_kd, sp_rate_nm, sp_rate, sp_itot, sp_inv_kd, '+@CR+
                  ' sp_tax_kd, sp_invtype, sp_rem, sp_tal_rec '+@CR+
                  ') '+@CR+
                  'select distinct sd_class as sp_class, '+@CR+
                  '       sd_slip_fg as sp_slip_fg, '+@CR+
                  '       sd_date as sp_date, '+@CR+
                  '       sd_date as sp_pdate, '+@CR+
                  '       sd_no as sp_no, '+@CR+
                  ' '+@CR+
                  '       '''+@sp_maker+''' as sp_maker, '+@CR+ -- �s��H��
                  '       sum(sd_stot) as sp_tot, '+@CR+ -- �p�p
                  
                  -- 2013/6/10 ����^�f���ݭn�|��
                  '       0 as sp_tax, '+@CR+ -- ��~�|(��)
                  
                  '       0 as sp_dis, '+@CR+ -- �������B(��)
                  ' '+@CR+
                  '       '''' as sp_pay_kd, '+@CR+ -- �������
                  '       ''NT'' as sp_rate_nm, '+@CR+ -- �ײv�W��
                  '       1 as sp_rate, '+@CR+ -- �ײv
                  '       0 as sp_itot, '+@CR+ -- �o�����B
                  '       0 as sp_inv_kd, '+@CR+ -- �o�����O(=1  �T�p��,=2  �G�p��,=3  ���Ⱦ�)
                  ' '+@CR+
                  '       0 as sp_tax_kd, '+@CR+ -- �|�O(=1���|,=2�s�| )
                  '       0 as sp_invtype, '+@CR+ -- �}�ߤ覡(=1���}, =2�H��}��, =3�妸�}��)
                  '       '''+@Rm+''' as sp_rem, '+@CR+ -- �Ƶ�
                  '       Count(1) as sp_tal_rec '+@CR+ -- �����`����
                  '  from TYEIPDBS2.lytdbta13.dbo.sslpdttmp m '+@CR+
                  ' where sd_class = '''+@sd_class+''' '+@CR+
                  '   and sd_slip_fg = '''+@sd_slip_fg+''' '+@CR+
                  '   and sd_date = '''+@sd_date+''' '+@CR+
                  '   and sd_lotno like '''+@RM+'%'' '+@CR+
                  ' group by sd_class, '+@CR+
                  '       sd_slip_fg, '+@CR+
                  '       sd_date, '+@CR+
                  '       sd_no '

     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

     Set @subject = 'Exec '+@Proc+' ���\!!'
     set @Msg ='���\�פJ ['+@sp_Date+'] LA�ܽվ���(���T�{)��ơC'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
     Execute uSP_Sys_Send_Mail @Proc, @Subject, 'fa@ta-yeh.com.tw', @Msg, ''
  end

  Return(0)
end
GO
