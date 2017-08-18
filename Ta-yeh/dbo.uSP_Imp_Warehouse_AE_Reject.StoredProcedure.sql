USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_Warehouse_AE_Reject]    Script Date: 07/24/2017 14:44:00 ******/
DROP PROCEDURE [dbo].[uSP_Imp_Warehouse_AE_Reject]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_Imp_Warehouse_AE_Reject]
as
begin
  /***********************************************************************************************************
    2014/10/14 Rickliu 
    ���p�n�D�G�C��T�w 10 �鲣�Ͷi�f�h�X��ڦ� TA13�A�Ӷi�f�h�X��ڶȰw�� AE �ܡA�ñN�Ҧ��~�����������Ӳ��X���ӡA
  ***********************************************************************************************************/
  Declare @Proc Varchar(50) = 'uSP_Imp_Warehouse_AE_Reject'
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
       
  --���Ͷi�f�h�X����A2014/10/14 ���p��ܨC��T�w�H 10 �����ͽվ���
  set @Last_Date = '10'
  set @sp_date = Substring(Convert(Varchar(10), getdate(), 111), 1, 8)+@Last_Date
  set @sd_date = @sp_date
  Set @sd_class = '0' -- �i�h�f��
  Set @sd_slip_fg = '1' -- �i�h��
  
  Declare @Rm Varchar(100), @sp_maker Varchar(20), @str Varchar(200)
  Declare @Pos Int

  Set @Rm = '�t�ζפJ�i�f�h�X��('+Convert(Varchar(10), getdate(), 111)+')'
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
     Set @subject = 'Exec '+@Proc+' AE�ܲ��ͽվ��ڥ���!!'
     set @Msg ='�w�� ['+@sp_Date+'] �i�f�h�X���(�w/���T�{)��ơA�Y�ݭn����h����M���T�{�Υ��T�{��ګ᭫��Y�i�Csslpdt, @sd_class=['+@sd_class+'], @sd_slip_fg=['+@sd_slip_fg+'], @sd_date=['+@sp_Date+']�C'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
     Execute uSP_Sys_Send_Mail @Proc, @Subject, 'pu@ta-yeh.com.tw', @Msg, ''

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
     --set @New_sdno = Convert(Varchar(10), Convert(Int, @New_sdno) +1)
     set @New_sdno = Substring(@New_sdno, 7, 4)
     set @Msg = '���o�̷s�i�f�h�X�渹(�����q�w�֭�H�Υ��֭㪺������),new_sdno=['+@new_sdno+'], sd_class=['+@sd_class+'], sd_slip_fg=['+@sd_slip_fg+'], sd_date=['+@sd_date+'].'
     Exec uSP_Sys_Write_Log @Proc, @Msg, '', @Cnt
     --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
     set @Msg = '�s�W['+@sd_date+'] ��վ���ӡC'
     set @strSQL ='insert into TYEIPDBS2.lytdbta13.dbo.sslpdttmp '+@CR+
                  '(sd_class, sd_slip_fg, sd_date, sd_no, sd_ctno, '+@CR+
                  ' sd_skno, sd_name, sd_whno, sd_whno2, sd_qty, '+@CR+
                  ' sd_price, sd_dis, sd_stot, sd_lotno, sd_unit, '+@CR+
                  ' sd_unit_fg, sd_ave_p, sd_rate, sd_seqfld, sd_ordno '+@CR+
                  ') '+@CR+
                  'select '''+@sd_class+''' as sd_class, '+@CR+ -- ���O
                  '       '''+@sd_slip_fg+''' sd_slip_fg, '+@CR+ -- ��ں���
                  '       convert(datetime, '''+@sd_date+''') as sd_date, '+@CR+ -- �f����
                  '       substring(convert(Varchar(10), convert(datetime, '''+@sd_date+'''), 112), 3, 6) + '+@CR+
                  '       Substring(Convert(Varchar(10), DENSE_RANK() over(order by m.s_upp)+10000+Convert(Int, '''+@New_sdno+''')), 2, 10) as sd_no, '+@CR+ -- �f��s��
                  ''+@CR+
                  '       m.s_upp as sd_ctno, '+@CR+ -- �����ӽs��
                  '       m.wd_skno as sd_skno, '+@CR+ -- �f�~�s��
                  '       m.sk_name as sd_name, '+@CR+ -- �~�W�W��
                  '       ''AE'' as sd_whno, '+@CR+ -- �ܮw(�J)
                  '       '''' as sd_whno2, '+@CR+ -- �ܮw(�X)
                  '       m.wd_last_qty as sd_qty, '+@CR+ -- ����ƶq
                  ''+@CR+
                  '       m.sk_save as sd_price, '+@CR+ -- ����
                  '       0 as sd_dis, '+@CR+ -- ���
                  '       m.wd_last_qty * m.sk_save as sd_stot, '+@CR+ -- �p�p
                  '       '''+@Rm+''' as sd_lotno, '+@CR+ -- �ƥ����
                  '       m.sk_unit as sd_unit, '+@CR+ -- ���
                  ' '+@CR+
                  '       0 as sd_unit_fg, '+@CR+ -- ���X��
                  '       m.sk_save as sd_ave_p, '+@CR+ -- ��즨��
                  '       1 as sd_rate, '+@CR+ -- �ײv
                  '       row_number() over(partition by m.s_upp order by m.wd_skno) as sd_seqfld, '+@CR+ -- ���ӧǸ�, ���Ǹ��|�]����V�ק�ӥ[�H�ܧ�A�ҥH�t�s�@���� sd_ordno
                  '       row_number() over(partition by m.s_upp order by m.wd_skno) as sd_ordno'+@CR+ -- XLS ���ӧǸ�
                  '  from (select m.wd_skno, d.sk_name, sk_unit, sk_save, '+@CR+
                  '               sum(isnull(wd_amt0+wd_amt1+wd_amt2+wd_amt3+wd_amt4+wd_amt5+wd_amt6+ '+@CR+
                  '                          wd_amt7+wd_amt8+wd_amt9+wd_amt10+wd_amt11+wd_amt12, 0)) as wd_last_qty, '+@CR+
                  '               Case '+@CR+
                  '                 when IsNull(sk_fld11, '''') = '''' '+@CR+
                  '                 then s_supp '+@CR+
                  '                 else sk_fld11 '+@CR+
                  '               end as s_upp '+@CR+
                  '          from SYNC_TA13.dbo.swaredt m '+@CR+
                  '               left join SYNC_TA13.dbo.sstock d '+@CR+
                  '                 on m.wd_skno = d.sk_no '+@CR+
                  -- 2014/10/18 Rickliu �쥻�ϥΨ���~�ת��w�s��ơA���o�{��V���w�s��ƬO�����ʤ~�|�i��L�b�A
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
                  '           and m.wd_no = ''AE'' '+@CR+
                  '           and m.wd_class=''0'' '+@CR+
                  '           and m.wd_skno like ''A%'' '+@CR+
                  '         group by m.wd_skno, d.sk_name, sk_unit, sk_save, '+@CR+
                  '                  Case '+@CR+
                  '                    when IsNull(sk_fld11, '''') = '''' '+@CR+
                  '                    then s_supp '+@CR+
                  '                    else sk_fld11 '+@CR+
                  '                  end '+@CR+
                  '         having sum(isnull(wd_amt0+wd_amt1+wd_amt2+wd_amt3+wd_amt4+wd_amt5+wd_amt6+ '+@CR+
                  '                           wd_amt7+wd_amt8+wd_amt9+wd_amt10+wd_amt11+wd_amt12, 0)) <> 0 '+@CR+
                  '       ) m '+@CR+
                  'where 1=1'+@CR+
                  '  and substring(m.s_upp,1,2) <> ''AZ'''+@CR+    --20150910 add by NanLiao ���p���X�n�ư�AZ�}�Y���t��
                  ' Order by 1, 2, 3, 4'
     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
     
     --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
     -- �s�W�վ�D�� PKey: sp_no (ASC), sp_slip_fg (ASC)
     set @Msg = '�s�W�վ�D�ɡC'
     set @strSQL ='insert into TYEIPDBS2.lytdbta13.dbo.ssliptmp '+@CR+
                  '(sp_class, sp_slip_fg, sp_date, sp_pdate, sp_no, '+@CR+
                  ' sp_ctno, sp_ctname, sp_ctadd2, sp_sales, sp_dpno, '+@CR+
                  ' sp_maker, sp_tot, sp_tax, sp_dis, '+@CR+
                  ' sp_pay_kd, sp_rate_nm, sp_rate, sp_itot, sp_inv_kd, '+@CR+
                  ' sp_tax_kd, sp_invtype, sp_rem, sp_tal_rec, '+@CR+
                  ' sp_ave_p, sp_ntnpay '+@CR+
                  ') '+@CR+
                  'select distinct sd_class as sp_class, '+@CR+
                  '       sd_slip_fg as sp_slip_fg, '+@CR+
                  '       sd_date as sp_date, '+@CR+
                  '       dbo.uFn_Getdate(ct_pmode, sd_date, ct_pdate) as sp_pdate, '+@CR+
                  /*
                  '       case '+@CR+
                  '         when ct_pdate = 0 then sd_date '+@CR+
                  '         when ct_pdate = 31 then  Convert(DateTime, Convert(Varchar(7), DateAdd(mm, 1, sd_date), 111)+''/01'')-1 '+@CR+
                  '         when day(sd_date) >= ct_pdate '+@CR+
                  '         then Convert(DateTime, Convert(Varchar(7), DateAdd(mm, 1, sd_date), 111)+''/''+Convert(Varchar(2), ct_pdate)) '+@CR+
                  '         else Convert(DateTime, Convert(Varchar(7), sd_date, 111)+''/''+Convert(Varchar(2), ct_pdate)) '+@CR+
                  '       end as sp_pdate, '+@CR+
                  */
                  '       sd_no as sp_no, '+@CR+
                  ' '+@CR+
                  '       isnull(sd_ctno, '''') as sp_ctno, '+@CR+ --������
                  '       isnull(ct_name, '''') as sp_ctname, '+@CR+ --�����ӦW��
                  '       convert(Varchar(255), isnull(ct_addr3, '''')) as sp_ctadd2, '+@CR+ --�o���a�}
                  '       isnull(ct_sales, '''') as sp_sales, '+@CR+ --�o���a�}
                  '       isnull(ct_dept, '''') as sp_dpno, '+@CR+ --�����s��
                  ' '+@CR+
                  
                  '       '''+@sp_maker+''' as sp_maker, '+@CR+ -- �s��H��
                  '       Round(sum(sd_stot), 0) as sp_tot, '+@CR+ -- �p�p
                  
                  -- 2013/6/10 ����^�f���ݭn�|��
                  '       Round(sum(sd_stot) * 0.05, 0) as sp_tax, '+@CR+ -- ��~�|(��)
                  
                  '       0 as sp_dis, '+@CR+ -- �������B(��)
                  ' '+@CR+
                  '       '''' as sp_pay_kd, '+@CR+ -- �������
                  '       ''NT'' as sp_rate_nm, '+@CR+ -- �ײv�W��
                  '       1 as sp_rate, '+@CR+ -- �ײv
                  '       Round(sum(sd_stot), 0) as sp_itot, '+@CR+ -- �o�����B
                  '       1 as sp_inv_kd, '+@CR+ -- �o�����O(=1  �T�p��,=2  �G�p��,=3  ���Ⱦ�)
                  ' '+@CR+
                  '       1 as sp_tax_kd, '+@CR+ -- �|�O(=1���|,=2�s�| )
                  '       1 as sp_invtype, '+@CR+ -- �}�ߤ覡(=1���}, =2�H��}��, =3�妸�}��)
                  '       '''+@Rm+''' as sp_rem, '+@CR+ -- �Ƶ�
                  '       Count(1) as sp_tal_rec, '+@CR+ -- �����`����
                  '       sum(sd_stot) as sp_ave_p, '+@CR+ -- ���i����
                  '       Round(sum(sd_stot) * 1.05, 0) as sp_ntnpay '+@CR+ -- ���h���B
                  '  from TYEIPDBS2.lytdbta13.dbo.sslpdttmp m '+@CR+
                  '       left join SYNC_TA13.dbo.pcust d '+@CR+
                  '         on m.sd_ctno = d.ct_no '+@CR+
                  '        and d.ct_class = ''2'' '+@CR+
                  ' where sd_class = '''+@sd_class+''' '+@CR+
                  '   and sd_slip_fg = '''+@sd_slip_fg+''' '+@CR+
                  '   and sd_date = '''+@sd_date+''' '+@CR+
                  '   and sd_lotno like '''+@RM+'%'' '+@CR+
                  '   and substring(sd_ctno,1,2) <> ''AZ'''+@CR+    --20150910 add by NanLiao ���p���X�n�ư�AZ�}�Y���t��
                  ' group by sd_class, sd_slip_fg, sd_date, '+@CR+ 
                  '          dbo.uFn_Getdate(ct_pmode, sd_date, ct_pdate), '+@CR+
                  '          sd_no, sd_ctno, ct_name, convert(Varchar(255), isnull(ct_addr3, '''')), ct_sales, ct_dept '

     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

     Set @subject = 'Exec '+@Proc+' ���\!!'
     set @Msg ='���\�פJ ['+@sp_Date+'] AE�ܽվ���(���T�{)��ơC'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
     Execute uSP_Sys_Send_Mail @Proc, @Subject, 'pu@ta-yeh.com.tw', @Msg, ''
     Return(0)
  end

end
GO
