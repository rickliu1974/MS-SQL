USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_Stock_Low_Save_Qty_To_PR]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_Imp_Stock_Low_Save_Qty_To_PR]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_Imp_Stock_Low_Save_Qty_To_PR]
as
begin
  Declare @Proc Varchar(50) = 'uSP_Imp_Stock_Low_Save_Qty_To_PR'
  Declare @Cnt Int = 0
  Declare @RowCnt Int = 0
  Declare @Msg Varchar(4000) =''
  Declare @Errcode int = -1
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @rDB Varchar(100) = 'SYNC_TA13.dbo.'
  Declare @wDB Varchar(100) = 'TYEIPDBS2.lytdbta13.dbo.'
  
  set @rDB = @wDB -- Rickliu 2017/03/24 �b�q�\�A���٨S�n�ɡA�Ȯɨϥ� @wDB ��Ʈw�����
  
  Declare @RM Varchar(100) = '�C�w�s�q�۰�����ʳ�'
  Declare @RowCount Table (cnt int)
  Declare @Sender Varchar(100) = 'pu@ta-yeh.com.tw;vp888@ta-yeh.com.tw'

  Declare @strSQL Varchar(Max)

  Declare @TB_tmp_Name Varchar(100) = 'Stock_Low_Save_Qty_To_PR_tmp'
  
  Declare @bd_no Varchar(20) = ''
  Declare @bd_date1 DateTime
  Declare @cnt_ctno Int, @cnt_skno Int, @sum_stot Int, @cnt_PR Int, @cnt_PR_DT Int
  
  
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_tmp_Name+']') AND type in (N'U'))
  begin
     Set @Msg = '�M���{����J������ƪ� ['+@TB_tmp_Name+']�C'

     Set @strSQL = 'DROP TABLE [dbo].['+@TB_tmp_Name+']'
     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end
  
  set @Msg = '�إ��{����J������ƪ� ['+@TB_tmp_Name+']�C'
  set @strSQL = 'select ''### �D�ɸ�� ###'' as ''### �D�ɸ�� ###'', '+@CR+
                -- ���ʳ渹
                '       br_no = Substring(max_no, 1, 6)+Substring(Convert(Varchar(5), Convert(Int, Substring(max_no, 7, 4))+ '+@CR+
                -- �̸s�խ��s����
                -- '               DENSE_RANK() over(order by m.s_supp)+10000), 2, 4), '+@CR+
                -- 2015/02/06 Rickliu ���p��ܱN�Ҧ��t�Ӥ������k�J�@�i���ʳ�A�ҥH�T�w�Ȭ� 1
                '               1+10000), 2, 4), '+@CR+                -- ���ʤ��
                '       br_date1 = convert(varchar(10), getdate(), 111), '+@CR+
                -- ���ʤH��
                --'       br_sales = ct_sales, '+@CR+
				-- 2015/05/11 Nanliao ���p��ܱN�Ҧ����ʤH���אּAdmin
                '       br_sales = ''Admin'', '+@CR+
                -- �����s��
				-- 2015/06/17 Rickliu �N�Τ@�אּ A5120
				-- 2016/07/22 NanLiao �N�Τ@�אּ A4210
                '       br_dpno = ''A4210'', '+@CR+
                -- �s��H��
                '       br_maker = ''Admin'', '+@CR+
                -- �X�p���B
                '       br_tot =  Round(Sum(case '+@CR+
                '                             when isnull(s_lprice1, 0) = 0 '+@CR+
                '                             then isnull(sk_save, 0) '+@CR+
                '                             else isnull(s_lprice1, 0) '+@CR+
                '                           end * '+@CR+
                -- 2015/02/04 Rickliu ���q�C��w���w�s�q �h�Ħw���w�s�q
                '                           case '+@CR+
                '                             when isnull(sk_bqty, 0) < (chg_wd_AA_sQty+chg_wd_AB_sQty+chg_wd_AC_sQty) '+@CR+
                '                             then (chg_wd_AA_sQty+chg_wd_AB_sQty+chg_wd_AC_sQty) '+@CR+
                '                           else isnull(sk_bqty, 0) '+@CR+
                -- �̸s�ղέp
                --'                     end * isnull(sk_bqty, 0)) Over(Partition By m.s_supp), '+@CR+
                -- 2015/02/06 Rickliu ���p��ܱN�Ҧ��t�Ӥ������k�J�@�i���ʳ�
                '                           end) Over(Partition By 1), 0), '+@CR+
                -- ���ӵ���
                -- �̸s�ղέp
                --'       or_tal_rec = Count(*) Over(Partition By m.s_supp), '+@CR+
                -- 2015/02/06 Rickliu ���p��ܱN�Ҧ��t�Ӥ������k�J�@�i���ʳ�
                '       br_tal_rec = Count(1) Over(Partition by 1),  '+@CR+
                -- ��ڪ���
                '       br_rem = '''+@RM+'''+Convert(Varchar(20), getdate(), 121), '+@CR+
                -- �w�槹�_ 0:���槹
                '       br_ispack = ''0'', '+@CR+
                -- ���槹����
                -- �̸s�խ��s����
                --'       or_npack = Count(*) Over(Partition By m.s_supp), '+@CR+
                -- 2015/02/06 Rickliu ���p��ܱN�Ҧ��t�Ӥ������k�J�@�i���ʳ�
                '       br_npack = Count(1) Over(Partition by 1), '+@CR+
                -- �w�f�֧_
                '       br_surefg = 0, '+@CR+
                '       ''### ���Ӹ�� ###'' as ''### ���Ӹ�� ###'', '+@CR+
                -- ��f���
                '       bd_date2 = convert(varchar(10), getdate()+1, 111), '+@CR+
                -- �t�ӽs��
                '       bd_ctno = s_supp, '+@CR+
                -- �t�ӦW��
                '       bd_ctname= chg_supp_name, '+@CR+
                -- �f�~�s��
                '       bd_skno = sk_no, '+@CR+
                -- �f�~�W��
                '       bd_name = sk_name, '+@CR+
                -- ���
                '       bd_unit = sk_unit, '+@CR+
                -- �ƶq -- 2014/02/04 Rickliu ���q�C��w���w�s�q �h�Ħw���w�s�q
                '       bd_qty = case '+@CR+
                '                  when isnull(sk_bqty, 0) < (chg_wd_AA_sQty+chg_wd_AB_sQty+chg_wd_AC_sQty) '+@CR+
                '                  then (chg_wd_AA_sQty+chg_wd_AB_sQty+chg_wd_AC_sQty) '+@CR+
                '                  else isnull(sk_bqty, 0) '+@CR+
                '                end, '+@CR+
                -- ��� (�Y�L�̪�@���i���h�����������)
                '       bd_price = case '+@CR+
                '                    when isnull(s_lprice1, 0) = 0 '+@CR+
                '                    then isnull(sk_save, 0) '+@CR+
                '                    else isnull(s_lprice1, 0) '+@CR+
                '                  end, '+@CR+
                -- �p�p
                '       bd_stot = Round(case '+@CR+
                '                         when isnull(s_lprice1, 0) = 0 '+@CR+
                '                         then isnull(sk_save, 0) '+@CR+
                '                         else isnull(s_lprice1, 0) '+@CR+
                '                       end * '+@CR+
                -- 2015/02/04 Rickliu ���q�C��w���w�s�q �h�Ħw���w�s�q
                '                       case '+@CR+
                '                         when isnull(sk_bqty, 0) < (chg_wd_AA_sQty+chg_wd_AB_sQty+chg_wd_AC_sQty) '+@CR+
                '                         then (chg_wd_AA_sQty+chg_wd_AB_sQty+chg_wd_AC_sQty) '+@CR+
                '                         else isnull(sk_bqty, 0) '+@CR+
                '                       end, 0), '+@CR+
                -- ������O(�ϥΰ򥻼ƶq)
                '       bd_unit_fg = 0, '+@CR+
                -- �f�~����
                '       bd_rem = '''+@RM+'''+Convert(Varchar(20), getdate(), 121), '+@CR+
                -- �ײv�W��
                '       bd_rate_nm = d.ct_curt_id, '+@CR+
                -- �ײv
                '       bd_rate = d.Chg_Rate, '+@CR+
                -- ��f�X��
                '       bd_is_pack = 0, '+@CR+
                -- �w�f�֧_
                '       bd_surefg = 0, '+@CR+
                -- ���ӽs��
                -- �̸s�խ��s����
                --'       od_seqfld = row_number() over(PARTITION BY m.s_supp order by sk_no), '+@CR+
                -- 2015/02/06 Rickliu ���p��ܱN�Ҧ��t�Ӥ������k�J�@�i���ʳ�
                '       bd_seqfld = row_number() over(order by sk_no), '+@CR+
                '       ''### �ƶq�ˬd ###'' as ''### �ƶq�ˬd ###'', '+@CR+
                '       sk_bqty as chk_sk_bqty, '+@CR+
                '       (chg_wd_AA_sQty+chg_wd_AB_sQty+chg_wd_AC_sQty) as chk_wd_save_qty, '+@CR+
                '       (chg_wd_AA_last_diff_qty+chg_wd_AB_last_diff_qty+chg_wd_AC_last_diff_qty) as chk_wd_last_diff_qty, '+@CR+
                '       isnull(d2.od_qty, 0) as chk_od_qty, '+@CR+
                '       isnull(d3.bd_qty, 0) as chk_bd_qty '+@CR+
                '       into '+@TB_tmp_Name+' '+@CR+
                '  from fact_sstock m '+@CR+
                '       left join fact_pcust d '+@CR+
                '         on m.s_supp = d.ct_no '+@CR+
                '        and d.ct_class = ''2'' '+@CR+
                -- 2015/02/04 �W�[ �w�ĥ��i ���P�O(���ʳ�)
                '       left join '+@CR+
                '       (select od_skno, Sum(Isnull(od_qty, 0)) as od_qty '+@CR+
                '          from '+@rDB+'sorddt '+@CR+
                '         where od_class = ''1'' '+@CR+
                '           and (od_is_pack = ''0'') '+@CR+
                '         group by od_skno '+@CR+
                '       ) d2 '+@CR+
                '        on m.sk_no = d2.od_skno '+@CR+
                -- 2015/02/04 �W�[ �w�Х��� ���P�O(���ʳ�)
                '       left join '+@CR+
                '       (select bd_skno, sum(isnull(bd_qty, 0)) as bd_qty '+@CR+
                '          from (select bd_skno, bd_qty '+@CR+
                '                  from '+@rDB+'sborddt '+@CR+
                '                 where bd_is_pack = ''0'' '+@CR+
                '                 union '+@CR+
                '                select bd_skno, bd_qty '+@CR+
                '                  from '+@rDB+'sborddttmp '+@CR+
                '                 where bd_is_pack = ''0'' '+@CR+
                '               ) m '+@CR+
                '          group by bd_skno '+@CR+
                '       ) d3 '+@CR+
                '         on m.sk_no = d3.bd_skno '+@CR+
                '       cross join '+@CR+ 
                '       (select isnull(max(br_no), Substring(convert(varchar(10), getdate(), 112), 3, 6)+''0000'') as max_no '+@CR+
                '          from (select br_no '+@CR+
                '                  from '+@rDB+'sborder '+@CR+ -- ���ʳ�(�w�T�{)
                '                 where br_date1 >= convert(varchar(10), getdate(), 111) '+@CR+
                '                 union '+@CR+
                '                select br_no '+@CR+ 
                '                  from '+@rDB+'sbordertmp '+@CR+ -- ���ʳ�(���T�{)
                '                 where br_date1 >= convert(varchar(10), getdate(), 111) '+@CR+
                '               ) m '+@CR+
                '       ) d1 '+@CR+
                ' where 1=1 '+@CR+
                '   and (chg_wd_AA_sQty+chg_wd_AB_sQty+chg_wd_AC_sQty) > '+@CR+ -- �C�w�s�q
                '       (chg_wd_AA_last_diff_qty+chg_wd_AB_last_diff_qty+chg_wd_AC_last_diff_qty + '+@CR+ -- �ثe�w�s�q
                '        isnull(od_qty, 0) + isnull(bd_qty, 0)) '+@CR+ -- ���ʶq, ���ʶq
                '   and sk_color in (''�L'', '''') '+@CR+
                '   and (chg_wd_AA_sQty+chg_wd_AB_sQty+chg_wd_AC_sQty + isnull(od_qty, 0) + isnull(bd_qty, 0)) <> 0 '+@CR+

                -- 2015/04/02 ���p�n�D�����w������
                /*
                '   and chg_supp_name not like ''%�j�~%'' '+@CR+
                '   and chg_supp_name not like ''%�w�B�S%'' '+@CR+
                '   and chg_supp_name not like ''%�_��Ƥu%'' '+@CR+
                */
                ' order by m.s_supp, sk_no '
                     
  Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  
  set @Cnt = 0
  Set @Msg = '�ˬd['+@TB_tmp_Name+']�����ɸ�ƬO�_�s�b�C'
  set @strSQL = 'select count(distinct br_no) from '+@TB_tmp_Name
  delete @RowCount
  print @strSQL
  insert into @RowCount Exec (@strSQL)
  select @Cnt=cnt from @RowCount
  
  if @Cnt <> 0
  begin
    set @Msg = '�s�W���ʳ�w�T�{�� �D�ɸ��'
    set @strSQL = 'Insert Into '+@wDB+'sborder '+@CR+
                  '(br_no, br_date1, br_sales, br_dpno, br_maker, '+@CR+
                  ' br_tot, br_tal_rec, br_rem, br_ispack, br_npack, '+@CR+
                  ' br_surefg) '+@CR+
                  'select distinct '+@CR+
                  '       br_no, br_date1, br_sales, br_dpno, br_maker, '+@CR+
                  '       br_tot, br_tal_rec, br_rem, br_ispack, br_npack, '+@CR+
                  '       br_surefg '+@CR+
                  '  from '+@TB_tmp_Name
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
                
    set @Msg = '�s�W���ʳ�w�T�{�� ���Ӹ��'
    set @strSQL = 'Insert Into '+@wDB+'sborddt '+@CR+
                  '(bd_no, bd_date1, bd_date2, bd_ctno, bd_ctname, '+@CR+
                  ' bd_skno, bd_name, bd_unit, bd_qty, bd_price, '+@CR+
                  ' bd_stot, bd_unit_fg, bd_rem, bd_rate_nm, '+@CR+
                  ' bd_rate, bd_is_pack, bd_surefg, bd_seqfld) '+@CR+
                  'select br_no, br_date1, bd_date2, bd_ctno, bd_ctname, '+@CR+
                  '       bd_skno, bd_name, bd_unit, bd_qty, bd_price, '+@CR+
                  '       bd_stot, bd_unit_fg, bd_rem, bd_rate_nm, '+@CR+
                  '       bd_rate, bd_is_pack, bd_surefg, bd_seqfld '+@CR+
                  '  from '+@TB_tmp_Name
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    

    set @Msg = '�d�ߵ���'
    select @bd_no = RTrim(bd_no), 
           @bd_date1 = bd_date1, 
           @cnt_ctno = count(distinct bd_ctno),
           @cnt_skno = count(distinct bd_skno),
           @sum_stot = sum(bd_stot),
           @cnt_PR = count(distinct bd_no),
           @cnt_PR_DT = count(1)
      from TYEIPDBS2.lytdbta13.dbo.sborddt
     where bd_rem like '%'+@RM+'%'
       and bd_date1 = convert(varchar(10), getdate(), 111)
     group by bd_no, bd_date1
    
    if isnull(@bd_no, '') <> ''
    begin
       Set @RM = @RM+'...'+convert(varchar(20), getdate(), 120)+' �w�۰ʲ��� �i'+Convert(Varchar(100), @cnt_PR)+'�j �i���ʳ�, �i'+Convert(Varchar(10), @Cnt_PR_DT)+'�j������!!'
       Set @Msg = '���ʳ渹�G�i'+@bd_no+'�j'+@CR+
                  '���ʤ���G�i'+Convert(Varchar(10), @bd_date1, 111)+'�j'+@CR+
                  '�t�Ӯa�ơG�i'+Convert(Varchar(10), @Cnt_ctno)+'�j'+@CR+
                  '�ӫ~���ơG�i'+Convert(Varchar(10), @Cnt_skno)+'�j'+@CR+
                  '�����`�B�G�i'+Convert(Varchar(10), @sum_stot)+'�j!!'+@CR+
                  ' '+@CR+
                  '�Щ��V�t�ν��ʤw�T�{�椺�d��!!'
       Exec uSP_Sys_Send_Mail @Proc, @RM, @Sender, @Msg, ''
    end
  end
  else
  begin
     Set @Msg = convert(varchar(20), getdate(), 120) +' �L�C�w��ơA�L������ʳ�!!'
     Exec uSP_Sys_Send_Mail @Proc, @Msg, @Sender, @Msg, ''
  end
end
GO
