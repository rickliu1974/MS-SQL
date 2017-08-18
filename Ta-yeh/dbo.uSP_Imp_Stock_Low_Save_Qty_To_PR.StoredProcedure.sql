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
  
  set @rDB = @wDB -- Rickliu 2017/03/24 在訂閱服務還沒好時，暫時使用 @wDB 資料庫的資料
  
  Declare @RM Varchar(100) = '低安存量自動轉請購單'
  Declare @RowCount Table (cnt int)
  Declare @Sender Varchar(100) = 'pu@ta-yeh.com.tw;vp888@ta-yeh.com.tw'

  Declare @strSQL Varchar(Max)

  Declare @TB_tmp_Name Varchar(100) = 'Stock_Low_Save_Qty_To_PR_tmp'
  
  Declare @bd_no Varchar(20) = ''
  Declare @bd_date1 DateTime
  Declare @cnt_ctno Int, @cnt_skno Int, @sum_stot Int, @cnt_PR Int, @cnt_PR_DT Int
  
  
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_tmp_Name+']') AND type in (N'U'))
  begin
     Set @Msg = '清除臨時轉入介面資料表 ['+@TB_tmp_Name+']。'

     Set @strSQL = 'DROP TABLE [dbo].['+@TB_tmp_Name+']'
     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end
  
  set @Msg = '建立臨時轉入介面資料表 ['+@TB_tmp_Name+']。'
  set @strSQL = 'select ''### 主檔資料 ###'' as ''### 主檔資料 ###'', '+@CR+
                -- 請購單號
                '       br_no = Substring(max_no, 1, 6)+Substring(Convert(Varchar(5), Convert(Int, Substring(max_no, 7, 4))+ '+@CR+
                -- 依群組重新分號
                -- '               DENSE_RANK() over(order by m.s_supp)+10000), 2, 4), '+@CR+
                -- 2015/02/06 Rickliu 美如表示將所有廠商之採購歸入一張請購單，所以固定值為 1
                '               1+10000), 2, 4), '+@CR+                -- 請購日期
                '       br_date1 = convert(varchar(10), getdate(), 111), '+@CR+
                -- 請購人員
                --'       br_sales = ct_sales, '+@CR+
				-- 2015/05/11 Nanliao 美如表示將所有採購人員改為Admin
                '       br_sales = ''Admin'', '+@CR+
                -- 部門編號
				-- 2015/06/17 Rickliu 將統一改為 A5120
				-- 2016/07/22 NanLiao 將統一改為 A4210
                '       br_dpno = ''A4210'', '+@CR+
                -- 製單人員
                '       br_maker = ''Admin'', '+@CR+
                -- 合計金額
                '       br_tot =  Round(Sum(case '+@CR+
                '                             when isnull(s_lprice1, 0) = 0 '+@CR+
                '                             then isnull(sk_save, 0) '+@CR+
                '                             else isnull(s_lprice1, 0) '+@CR+
                '                           end * '+@CR+
                -- 2015/02/04 Rickliu 當批量低於安全庫存量 則採安全庫存量
                '                           case '+@CR+
                '                             when isnull(sk_bqty, 0) < (chg_wd_AA_sQty+chg_wd_AB_sQty+chg_wd_AC_sQty) '+@CR+
                '                             then (chg_wd_AA_sQty+chg_wd_AB_sQty+chg_wd_AC_sQty) '+@CR+
                '                           else isnull(sk_bqty, 0) '+@CR+
                -- 依群組統計
                --'                     end * isnull(sk_bqty, 0)) Over(Partition By m.s_supp), '+@CR+
                -- 2015/02/06 Rickliu 美如表示將所有廠商之採購歸入一張請購單
                '                           end) Over(Partition By 1), 0), '+@CR+
                -- 明細筆數
                -- 依群組統計
                --'       or_tal_rec = Count(*) Over(Partition By m.s_supp), '+@CR+
                -- 2015/02/06 Rickliu 美如表示將所有廠商之採購歸入一張請購單
                '       br_tal_rec = Count(1) Over(Partition by 1),  '+@CR+
                -- 單據附註
                '       br_rem = '''+@RM+'''+Convert(Varchar(20), getdate(), 121), '+@CR+
                -- 已交完否 0:未交完
                '       br_ispack = ''0'', '+@CR+
                -- 未交完筆數
                -- 依群組重新分號
                --'       or_npack = Count(*) Over(Partition By m.s_supp), '+@CR+
                -- 2015/02/06 Rickliu 美如表示將所有廠商之採購歸入一張請購單
                '       br_npack = Count(1) Over(Partition by 1), '+@CR+
                -- 已審核否
                '       br_surefg = 0, '+@CR+
                '       ''### 明細資料 ###'' as ''### 明細資料 ###'', '+@CR+
                -- 交貨日期
                '       bd_date2 = convert(varchar(10), getdate()+1, 111), '+@CR+
                -- 廠商編號
                '       bd_ctno = s_supp, '+@CR+
                -- 廠商名稱
                '       bd_ctname= chg_supp_name, '+@CR+
                -- 貨品編號
                '       bd_skno = sk_no, '+@CR+
                -- 貨品名稱
                '       bd_name = sk_name, '+@CR+
                -- 單位
                '       bd_unit = sk_unit, '+@CR+
                -- 數量 -- 2014/02/04 Rickliu 當批量低於安全庫存量 則採安全庫存量
                '       bd_qty = case '+@CR+
                '                  when isnull(sk_bqty, 0) < (chg_wd_AA_sQty+chg_wd_AB_sQty+chg_wd_AC_sQty) '+@CR+
                '                  then (chg_wd_AA_sQty+chg_wd_AB_sQty+chg_wd_AC_sQty) '+@CR+
                '                  else isnull(sk_bqty, 0) '+@CR+
                '                end, '+@CR+
                -- 單價 (若無最近一次進價則抓取平均成本)
                '       bd_price = case '+@CR+
                '                    when isnull(s_lprice1, 0) = 0 '+@CR+
                '                    then isnull(sk_save, 0) '+@CR+
                '                    else isnull(s_lprice1, 0) '+@CR+
                '                  end, '+@CR+
                -- 小計
                '       bd_stot = Round(case '+@CR+
                '                         when isnull(s_lprice1, 0) = 0 '+@CR+
                '                         then isnull(sk_save, 0) '+@CR+
                '                         else isnull(s_lprice1, 0) '+@CR+
                '                       end * '+@CR+
                -- 2015/02/04 Rickliu 當批量低於安全庫存量 則採安全庫存量
                '                       case '+@CR+
                '                         when isnull(sk_bqty, 0) < (chg_wd_AA_sQty+chg_wd_AB_sQty+chg_wd_AC_sQty) '+@CR+
                '                         then (chg_wd_AA_sQty+chg_wd_AB_sQty+chg_wd_AC_sQty) '+@CR+
                '                         else isnull(sk_bqty, 0) '+@CR+
                '                       end, 0), '+@CR+
                -- 單位類別(使用基本數量)
                '       bd_unit_fg = 0, '+@CR+
                -- 貨品附註
                '       bd_rem = '''+@RM+'''+Convert(Varchar(20), getdate(), 121), '+@CR+
                -- 匯率名稱
                '       bd_rate_nm = d.ct_curt_id, '+@CR+
                -- 匯率
                '       bd_rate = d.Chg_Rate, '+@CR+
                -- 交貨旗標
                '       bd_is_pack = 0, '+@CR+
                -- 已審核否
                '       bd_surefg = 0, '+@CR+
                -- 明細編號
                -- 依群組重新分號
                --'       od_seqfld = row_number() over(PARTITION BY m.s_supp order by sk_no), '+@CR+
                -- 2015/02/06 Rickliu 美如表示將所有廠商之採購歸入一張請購單
                '       bd_seqfld = row_number() over(order by sk_no), '+@CR+
                '       ''### 數量檢查 ###'' as ''### 數量檢查 ###'', '+@CR+
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
                -- 2015/02/04 增加 已採未進 的判別(採購單)
                '       left join '+@CR+
                '       (select od_skno, Sum(Isnull(od_qty, 0)) as od_qty '+@CR+
                '          from '+@rDB+'sorddt '+@CR+
                '         where od_class = ''1'' '+@CR+
                '           and (od_is_pack = ''0'') '+@CR+
                '         group by od_skno '+@CR+
                '       ) d2 '+@CR+
                '        on m.sk_no = d2.od_skno '+@CR+
                -- 2015/02/04 增加 已請未採 的判別(請購單)
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
                '                  from '+@rDB+'sborder '+@CR+ -- 請購單(已確認)
                '                 where br_date1 >= convert(varchar(10), getdate(), 111) '+@CR+
                '                 union '+@CR+
                '                select br_no '+@CR+ 
                '                  from '+@rDB+'sbordertmp '+@CR+ -- 請購單(未確認)
                '                 where br_date1 >= convert(varchar(10), getdate(), 111) '+@CR+
                '               ) m '+@CR+
                '       ) d1 '+@CR+
                ' where 1=1 '+@CR+
                '   and (chg_wd_AA_sQty+chg_wd_AB_sQty+chg_wd_AC_sQty) > '+@CR+ -- 低安存量
                '       (chg_wd_AA_last_diff_qty+chg_wd_AB_last_diff_qty+chg_wd_AC_last_diff_qty + '+@CR+ -- 目前庫存量
                '        isnull(od_qty, 0) + isnull(bd_qty, 0)) '+@CR+ -- 請購量, 採購量
                '   and sk_color in (''無'', '''') '+@CR+
                '   and (chg_wd_AA_sQty+chg_wd_AB_sQty+chg_wd_AC_sQty + isnull(od_qty, 0) + isnull(bd_qty, 0)) <> 0 '+@CR+

                -- 2015/04/02 美如要求不限定供應商
                /*
                '   and chg_supp_name not like ''%大業%'' '+@CR+
                '   and chg_supp_name not like ''%安伯特%'' '+@CR+
                '   and chg_supp_name not like ''%寶麟化工%'' '+@CR+
                */
                ' order by m.s_supp, sk_no '
                     
  Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  
  set @Cnt = 0
  Set @Msg = '檢查['+@TB_tmp_Name+']待轉檔資料是否存在。'
  set @strSQL = 'select count(distinct br_no) from '+@TB_tmp_Name
  delete @RowCount
  print @strSQL
  insert into @RowCount Exec (@strSQL)
  select @Cnt=cnt from @RowCount
  
  if @Cnt <> 0
  begin
    set @Msg = '新增請購單已確認單 主檔資料'
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
                
    set @Msg = '新增請購單已確認單 明細資料'
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
    

    set @Msg = '查詢筆數'
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
       Set @RM = @RM+'...'+convert(varchar(20), getdate(), 120)+' 已自動產生 【'+Convert(Varchar(100), @cnt_PR)+'】 張請購單, 【'+Convert(Varchar(10), @Cnt_PR_DT)+'】筆明細!!'
       Set @Msg = '請購單號：【'+@bd_no+'】'+@CR+
                  '請購日期：【'+Convert(Varchar(10), @bd_date1, 111)+'】'+@CR+
                  '廠商家數：【'+Convert(Varchar(10), @Cnt_ctno)+'】'+@CR+
                  '商品項數：【'+Convert(Varchar(10), @Cnt_skno)+'】'+@CR+
                  '請購總額：【'+Convert(Varchar(10), @sum_stot)+'】!!'+@CR+
                  ' '+@CR+
                  '請於凌越系統請購已確認單內查詢!!'
       Exec uSP_Sys_Send_Mail @Proc, @RM, @Sender, @Msg, ''
    end
  end
  else
  begin
     Set @Msg = convert(varchar(20), getdate(), 120) +' 無低安資料，無須轉請購單!!'
     Exec uSP_Sys_Send_Mail @Proc, @Msg, @Sender, @Msg, ''
  end
end
GO
