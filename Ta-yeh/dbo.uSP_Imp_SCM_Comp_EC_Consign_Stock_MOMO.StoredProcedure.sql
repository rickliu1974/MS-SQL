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
   程式修訂歷程：
   1.2016/01/01 Rickliu 新增 EC MOMO 寄倉庫存對帳程式
   2.2017/06/22 MOMO 變更對帳 EXCEL 格式, Rickliu 修訂此程式
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
       
  --產生貨單日期，2013/8/27 林課確定每月固定以24號為轉檔日
  set @Last_Date = '24'
  set @sp_date = Substring(Convert(Varchar(10), getdate(), 111), 1, 8)+@Last_Date
  set @sd_date = @sp_date

  Declare @CompanyName Varchar(100), @CompanyLikeCode Varchar(10), @CompanyLikeName Varchar(100), @Rm Varchar(100), @sp_maker Varchar(20), @str Varchar(200)
  Declare @Pos Int
  
  Set @CompanyLikeCode = 'I90120011'
  Set @CompanyName = 'MOMO'
  Set @CompanyLikeName = '%'+@CompanyName+'%'
  Set @Rm = '系統匯入'+@CompanyName
  Set @sp_maker = 'Admin'

print '4'
  Set @xls_head = 'Ori_Xls#'
  Set @TB_head = 'Comp_EC_Consign_Stock_'+@CompanyName
  Set @TB_xls_Name = Isnull(@xls_head+@TB_Head, '') -- Ex: Ori_Xls#Comp_EC_Consign_Stock__MOMO -- Excel 原檔資料
  Set @TB_tmp_Name = Isnull(@TB_Head, '')+'_tmp' -- Ex: Comp_EC_Consign_Stock_MOMO_tmp -- 臨時轉入介面
  Set @TB_OD_Name = Isnull(@TB_Head, '')
    
  -- Check 匯入檔案是否存在
  print @TB_xls_Name
  IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_xls_Name+']') AND type in (N'U'))
  begin
print '5'
     --Send Message
     Set @Msg = '外部 Excel 匯入資料表 ['+@TB_xls_Name+']不存在，終止進行轉檔作業。'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, @Errcode
     -- 2013/11/28 增加失敗回傳值
     Return(@Errcode)
  end
      
    -- 判別 Excel 列印日期 是否存在
print '6'
  set @Cnt = 0
  set @Print_Date  = ''
    
  set @strSQL = 'Select Top 1 SplitFileName as Print_Date from [dbo].['+@TB_xls_Name+']  '
  set @Msg = '判別 Excel 列印日期 是否存在'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
    
  delete @RowData
  insert into @RowData exec (@strSQL)
  select @Print_Date=Rtrim(isnull(aData, '')) from @RowData
  set @Print_Date = Convert(Varchar(10), CONVERT(Date, @Print_Date) , 111)
    
  --2013/11/25 林課確定每月固定以24號為轉檔日 
  --set @Print_Date = Substring(Convert(Varchar(10), @Print_Date, 111), 1, 8)+@Last_Date
    
  if Rtrim(@Print_Date) = ''
  begin
print '4'
     set @Cnt = @Errcode
     set @strSQL = ''
     Set @Msg = '找不到 Excel 資料內的列印日期，終止進行轉檔作業。'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
     -- 2013/11/28 增加失敗回傳值
     Return(@Errcode)
  end
  else
  begin
print '9'
    IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_tmp_Name+']') AND type in (N'U'))
    begin
print '10'
       Set @Msg = '清除'+@CompanyName+'臨時轉入介面資料表 ['+@TB_tmp_Name+']。'
       Set @strSQL = 'DROP TABLE [dbo].['+@TB_tmp_Name+']'
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    end
    
    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    --將每筆資料加入唯一鍵值(ps.此資料表請勿排序，保留原始 XLS 樣貌，以利日後產出對帳表)
    --建立暫存檔且建立序號，建立序號步驟很重要，因為後續要依序產生店名
print '11'
    Set @Msg = '建立臨時轉入介面資料表每筆唯一序號 ['+@TB_tmp_Name+']。'
    Set @strSQL = 'select *, print_date=Convert(date, rtrim('''+@Print_Date+''')) '+@CR+
                  '       into '+@TB_tmp_Name+@CR+
                  '  from '+@TB_xls_Name
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    
    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '12'
    Set @Msg = '刪除臨時轉入介面資料表之表頭資料 ['+@TB_tmp_Name+']。'
    Set @strSQL = 'delete '+@TB_tmp_Name+@CR+
                  ' where rowid <= 1 '
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

print '17'
    IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_OD_Name+']') AND type in (N'U'))
    begin
print '18'
       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
       Set @Msg = '重建'+@TB_OD_Name+'對帳總表 ['+@TB_head+']。'
    
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

    Set @Msg = '刪除資料表 ['+@TB_OD_Name+']'
    Set @strSQL = 'Delete [dbo].['+@TB_OD_Name+'] '+@CR+
                  ' where Print_Date = '''+@Print_Date+''' '
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

print '19'
    /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
    Excel 原始格式
    20151031_Comp_EC_Consign_Stock_Momo 格式:
    F01:倉別          F02:商品原廠編號(V) F03:品號(V)     F04:品名(V)     F05:規格
    F06:商品狀態      F07:期初庫存        F08:進貨量      F09:客退數      F10:銷貨數
    F11:待退量        F12:期末庫存        F13:即時庫存(V) F14:寄倉在途
    *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/

    Set @Msg = '進行比對商品編號並寫入 ['+@TB_OD_Name+']資料表。'
    set @strSQL = -- CTE_Q1 原始 Excel 資料僅取部分欄位
                 'With CTE_Q1 as ( '+@CR+
                 '  Select F2 as xls_skno, '+@CR+ -- F02:商品條碼(Excel)
                 '         F4 as xls_skname, '+@CR+ -- F04: 商品名稱(Excel)
                 '         F3 as xls_cono, '+@CR+ -- F03: 輔助編號(Excel)
                 '         Sum(isnull(Convert(Numeric(20, 6), Replace(F13, '','', '''')), 0)) as xls_qty '+@CR+ -- F13:寄賣庫存 (Excel)
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
                 -- 只查託售回貨部分
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
                 '         sum(case when sd_slip_fg = ''6'' then sd_qty else 0 end) as fg6_qty, '+@CR+ -- 托售數量
                 '         sum(case when sd_slip_fg = ''7'' then sd_qty else 0 end) as fg7_qty, '+@CR+ -- 回貨數量
                 '         sum(case when sd_slip_fg = ''7'' then sd_qty * -1 else sd_qty end) as sum_qty, '+@CR+ -- 合計數量<-------CHECK
                 '         sum(case when sd_slip_fg = ''6'' then sd_stot else 0 end) as fg6_amt, '+@CR+ -- 托售金額
                 '         sum(case when sd_slip_fg = ''7'' then sd_stot else 0 end) as fg7_amt, '+@CR+ -- 托售數量
                 '         sum(case when sd_slip_fg = ''7'' then sd_stot * -1 else sd_stot end) as sum_amt '+@CR+ -- 合計金額
                 '    from Fact_sslpdt '+@CR+
                 '   where 1=1 '+@CR+
                 '     and sd_slip_fg IN (''6'',''7'')  '+@CR+
                 '     and sd_date >= ''2012/12/01''  '+@CR+
                 -- 範圍截取至 2013/01/01 ~ 列印日期
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
   
                 -- 商品編號
                 '       Convert(Varchar(50), RTrim(isnull(isnull(d.co_skno, d1.sk_no), ''N/A''))) as sk_no, '+@CR+
                 -- 商品名稱
                 '       Convert(Varchar(50), RTrim(isnull(isnull(d.sk_name, d1.sk_name), ''N/A''))) as sk_name, '+@CR+
                 -- 商品條碼
                 '       Convert(Varchar(50), RTrim(isnull(isnull(d.sk_bcode, d1.sk_bcode), ''N/A''))) as sk_bcode, '+@CR+
   
                 -- 商品編號(Excel)
                 '       Convert(Varchar(50), RTrim(Isnull(xls_skno, ''N/A''))) as xls_skno, '+@CR+ 
                 -- 商品輔助編號(Excel)
                 '       Convert(Varchar(50), RTrim(Isnull(xls_cono, ''N/A''))) as xls_cono, '+@CR+ 
                 -- 商品名稱(Excel)
                 '       Convert(Varchar(50), RTrim(isnull(xls_skname, ''N/A''))) as xls_skname, '+@CR+ 
   
                 -- 托售數量(凌越)
                 '       isnull(fg6_qty, 0) as fg6_qty, '+@CR+ 
                 -- 回貨數量(凌越)
                 '       isnull(fg7_qty, 0) as fg7_qty, '+@CR+ 
                 -- 合計數量(凌越)
                 '       isnull(sum_qty, 0 ) as sum_qty, '+@CR+ 
   
                 -- 托售金額(凌越)
                 '       isnull(fg6_amt, 0) as fg6_amt, '+@CR+ 
                 -- 回貨金額(凌越)
                 '       isnull(fg7_amt, 0) as fg7_amt, '+@CR+ 
                 -- 合計金額(凌越)
                 '       isnull(sum_amt, 0) as sum_amt, '+@CR+ 
                 -- 寄賣庫存 (Excel) 
                 '       isnull(Convert(Numeric(20, 6), xls_qty), 0) as xls_qty, '+@CR+
                 -- 差異數量
                 '       isnull(sum_qty - Convert(Numeric(20, 6), xls_qty), 0) as diff_qty, '+@CR+
                 '       '''+@Print_Date+''' as Print_Date, '+@CR+
                 -- 異常
                 '       case '+@CR+
                 '         when (isnull(d.co_skno, d1.sk_no) is null) then ''XLS 編號對不到凌越商品資料'' '+@CR+
                 '         when (fg6_qty is null) then ''XLS 有資料，但凌越無託售回貨資料'' '+@CR+
                 '         when (xls_skno is null and sum_qty <>0) then ''凌越有託售回貨資料，但XLS無資料'' '+@CR+
                 '         when (sum_qty - xls_qty <>0) then ''比對數量異常'' '+@CR+
                 '         else '''' '+@CR+
                 '       end as isfound, '+@CR+
                 '       getdate() as Exec_DateTime '+@CR+
                 -- 客戶編號
                 '  from CTE_Q1 m '+@CR+
                 -- 商品資料
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
                 --  凌越托售回貨資料 sp_slip_fg = 6 托售, sp_slip_fg = 7 回貨
                 '       left join CTE_Q3 d2 '+@CR+
                 '         on 1=1 '+@CR+
                 '        and isnull(d.co_skno, d1.sk_no) = d2.sd_skno '+@CR+
                 ' where 1=1 '+@CR+
                 -- 不顯示之前採購過的商品且數量為 0 的商品, 也就是只顯示剩餘庫存部分
                 --'   and (F3 is not null) '+@CR+
                 --'   and (isnull(sum_qty, 0) <> 0) '+@CR+
                 ' order by 1, 2 '
                 --group by ct_no, ct_sname, ct_ssname, sk_no, sk_name, sk_bcode, f3, f6
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    --2013/12/24將重複的資料也列入異常
print '20'
/*
    set @Msg = '將重複的資料標註異常'
    set @Cnt = 0
    set @strSQL ='update '+@TB_OD_Name+@CR+
                 '   set isfound = isfound+''，重複比對'' '+@CR+
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
  -- 2013/11/28 增加失敗回傳值
  Return(0)
end
GO
