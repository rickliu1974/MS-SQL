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
     八百屋 後台拉單一律轉訂單，訂單日期一律抓取製表日期，而 指定到貨日 則寫入 訂單的交貨日期
     切記：
     1.由於 3C 與 百貨 的客戶編號不相同，所以拉單時必須分開處理。
     2.此程序僅處理客戶揀貨所產生的揀貨單，透過此單轉成我司之訂單，除此以外則不在此處理範圍內。
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
  Declare @or_wkno varchar(20) -- 單據上的採購單號當作客戶訂單單號 
  Declare @Or_Cnt int

  Declare @txt_head Varchar(100), @TB_head Varchar(200), @TB_txt_Name Varchar(200), @TB_OR_Name Varchar(200), @TB_tmp_name Varchar(200)

  Declare @strSQL Varchar(Max)
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @RowCount Table (cnt int)
  Declare @RowData Table (aData Varchar(255))
  
  Declare @New_Orno varchar(10)
  Declare @Last_Date Varchar(2)
  Declare @Sct_SubQuery Varchar(1000) -- 促銷策略子查詢

  Set @CompanyName = '八百屋' --> 請勿亂變動
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
           (select 'Order' as OR_Name, '揀貨轉訂單' as OR_CName, '3' as OR_Class
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
    Set @TB_txt_Name = @txt_head+@TB_Head+@KindName -- Ex: Ori_txt#Mall_Normal_Order_800yaoya_3C -- Text 原檔資料
    Set @TB_tmp_Name = @TB_Head+@KindName+'_'+'tmp' -- Ex: Mall_Normal_Order_800yaoya_3C_tmp -- 臨時轉入介面
    Set @TB_OR_Name = @TB_Head+@KindName -- Ex: Mall_Normal_Order_800yaoya_3C
    Set @Rm = '系統匯入'+@CompanyName+@KindName

    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    -- Check 匯入檔案是否存在
    IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_txt_Name+']') AND type in (N'U'))
    begin
print '4'
       --Send Message
       set @Cnt = @Errcode
       set @strSQL = ''
       Set @Msg = '外部 Text 匯入資料表 ['+@TB_txt_Name+']不存在，終止進行轉檔作業。'
       Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

       close Cur_Mall_Normal_Order_800yaoya_DataKind
       deallocate Cur_Mall_Normal_Order_800yaoya_DataKind
       Return(@Errcode)
    end
    
    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    -- 判別 Text 列印日期 是否存在
print '5'
    set @Cnt = 0
    set @strSQL = 'Select Distinct SUBSTRING(Convert(Text, F1), 11, 10) as Or_Date1 from [dbo].['+@TB_txt_Name+'] where F1 like ''%製表日期%'' '
    set @Msg = '判別 Text 製表日期 是否存在'
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

    --由於拉單日期不固定，所以一律改以表單的製表日期為主。
    --目前僅支援同一天的製表日期只能轉一次，避免因人員改單造成再次轉檔覆蓋
    delete @RowData
    insert into @RowData exec (@strSQL)
    select @Or_Date1=Rtrim(isnull(aData, '')) from @RowData
    set @Or_Date1 = Convert(Varchar(10), CONVERT(Date, @Or_Date1) , 111)
    
    if @Or_Date1 = ''
    begin
print '6'
       set @Cnt = @Errcode
       set @strSQL = ''
       Set @Msg = '找不到 Text 資料內的製表日期，終止進行轉檔作業。'
       Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

       close Cur_Mall_Normal_Order_800yaoya_DataKind
       deallocate Cur_Mall_Normal_Order_800yaoya_DataKind
       Return(@Errcode)
    end

    --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
    -- 增加判別 確認 及 未確認 單據 若是存在則不得進行重轉作業。
    -- 2013/10/28 增加判斷若當月24日已由未確認轉確認單則不再重轉，避免凌越未確認單據不斷刪除新增造成單號累增情況
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
       set @Msg ='[ 已確認 或 未確認 單 ] 已有 ['+@Or_Date1+' '+@OR_CName+'] 資料，若需要重轉則先行清除確認及未確認單據後重轉即可。sorder, @or_class=['+@OR_class+'], @SP_Date1=['+@Or_Date1+'], @or_lotno=['+@RM+']。'
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
          Set @Msg = '清除 '+@CompanyName+' 臨時轉入介面資料表 ['+@TB_tmp_Name+']。'

          Set @strSQL = 'DROP TABLE [dbo].['+@TB_tmp_Name+']'
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       end

       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
       --將每筆資料加入唯一鍵值(ps.此資料表請勿排序，保留原始 Text 樣貌，以利日後產出對帳表)
       --建立暫存檔且建立序號，建立序號步驟很重要，因為後續要依序產生店名
print '11'
       Set @Msg = '新增備用欄位 F2 ~ F11 ['+@TB_tmp_Name+']。'
       Set @strSQL = 'select rowid, '+@CR+ -- 列號
                     '       F1, '+@CR+ --原資料
                     '       Convert(Varchar(100), '''') as F2,  '+@CR+  --店名
                     '       Convert(Varchar(100), '''') as F3,  '+@CR+  --序      
                     '       Convert(Varchar(100), '''') as F4,  '+@CR+  --商品編號
                     '       Convert(Varchar(100), '''') as F5,  '+@CR+  --商品名稱
                     '       Convert(Varchar(100), '''') as F6,  '+@CR+  --訂貨量  
                     '       Convert(Varchar(100), '''') as F7,  '+@CR+  --贈品否  
                     '       Convert(Varchar(100), '''') as F8,  '+@CR+  --揀貨量  
                     '       Convert(Varchar(100), '''') as F9,  '+@CR+  --進價    
                     '       Convert(Varchar(100), '''') as F10,  '+@CR+ --小計    
                     '       Convert(Varchar(100), '''') as F11,  '+@CR+ --備註    
                     '       Convert(Varchar(100), '''') as F12,  '+@CR+ --採購單號    
                     '       Convert(Varchar(100), '''') as F13,  '+@CR+ --指定到貨日    
                     '       print_date='''+@Or_Date1+''', '+@CR+
                     /*
                     '       xlsFileName, '+@CR+ -- 原始 XLS 檔案名稱
                     '       imp_date, '+@CR+ -- 執行 SP_Imp_xls_to_db 匯入日期
                     */
                     '       SP_Exec_date=getdate() '+@CR+ -- 執行本程序日期
                     '       into ['+@TB_tmp_Name+']'+@CR+
                     '  from ['+@TB_txt_Name+']'
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '12'
       Set @Msg = '抽取店別資料 ['+@TB_tmp_Name+']。'
       Set @strSQL = 'update ['+@TB_tmp_Name+']'+@CR+
                     '   set F2=Replace(Replace(REPLACE(F1, '' '', ''''), ''店)廠商揀貨用(不送門市)'', ''''), ''('', '''') '+@CR+
                     ' where F1 like ''%'+SPACE(10)+'(%'''
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       
       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '13'
       -- 此處會將每一筆的店名給寫入 F2 欄位
       Set @Msg = '處理臨時轉入介面資料表之群組小計店名資料 ['+@TB_tmp_Name+']。'
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
                     -- 由於每一家有自己的採購單號與指定到貨日，所以額外處理
             --'  print @f1  '+@CR+
                     '  set @f12 = substring(Convert(Text, @f1), 52, 8) '+@CR+
             --'  print @f12  '+@CR+
                     '  if @f12 = ''採購單號'' '+@CR+
                     '     set @vf12 = Substring(Convert(Text, @f1), 64, 10) '+@CR+
             --'  print @vf12  '+@CR+
                     ''+@CR+
                     '  set @f13 = substring(Convert(Text, @f1), 52, 10) '+@CR+
             --'  print @f13  '+@CR+
                     '  if @f13 = ''指定到貨日'' '+@CR+
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
       Set @Msg = '處理各個欄位資料 ['+@TB_tmp_Name+']。'
       Set @strSQL = 'update ['+@TB_tmp_Name+']'+@CR+
                     '   set  F3=Replace(LTrim(RTrim(Substring(Convert(Text, F1),   1,  3))), '','', ''''), '+@CR+ --序
                     '        F4=Replace(LTrim(RTrim(Substring(Convert(Text, F1),   5, 14))), '','', ''''), '+@CR+ --商品編號
                     '        F5=Replace(LTrim(RTrim(Substring(Convert(Text, F1),  20, 40))), '','', ''''), '+@CR+ --商品名稱
                     '        F6=Replace(LTrim(RTrim(Substring(Convert(Text, F1),  61,  4))), '','', ''''), '+@CR+ --訂貨量
                     '        F7=Replace(LTrim(RTrim(Substring(Convert(Text, F1),  66,  4))), '','', ''''), '+@CR+ --贈品否  
                     '        F8=Replace(LTrim(RTrim(Substring(Convert(Text, F1),  72,  6))), '','', ''''), '+@CR+ --揀貨量
                     '        F9=Replace(LTrim(RTrim(Substring(Convert(Text, F1),  78, 10))), '','', ''''), '+@CR+ --進價
                     '       F10=Replace(LTrim(RTrim(Substring(Convert(Text, F1),  89, 12))), '','', ''''), '+@CR+ --小計
                     '       F11=LTrim(RTrim(Substring(Convert(Text, F1), 103, 10)))  '  --備註
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '15'
       Set @Msg = '刪除臨時轉入介面資料表之非必要資料 ['+@TB_tmp_Name+']。'
       Set @strSQL = 'delete ['+@TB_tmp_Name+']'+@CR+
                     ' where Isnumeric(Ltrim(Rtrim(substring(F1, 1, 3)))) <> 1'+@CR+
                     '    or Rtrim(Isnull(F4, '''')) = '''' '

       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
--=-[產生單據明細資料，數量為負數的要產生退貨單]=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '16'
       IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_OR_Name+']') AND type in (N'U'))
       begin
print '17'
          Set @Msg = '刪除資料表 ['+@TB_OR_Name+']'
          Set @strSQL = 'DROP TABLE [dbo].['+@TB_OR_Name+']'
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       end

       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
       --2013/04/25 雅婷說逐筆轉入就好不用加總，比較好對
       --在此進行比對商品編號，會拿輔助編號以及商品基本資料檔進行比對作業
print '18'
       Set @Msg = '進行比對商品編號並寫入 ['+@TB_OR_Name+']資料表。'
       set @strSQL = 'select rowid, ct_no, ct_sname, ct_ssname, '+@CR+
                     '       isnull(d.co_skno, d1.sk_no) as sk_no, '+@CR+
                     '       isnull(d.sk_name, d1.sk_name) as sk_name, '+@CR+
                     -- 原商品名稱
                     '       f5, '+@CR+ 
                     '       isnull(d.sk_bcode, d1.sk_bcode) as sk_bcode, '+@CR+
                     -- 原商品編號
                     '       f4, '+@CR+ 
                          -- sum(case when f13 < 0 then f13 * -1 else f13 end) as sale_qty, 
                          -- sum(case when f18 < 0 then f18 * -1 else f18 end) as sale_amt
                     --'       case when Convert(Numeric(20, 6), f13) < 0 then Convert(Numeric(20, 6), f13) * -1 else Convert(Numeric(20, 6), f13) end as sale_qty, '+@CR+
                     -- 原商品數量
                     '       Convert(Numeric(20, 6), F6) as Ori_sale_qty, '+@CR+ -- 2015/08/19 Rickliu 保留原本欄位，避免後面修改
                     '       Convert(Numeric(20, 6), F6) as sale_qty, '+@CR+
                     -- 單價
                     '       (Convert(Numeric(20, 6), F10)) / (Convert(Numeric(20, 6), F6)) as Ori_sale_amt, '+@CR+ -- 2015/08/19 Rickliu 保留原本欄位，避免後面修改
                     '       (Convert(Numeric(20, 6), F10)) / (Convert(Numeric(20, 6), F6)) as sale_amt, '+@CR+
                     -- 小計
                     '       (Convert(Numeric(20, 6), F10)) as Ori_sale_tot, '+@CR+ -- 2015/08/19 Rickliu 保留原本欄位，避免後面修改
                     '       (Convert(Numeric(20, 6), F10)) as sale_tot, '+@CR+ 
                     -- 備註
                     '       F11,  '+@CR+ 
                     -- 採購單號
                     '       F12,  '+@CR+ 
                     -- 指定到貨日
                     '       F13,  '+@CR+ 
                     -- 是否對應
                     '       case when isnull(d.co_skno, d1.sk_no) is null then ''N'' else ''Y'' end as isfound, '+@CR+
                     -- 2015/07/23 Rickliu 增加促銷策略欄位，用此寫法是為了避免 Cursor 耗掉大量資源，因此採 subQuery 的作法，雖然不是最好做法。
                     -- 促銷方案代號
                     '       sct_ss_csno = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_csno')+', ''''), '+@CR+ 
                     -- 促銷方案名稱
                     '       sct_ss_csname = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_csname')+', ''''), '+@CR+ 
                     -- 明細代碼
                     '       sct_ss_rec = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_rec')+', 0), '+@CR+ 
                     -- 客戶類型及編號
                     '       sct_ss_ctkind = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_ctkind')+', ''''), '+@CR+ 
                     -- 客戶編號
                     '       sct_ss_ctno = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_ctno')+', ''''), '+@CR+ 
                     -- 貨品類型及編號
                     '       sct_ss_nokind = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_nokind')+', ''''), '+@CR+ 
                     -- 貨品編號
                     '       sct_ss_no = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_no')+', ''''), '+@CR+ 
                     -- 贈品編號
                     '       sct_ss_sendno = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_sendno')+', ''''), '+@CR+ 
                     -- 品項數量 
                     '       sct_ss_noqty = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_noqty')+', 0), '+@CR+ 
                     -- 贈品數量
                     '       sct_ss_sendqty = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_sendqty')+', 0), '+@CR+ 
                     -- 單次贈品限制 
                     '       sct_ss_oneqty = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_oneqty')+', 0), '+@CR+ 
                     -- 客戶贈品限制
                     '       sct_ss_itmqty = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_itmqty')+', 0), '+@CR+ 
                     -- 總贈品限制
                     '       sct_sendtot = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_sendtot')+', 0), '+@CR+ 
                     -- 有效起始日期
                     '       sct_ss_sdate = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_sdate')+', getdate()), '+@CR+ 
                     -- 有效截止日期
                     '       sct_ss_edate = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_edate')+', getdate()), '+@CR+ 
                     -- 起始數量
                     '       sct_ss_sqty = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_sqty')+', 0), '+@CR+ 
                     -- 中止數量
                     '       sct_ss_eqty = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_eqty')+', 0), '+@CR+ 
                     -- 價格公式
                     '       sct_ss_form = Isnull('+Replace(@Sct_SubQuery, '@@', 'ss_form')+', ''''), '+@CR+ 
                     -- 是否為贈品 0: 是, 1: 否
                     '       1 as isSend '+@CR+
                     '       into ['+@TB_OR_Name+'] '+@CR+
                     '  from (select * '+@CR+
                     '          from (select ct_no, rtrim(ct_sname) as ct_sname, rtrim(ct_name)+''#''+rtrim(ct_sname) as ct_ssname '+@CR+
                     '                  from SYNC_TA13.dbo.PCUST  '+@CR+
                     '                 where ct_class =''1''  '+@CR+
                     '                   and ct_name like '''+@CompanyLikeName+''' '+@CR+
                     '                   and substring(ct_no, 9, 1) ='''+@Kind+''' '+@CR+ -- -- 1:百貨, 2:3C, 3:寄倉, 4:代工OEM
                     '                ) m, ['+@TB_tmp_Name+'] d '+@CR+
                     '         where m.ct_ssname collate Chinese_Taiwan_Stroke_CI_AS like ''%''+rtrim(d.f2)+''%'' collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                     --'           and Convert(Numeric(20, 6), f6) '+@oper+' 0 '+@CR+ -- 判斷數量是否為負數代表退貨
                     '       )m '+@CR+
                     '       left join '+@CR+ 
                     '         (select distinct co_ctno, co_skno, co_cono, sk_no, sk_name, sk_bcode '+@CR+
                     '            from SYNC_TA13.dbo.sauxf m '+@CR+ -- (進銷存)客戶輔助編號
                     '                 left join SYNC_TA13.dbo.sstock d '+@CR+
                     '                   on m.co_skno = d.sk_no '+@CR+
                     '           where co_class=''1'' '+@CR+
                     '         ) d '+@CR+
                     '          on 1=1 '+@CR+
                     '         and m.ct_no=d.co_ctno '+@CR+
                     -- F4 原商品編號
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

       -- 2015/08/19 Rickliu 不支援數量級距公式，因 SS_Mult_Text 之資料內文如右：@@19.000000;1:104.76@@1000.000000;1:61.9@@0.000000;:@@
       --                    @@19.000000;1:104.76 ==>    0 ~   19 價格公式 = 104.76
       --                    @@1000.000000;1:61.9 ==>   20 ~ 1000 價格公式 = 61.9
       --                    @@0.000000;:@@       ==> 1001 ~      價格公式 = ""
       -- 2015/08/19 Rickliu 增加支援 單一數量區間價格
print '19'
       Set @Msg = '進行比對商品編號並寫入 ['+@TB_OR_Name+']資料表。'
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
       -- 2015/08/19 Rickliu 增加支援 單一贈品
       Set @Msg = '進行比對商品編號並寫入 ['+@TB_OR_Name+']資料表。'
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

       Set @Msg = '進行更新重複的rowid ['+@TB_OR_Name+']資料表。'
       set @strSQL = 'update ['+@TB_OR_Name+'] '+@CR+
                     '   set rowid = rowid * 10  '+@CR+
                     ' where sk_no=sct_ss_sendno '+@CR+
                     '   and sk_no > ''''        '+@CR+
	                 ' and sct_ss_no > ''''      ' 
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL            
       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
       --產生出貨退回單
print '21'
       set @Cnt = 0
       Set @Msg = '檢查['+@TB_OR_Name+']待轉檔資料是否存在。'
       set @strSQL = 'select count(1) from ['+@TB_OR_Name+']'
       delete @RowCount
       insert into @RowCount Exec (@strSQL)

       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
       -- 因為使用 Memory Variable Table, 所以使用 >0 判斷
       select @Cnt=cnt from @RowCount
print @Cnt
print '22'
       if @Cnt >0
       begin
       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '23'
          set @Msg = @Msg + '...存在，將進行轉入程序。'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', @Cnt
          --取得最新出貨退回單號(必須從已核准以及未核准的表單取號)
          --請勿再加+1，因為新增時會自動加入
          --2013/11/24 改為 @od_date (每月固定使用 25日做為轉檔日)
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
          set @Msg = '取得最新訂單單號(必須從已核准以及未核准的表單取號),new_orno=['+@New_Orno+'], or_class=['+@Or_Class+'], or_date=['+@Or_Date1+'].'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', @Cnt
                         
          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '24'
          set @Msg = '刪除 ['+@Or_Date1+'] 日轉入的訂單單據明細。'
          --2014/04/22 副總指示直接轉確認訂單，無須再行審核。

          set @strSQL ='delete TYEIPDBS2.lytdbta13.dbo.sorddt '+@CR+
                       ' where Convert(Varchar(10), od_date1, 111) = '''+@Or_Date1+''' '+@CR+
                       '   and od_class = '''+@Or_Class+''' '+@CR+
                       '   and od_lotno like ''%'+@Rm+'%'' '
           
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '25'
          set @Msg = '刪除 ['+@Or_Date1+'] 日轉入的訂單單據主檔。'
          --2014/04/22 副總指示直接轉確認訂單，無須再行審核。
          set @strSQL ='delete TYEIPDBS2.lytdbta13.dbo.sorder '+@CR+
                       ' where Convert(Varchar(10), or_date1, 111) = '''+@Or_Date1+''' '+@CR+
                       '   and or_class = '''+@Or_Class+''' '+@CR+
                       '   and or_rem like ''%'+@Rm+'%'' '
 
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '26'
          set @Msg = '新增 ['+@Or_Date1+'] 訂單主檔。'
          --2014/04/22 副總指示直接轉確認訂單，無須再行審核。
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
                       '       d1.ct_no as or_ctno, '+@CR+ -- 客戶編號
                       
                       '       d1.ct_name as or_ctname, '+@CR+ -- 客戶名稱
                       '       d1.ct_sales as or_sales, '+@CR+ -- 業務員
                       '       d1.ct_dept as or_dpno, '+@CR+ -- 部門編號
                       '       '''+@Or_Maker+''' as or_maker, '+@CR+ -- 製單人員
                       '       sum(m.sale_tot) as or_tot, '+@CR+ -- 小計
                       
                       '       sum(m.sale_tot)*0.05 as or_tax, '+@CR+ -- 營業稅(原)
                       '       ''NT'' as or_rate_nm, '+@CR+ -- 匯率名稱
                       '       1 as or_rate, '+@CR+ -- 匯率
                       '       Substring(F12+'':''+Convert(Varchar(10), rowid), 1, 10) as or_wkno, '+@CR+ -- 客戶訂單單號
                       '       '''+@Rm+'於 ''+Convert(Varchar(20), Getdate(), 120)+char(13)+char(10)+'''+@OR_CName+''' as or_rem, '+@CR+ -- 備註
                       '       ''0'' as od_is_pack  '+@CR+ -- 交貨值
                       '  from ['+@TB_OR_Name+'] m '+@CR+
                       '       left join SYNC_TA13.dbo.sstock d '+@CR+
                       '         on m.sk_no = d.sk_no '+@CR+
                       '       left join SYNC_TA13.dbo.pcust d1 '+@CR+
                       '         on m.ct_no = d1.ct_no '+@CR+
                       '        and ct_class =''1'' '+@CR+
                       '       left join  '+@CR+
                       '        (select ct_no,  '+@CR+
                       '                convert(varchar(10), '+@New_Orno+'+row_number() over(order by m.ct_no)) as od_no '+@CR+ -- 貨單編號
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
                       '       d1.ct_no, '+@CR+ -- 客戶編號
                       '       d1.ct_name, '+@CR+ -- 客戶名稱
                       '       substring(d1.ct_addr3, 1, 255), '+@CR+ -- 送貨地址
                       '       d1.ct_sales, '+@CR+ -- 業務員
                       '       d1.ct_dept, '+@CR+ -- 部門編號
                       '       d1.ct_porter, '+@CR+ -- 貨運公司
                       '       Substring(F12+'':''+Convert(Varchar(10), rowid), 1, 10), '+@CR+ -- 客戶訂單單號
                       '       d1.ct_payfg '
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL


          set @Msg = '新增 ['+@Or_Date1+'] 日訂單明細。'
          --2014/04/22 副總指示直接轉確認訂單，無須再行審核。
          set @strSQL ='insert into TYEIPDBS2.lytdbta13.dbo.sorddt '+@CR+
                       '(od_class, or_no, od_date1, od_date2, od_ctno, '+@CR+
                       ' od_skno, od_name, od_price, od_unit, od_qty, '+@CR+
                       ' od_unit_fg, od_dis, od_stot, od_rate, od_ave_p, '+@CR+
                       ' od_lotno, od_seqfld, od_spec, od_sendfg, '+@CR+
                       ' od_csno, od_csrec '+@CR+
                       ' ,od_is_pack '+@CR+-- 已交完否
                       ') '+@CR+
                       'select '''+@Or_Class+''' as od_class, '+@CR+ -- 類別
                       '       d2.OR_NO, '+@CR+ -- 貨單編號
                       '       d2.OR_DATE1 as od_date1, '+@CR+ -- 貨單日期
					   -- 20160205 modify by Nan 佳瑄提出，可將交貨日期與貨單日期相同，若此欄位空白，此單將無法修改交貨狀態。
                       '       d2.OR_DATE2 as od_date2, '+@CR+ -- 交貨日期
                       --'       --convert(datetime, f13) as od_date2, '+@CR+ -- 交貨/有效日期
                       '       d2.OR_CTNO as od_ctno, '+@CR+ -- 客戶編號

                       '       d.sk_no as od_skno, '+@CR+ -- 貨品編號
                       '       d.sk_name as od_name, '+@CR+ -- 品名規格
                       '       m.sale_amt as od_price, '+@CR+ -- 價格
                       '       d.sk_unit as od_unit, '+@CR+ -- 單位
                       '       m.sale_qty as od_qty, '+@CR+ -- 交易數量
                       
                       '       0 as od_unit_fg, '+@CR+ -- 單位旗標
                       '       1 as od_dis, '+@CR+ -- 折數
                       '       m.sale_tot as od_stot, '+@CR+ -- 小計
                       '       1 as od_rate, '+@CR+ -- 匯率
                       '       d.s_price4 as od_ave_p, '+@CR+ -- 單位成本

                       '       '''+@Rm+@OR_CName +''' as od_lotno, '+@CR+ -- 備用欄位
                       '       rowid as od_seqfld, '+@CR+ -- 明細序號
                       '       F12+'':''+Convert(Varchar(10), rowid) as od_spec, '+@CR+ -- 明細序號, 此序號會因為凌越修改而加以變更，所以另存一份到 od_spec
                       '       Case '+@CR+
                       '         when m.sale_amt = 0 '+@CR+ --2013/12/28 客服芝惠表示，若金額為零則代表為贈品
                       '         then 1 '+@CR+
                       '         else 0 '+@CR+
                       '       end as od_sendfg, '+@CR+ -- 是否為贈品
                       '       sct_ss_csno as od_csno,  '+@CR+ -- 促銷方案代號
                       '       sct_ss_rec as od_csrec,  '+@CR+ -- 促銷方案明細代碼
                       '       ''0'' as od_is_pack  '+@CR+ -- 交貨值
                       '  from ['+@TB_OR_Name+'] m '+@CR+
                       '       left join SYNC_TA13.dbo.sstock d '+@CR+
                       '         on m.sk_no = d.sk_no '+@CR+
                       '       left join SYNC_TA13.dbo.pcust d1 '+@CR+
                       '         on m.ct_no = d1.ct_no '+@CR+
                       '        and ct_class =''1'' '+@CR+					   
                       '       left join TYEIPDBS2.LYTDBTA13.dbo.sorder d2 '+@CR+ -- 2017/03/28 Rickliu 因為訂閱資料沒法這麼迅速即時，因此只能回頭去抓取凌越資料庫
                       '         on m.ct_no = d2.OR_CTNO '+@CR+	
                       --'       left join  '+@CR+
                       --'        (select ct_no,  '+@CR+
                       ---- 2015/11/19 Rickliu, 由於同一天同一家分店可能轉入多張訂單情況，因此增加 F12 原訂單的採購單號 作為分單號碼區別
                       ----'                convert(varchar(10), '+@New_Orno+'+row_number() over(order by m.ct_no)) as od_no '+@CR+ -- 貨單編號
                       --'                convert(varchar(10), '+@New_Orno+'+row_number() over(order by m.ct_no, Ori_Buyno)) as od_no, '+@CR+ -- 貨單編號
                       --'                Ori_Buyno '+@CR+ -- 原始採單編號
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





        --  set @Msg = '新增 ['+@Or_Date1+'] 日訂單明細。'
        --  --2014/04/22 副總指示直接轉確認訂單，無須再行審核。
        --  set @strSQL ='insert into TYEIPDBS2.lytdbta13.dbo.sorddt '+@CR+
        --               '(od_class, or_no, od_date1, od_date2, od_ctno, '+@CR+
        --               ' od_skno, od_name, od_price, od_unit, od_qty, '+@CR+
        --               ' od_unit_fg, od_dis, od_stot, od_rate, od_ave_p, '+@CR+
        --               ' od_lotno, od_seqfld, od_spec, od_sendfg, '+@CR+
        --               ' od_csno, od_csrec '+@CR+
        --               ') '+@CR+
        --               'select '''+@Or_Class+''' as od_class, '+@CR+ -- 類別
        --               '       d2.od_no, '+@CR+ -- 貨單編號
        --               '       convert(datetime, '''+@Or_Date1+''') as od_date1, '+@CR+ -- 貨單日期
					   ---- 20160205 modify by Nan 佳瑄提出，可將交貨日期與貨單日期相同，若此欄位空白，此單將無法修改交貨狀態。
        --               '       convert(datetime, '''+@Or_Date1+''') as od_date2, '+@CR+ -- 交貨日期
        --               --'       --convert(datetime, f13) as od_date2, '+@CR+ -- 交貨/有效日期
        --               '       m.ct_no as od_ctno, '+@CR+ -- 客戶編號

        --               '       d.sk_no as od_skno, '+@CR+ -- 貨品編號
        --               '       d.sk_name as od_name, '+@CR+ -- 品名規格
        --               '       m.sale_amt as od_price, '+@CR+ -- 價格
        --               '       d.sk_unit as od_unit, '+@CR+ -- 單位
        --               '       m.sale_qty as od_qty, '+@CR+ -- 交易數量
                       
        --               '       0 as od_unit_fg, '+@CR+ -- 單位旗標
        --               '       1 as od_dis, '+@CR+ -- 折數
        --               '       m.sale_tot as od_stot, '+@CR+ -- 小計
        --               '       1 as od_rate, '+@CR+ -- 匯率
        --               '       d.s_price4 as od_ave_p, '+@CR+ -- 單位成本

        --               '       '''+@Rm+@OR_CName +''' as od_lotno, '+@CR+ -- 備用欄位
        --               '       rowid as od_seqfld, '+@CR+ -- 明細序號
        --               '       F12+'':''+Convert(Varchar(10), rowid) as od_spec, '+@CR+ -- 明細序號, 此序號會因為凌越修改而加以變更，所以另存一份到 od_spec
        --               '       Case '+@CR+
        --               '         when m.sale_amt = 0 '+@CR+ --2013/12/28 客服芝惠表示，若金額為零則代表為贈品
        --               '         then 1 '+@CR+
        --               '         else 0 '+@CR+
        --               '       end as od_sendfg, '+@CR+ -- 是否為贈品
        --               '       sct_ss_csno as od_csno,  '+@CR+ -- 促銷方案代號
        --               '       sct_ss_rec as od_csrec  '+@CR+ -- 促銷方案明細代碼
        --               '  from ['+@TB_OR_Name+'] m '+@CR+
        --               '       left join TYEIPDBS2.lytdbta13.dbo.sstock d '+@CR+
        --               '         on m.sk_no = d.sk_no '+@CR+
        --               '       left join TYEIPDBS2.lytdbta13.dbo.pcust d1 '+@CR+
        --               '         on m.ct_no = d1.ct_no '+@CR+
        --               '        and ct_class =''1'' '+@CR+
        --               '       left join  '+@CR+
        --               '        (select ct_no,  '+@CR+
        --               -- 2015/11/19 Rickliu, 由於同一天同一家分店可能轉入多張訂單情況，因此增加 F12 原訂單的採購單號 作為分單號碼區別
        --               --'                convert(varchar(10), '+@New_Orno+'+row_number() over(order by m.ct_no)) as od_no '+@CR+ -- 貨單編號
        --               '                convert(varchar(10), '+@New_Orno+'+row_number() over(order by m.ct_no, Ori_Buyno)) as od_no, '+@CR+ -- 貨單編號
        --               '                Ori_Buyno '+@CR+ -- 原始採單編號
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
          set @Msg = '回寫商品比對重複備註。'
          --2014/04/22 Rickliu 副總指示直接轉確認訂單，無須再行審核。
          --2017/06/21 Rickliu 不知為何當以 Join Update 的方式回寫 TYEIPDBS2，會大量產生 Lock，之後便產生以下錯誤訊息，因此改以 Cursor 方式 Updated 回去
          -- 錯誤訊息:無法從連結伺服器 "TYEIPDBS2" 的 OLE DB 提供者 "SQLNCLI10" 取得資料列的資料。
          --set @strSQL ='update TYEIPDBS2.lytdbta13.dbo.sorddt '+@CR+
          --             '   set od_rem='''+@Rm+@OR_CName+'商品比對重複 RowID:''+D.Rowid '+@CR+
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
                       '     set od_rem='''+@Rm+@OR_CName+'商品比對重複 RowID:''+@Rowid '+@CR+
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
          -- 新增出貨退回主檔 PKey: sp_no (ASC), sp_slip_fg (ASC)
print '28'
          --set @Msg = '新增 ['+@Or_Date1+'] 訂單主檔。'
          ----2014/04/22 副總指示直接轉確認訂單，無須再行審核。
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
          --             '       d1.ct_no as or_ctno, '+@CR+ -- 客戶編號
                       
          --             '       d1.ct_name as or_ctname, '+@CR+ -- 客戶名稱
          --             '       d1.ct_sales as or_sales, '+@CR+ -- 業務員
          --             '       d1.ct_dept as or_dpno, '+@CR+ -- 部門編號
          --             '       '''+@Or_Maker+''' as or_maker, '+@CR+ -- 製單人員
          --             '       sum(od_stot) as or_tot, '+@CR+ -- 小計
                       
          --             '       sum(od_stot)*0.05 as or_tax, '+@CR+ -- 營業稅(原)
          --             '       ''NT'' as or_rate_nm, '+@CR+ -- 匯率名稱
          --             '       1 as or_rate, '+@CR+ -- 匯率
          --             '       Substring(od_spec, 1, 10) as or_wkno, '+@CR+ -- 客戶訂單單號
          --             '       '''+@Rm+'於 ''+Convert(Varchar(20), Getdate(), 120)+char(13)+char(10)+'''+@OR_CName+''' as or_rem, '+@CR+ -- 備註
          --             '       ''0'' '+@CR+ -- 已交完否
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
          --             '       d1.ct_no, '+@CR+ -- 客戶編號
          --             '       d1.ct_name, '+@CR+ -- 客戶名稱
          --             '       substring(d1.ct_addr3, 1, 255), '+@CR+ -- 送貨地址
          --             '       d1.ct_sales, '+@CR+ -- 業務員
          --             '       d1.ct_dept, '+@CR+ -- 部門編號
          --             '       d1.ct_porter, '+@CR+ -- 貨運公司
          --             '       Substring(od_spec, 1, 10), '+@CR+ -- 客戶訂單單號
          --             '       d1.ct_payfg '
          --Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
          -- 2015/11/25 Rickliu 增加刪除無頭的明細資料，避免客服人員人工打訂單轉銷貨時，產生因轉檔失敗的資料出現。
          -- 2017/07/17 Rickliu 自從資料同步改為訂閱方式後，不知何故，跨DB進行 Update 會常態性失敗並且產生以下錯誤訊息，實無找不到問題所在，只好改寫為 Cursor 方式
          -- 錯誤訊息:無法從連結伺服器 "TYEIPDBS2" 的 OLE DB 提供者 "SQLNCLI10" 取得資料列的資料。
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

          set @Msg = '刪除 ['+@Or_Date1+'] 無頭之訂單明細檔。'
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
          set @Msg = @Msg + '...不存在，終止轉入程序。'
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
