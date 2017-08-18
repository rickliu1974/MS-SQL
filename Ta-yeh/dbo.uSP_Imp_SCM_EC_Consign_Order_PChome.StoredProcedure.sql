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
     PCHome 後台拉單時，請選擇依分店抓取資料，如此 XLS 格式才會正確
     -- 2013/11/28 增加失敗回傳值
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
       
  --產生貨單日期，2013/8/27 林課確定每月固定以25號為轉檔日
  set @Last_Date = '25'
  set @sp_date = Substring(Convert(Varchar(10), getdate(), 111), 1, 8)+@Last_Date
  set @sd_date = @sp_date

  Declare @CompanyName Varchar(100), @CompanyLikeName Varchar(100), @Rm Varchar(100), @sp_maker Varchar(20), @str Varchar(200)
  Declare @Pos Int
  
  Set @CompanyName = '網路家庭國際' --> 請勿亂變動
  Set @CompanyLikeName = '%'+@CompanyName+'%'
  Set @Rm = '系統匯入'+@CompanyName
  Set @sp_maker = 'Admin'

print '1'
  Declare Cur_Mall_Consign_Order_PCHome_DataKind cursor for
    select *
      from (select 'Reject' as OD_Name, '退回轉託售' as OD_CName, '1' as sd_class, '3' as sd_slip_fg, '<' as oper, '退貨貨款明細(商品彙總) - 寄倉訂單' as F1_Tag
             union
            select 'Return' as OD_Name, '託回轉銷售' as OD_CName, '3' as sd_class, '7' as sd_slip_fg, '>' as oper, '訂單貨款明細(商品彙總) - 寄倉訂單' as F1_Tag
           )d
     
  open Cur_Mall_Consign_Order_PCHome_DataKind
  fetch next from Cur_Mall_Consign_Order_PCHome_DataKind into @OD_Name, @OD_CName, @sd_class, @sd_slip_fg, @oper, @F1_Tag

print '2'
  while @@fetch_status =0
  begin
print '3'
    Set @xls_head = 'Ori_Xls#'
    Set @TB_head = 'EC_Consign_Order_PChome'
    Set @TB_xls_Name = @xls_head+@TB_Head -- Ex: Ori_Xls#EC_Consign_Order_PChome -- Excel 原檔資料
    Set @TB_tmp_Name = @TB_Head+'_'+'tmp' -- Ex: EC_Consign_Order_PChome_tmp -- 臨時轉入介面
    Set @TB_OD_Name = @TB_Head+'_'+@OD_Name -- Ex: EC_Consign_Order_PChome_Reject  

    set @strSQL = ''
    set @Cnt = 0
    set @Msg = '進行 '+@F1_Tag+' ==> '+@OD_CName+ ' 轉檔處理。'
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

    -- Check 匯入檔案是否存在
    IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_xls_Name+']') AND type in (N'U'))
    begin
print '4'
       --Send Message
       set @Cnt = @Errcode
       set @strSQL = ''
       Set @Msg = '外部 Excel 匯入資料表 ['+@TB_xls_Name+']不存在，終止進行轉檔作業。'
       Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

       fetch next from Cur_Mall_Consign_Order_PCHome_DataKind into @OD_Name, @OD_CName, @sd_class, @sd_slip_fg, @oper, @F1_Tag
       Continue
    end
    
    -- 判別 Excel 列印日期 是否存在
print '5'
    set @Cnt = 0
    set @Print_Date  = ''

    set @strSQL = 'Select Distinct SUBSTRING(F2, CHARINDEX(''~'', F2)+1, 10) as Print_date from [dbo].['+@TB_xls_Name+'] where F2 like ''列帳日期%'' '
    set @Msg = '判別 Excel 列帳日期 是否存在'
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
       Set @Msg = '找不到 Excel 資料內的列印日期，終止進行轉檔作業。'
       Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
       -- 2013/11/28 增加失敗回傳值
       close Cur_Mall_Consign_Order_PCHome_DataKind
       deallocate Cur_Mall_Consign_Order_PCHome_DataKind
       Return(@Errcode)

--       fetch next from Cur_Mall_Consign_Order_PCHome_DataKind into @OD_Name, @OD_CName, @sd_class, @sd_slip_fg, @oper, @F1_Tag
--       Continue
    end
   
    -- 增加判別 確認 及 未確認 單據 若是存在則不得進行重轉作業。
    -- 2013/10/28 增加判斷若當月24日已由未確認轉確認單則不再重轉，避免凌越未確認單據不斷刪除新增造成單號累增情況
print '7'
    set @Print_Date = Substring(Convert(Varchar(10), @Print_Date, 111), 1, 8)+@Last_Date
    Set @sp_date = @Print_Date
    Set @sd_date = @Print_Date

print '日期:'+@Print_Date
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
       set @Msg ='[ 已確認 或 未確認 單 ] 已有 ['+@Print_Date+' '+@OD_CName+'] 資料，若需要重轉則先行清除確認及未確認單據後重轉即可。sslpdt, @sd_class=['+@sd_class+'], @sd_slip_fg=['+@sd_slip_fg+'], @sd_date=['+@Print_Date+'], sd_lotno=['+@RM+']。'
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
          Set @Msg = '清除 '+@CompanyName+' 臨時轉入介面資料表 ['+@TB_tmp_Name+']。'

          Set @strSQL = 'DROP TABLE [dbo].['+@TB_tmp_Name+']'
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       end

       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
       --將每筆資料加入唯一鍵值(ps.此資料表請勿排序，保留原始 XLS 樣貌，以利日後產出對帳表)
       --建立暫存檔且建立序號，建立序號步驟很重要，因為後續要依序產生店名
print '11'
       Set @Msg = '建立臨時轉入介面資料表每筆唯一序號 ['+@TB_tmp_Name+']。'
       Set @strSQL = 'select *, print_date='''+@Print_Date+''', '+@CR+
                     '       SP_Exec_date=convert(date, getdate()) '+@CR+ -- 執行本程序日期
                     '  into '+@TB_tmp_Name+
                     '  from ['+@TB_xls_Name+'] '
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '12-1'
       Set @Msg = '刪除臨時轉入介面資料表之非必要資料 ['+@TB_tmp_Name+']。'
       Set @strSQL = 'update ['+@TB_tmp_Name+'] '+
                     '   set F1=F2 '+
                     ' where F2 like ''%對帳項目%'' '
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       
print '12'
       Set @Msg = '刪除臨時轉入介面資料表之非必要資料 ['+@TB_tmp_Name+']。'
       Set @strSQL = 'delete ['+@TB_tmp_Name+'] '+
                     ' where (Rtrim(F2) like ''%對帳明細表%'') '+
                     '    or (Rtrim(F2) like ''%對帳單%'') '+
                     '    or (Rtrim(F2) like ''%總額：%'') '+
                     '    or (Rtrim(F2) like ''商品名稱'') '+
                     '    or (Rtrim(F2) like ''END'') '
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
print '13'
       -- 此處會將每一筆的店名給寫入
       Set @Msg = '處理臨時轉入介面資料表之群組小計店名資料 ['+@TB_tmp_Name+']。'
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
       --Set @Msg = '刪除 F1 欄位為空白資料 ['+@TB_tmp_Name+']'

       --Set @strSQL = 'delete ['+@TB_tmp_Name+'] where isnull(f3, '''') ='''' and isnull(f6, '''') ='''' and isnull(f13, '''') ='''' '
       --Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
--=-[產生單據明細資料，數量為負數的要產生退貨單]=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
print '19'
       IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_OD_Name+']') AND type in (N'U'))
       begin
print '20'
          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
          Set @Msg = '刪除資料表 ['+@TB_OD_Name+']'
          Set @strSQL = 'DROP TABLE [dbo].['+@TB_OD_Name+']'
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       end

       --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
       --2013/04/25 雅婷說逐筆轉入就好不用加總，比較好對
       --在此進行比對商品編號，會拿輔助編號以及商品基本資料檔進行比對作業
print '21'
       Set @Msg = '進行比對商品編號並寫入 ['+@TB_OD_Name+']資料表。'
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
                     -- 2013/11/25 小嫻說金額要改為倒算回未稅金額
                     '       (Convert(Numeric(20, 6), F5)) / (Convert(Numeric(20, 6), F4)) / 1.05 as sale_amt, '+@CR+ -- 單價
                     '       (Convert(Numeric(20, 6), F5)) / 1.05 as sale_tot, '+@CR+ -- 小計
                     '       case when isnull(d.co_skno, d1.sk_no) is null then ''N'' else ''Y'' end as isfound '+@CR+
                     '       into '+@TB_OD_Name+' '+@CR+
                     '  from (select * '+@CR+
                     '          from (select ct_no, rtrim(ct_sname) as ct_sname, rtrim(ct_name)+''#''+rtrim(ct_sname) as ct_ssname '+@CR+
                     '                  from SYNC_TA13.dbo.PCUST  '+@CR+
                     '                 where ct_class =''1''  '+@CR+
                     '                   and ct_name like '''+@CompanyLikeName+''' '+@CR+
                     '                   and substring(ct_no, 9, 1) =''3'' '+@CR+ -- 一般為 1, 3C 為 2， 百貨為 3
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

       --產生出貨退回單
print '22'
       set @Cnt = 0
       Set @Msg = '檢查['+@TB_OD_Name+']待轉入資料是否存在。'
       set @strSQL = 'select count(1) from ['+@TB_OD_Name+'] '
       delete @RowCount
       print @strSQL
       insert into @RowCount Exec (@strSQL)

       -- 因為使用 Memory Variable Table, 所以使用 >0 判斷
       select @Cnt=cnt from @RowCount
print '23'
       if @Cnt >0
       begin
print '24'
          set @Msg = @Msg + '...存在，將進行轉入程序。'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', @Cnt
          --取得最新出貨退回單號(必須從已核准以及未核准的表單取號)
          --請勿再加+1，因為新增時會自動加入
          --2013/11/24 改為 @sd_date (每月固定使用 25日做為轉檔日)
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
          set @Msg = '取得最新出貨退回單號(必須從已核准以及未核准的表單取號),new_sdno=['+@new_sdno+'], sd_class=['+@sd_class+'], sd_slip_fg=['+@sd_slip_fg+'], sd_date=['+@sd_date+'].'
          Exec uSP_Sys_Write_Log @Proc, @Msg, '', @Cnt
           
                         
          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '25'
          set @Msg = '刪除['+@sd_date+'] 日轉入的單據明細。'
          
          set @strSQL ='delete TYEIPDBS2.lytdbta13.dbo.sslpdttmp '+@CR+
                       ' where Convert(Varchar(10), sd_date, 111) = '''+@sd_date+''' '+@CR+
                       '   and sd_class = '''+@sd_class+''' '+@CR+
                       '   and sd_slip_fg = '''+@sd_slip_fg+''' '+@CR+
                       '   and sd_lotno like ''%'+@Rm+'%'' '
           
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '26'
          set @Msg = '刪除['+@sd_date+'] 日轉入的單據主檔。'
          set @strSQL ='delete TYEIPDBS2.lytdbta13.dbo.ssliptmp '+@CR+
                       ' where Convert(Varchar(10), sp_date, 111) = '''+@sd_date+''' '+@CR+
                       '   and sp_class = '''+@sd_class+''' '+@CR+
                       '   and sp_slip_fg = '''+@sd_slip_fg+''' '+@CR+
                       '   and sp_rem like ''%'+@Rm+'%'' '
 
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '27'
          set @Msg = '新增['+@sd_date+'] 日出貨退回明細。'
          set @strSQL ='insert into TYEIPDBS2.lytdbta13.dbo.sslpdttmp '+@CR+
                       '(sd_class, sd_slip_fg, sd_date, sd_no, sd_ctno, '+@CR+
                       ' sd_skno, sd_name, sd_whno, sd_whno2, sd_qty, '+@CR+
                       ' sd_price, sd_dis, sd_stot, sd_lotno, sd_unit, '+@CR+
                       ' sd_unit_fg, sd_ave_p, sd_rate, sd_seqfld, sd_ordno '+@CR+
                       ') '+@CR+
                       'select '''+@sd_class+''' as sd_class, '+@CR+ -- 類別
                       '       '''+@sd_slip_fg+''' sd_slip_fg, '+@CR+ -- 單據種類
                       '       convert(datetime, '''+@sd_date+''') as sd_date, '+@CR+ -- 貨單日期
                       '       d2.sd_no, '+@CR+ -- 貨單編號
                       '       m.ct_no as sd_ctno, '+@CR+ -- 客戶編號
                       ''+@CR+
                       '       d.sk_no as sd_skno, '+@CR+ -- 貨品編號
                       '       d.sk_name as sd_name, '+@CR+ -- 品名規格
                       '       ''LB'' as sd_whno, '+@CR+ -- 倉庫(入)
                       '       '''' as sd_whno2, '+@CR+ -- 倉庫(出)
                       '       m.sale_qty as sd_qty, '+@CR+ -- 交易數量
                       ''+@CR+
                       '       m.sale_amt as sd_price, '+@CR+ -- 價格
                       '       1 as sd_dis, '+@CR+ -- 折數
                       '       m.sale_tot as sd_stot, '+@CR+ -- 小計
                       --2013/5/3 為了抓取重複資料並回寫到備註內，所以改寫到 sd_lotno 此欄位目前似乎未使用，只好當作臨時備註使用
                       --'       '''+@Rm+'於 ''+Convert(Varchar(20), Getdate(), 120)+char(13)+char(10)+'''+@OD_CName +''' as sd_rem, '+@CR+ -- 備用欄位
                       '       '''+@Rm+@OD_CName +''' as sd_lotno, '+@CR+ -- 備用欄位
                       '       d.sk_unit as sd_unit, '+@CR+ -- 單位
                       ' '+@CR+
                       '       0 as sd_unit_fg, '+@CR+ -- 單位旗標
                       '       d.s_price4 as sd_ave_p, '+@CR+ -- 單位成本
                       '       1 as sd_rate, '+@CR+ -- 匯率
                       '       rowid as sd_seqfld, '+@CR+ -- 明細序號, 此序號會因為凌越修改而加以變更，所以另存一份到 sd_ordno
                       '       rowid as sd_ordno'+@CR+ -- XLS 明細序號
                       --'       row_number() over(order by m.ct_no, d.sk_no) as sd_seqfld '+@CR+
                       '  from ['+@TB_OD_Name+'] m '+@CR+
                       '       left join SYNC_TA13.dbo.sstock d '+@CR+
                       '         on m.sk_no = d.sk_no '+@CR+
                       '       left join SYNC_TA13.dbo.pcust d1 '+@CR+
                       '         on m.ct_no = d1.ct_no '+@CR+
                       '        and ct_class =''1'' '+@CR+
                       '       left join  '+@CR+
                       '        (select ct_no,  '+@CR+
                       '                convert(varchar(10), '+@new_sdno+'+row_number() over(order by m.ct_no)) as sd_no '+@CR+ -- 貨單編號
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
          set @Msg = '回寫商品比對重複備註。'
          set @strSQL ='update TYEIPDBS2.lytdbta13.dbo.sslpdttmp '+@CR+
                       '   set sd_rem='''+@Rm+@OD_CName+'商品比對重複 RowID:''+D.Rowid '+@CR+
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
          -- 新增出貨退回主檔 PKey: sp_no (ASC), sp_slip_fg (ASC)
print '29'
          set @Msg = '新增出貨退回主檔。'
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
                       '       d1.ct_no as sp_ctno, '+@CR+ -- 客戶編號
                       '       d1.ct_name as sp_ctname, '+@CR+ -- 客戶名稱
                       '       substring(d1.ct_addr3, 1,255) as sp_ctadd2, '+@CR+ -- 送貨地址
                       '       d1.ct_sales as sp_sales, '+@CR+ -- 業務員
                       '       d1.ct_dept as sp_dpno, '+@CR+ -- 部門編號
                       ' '+@CR+
                       '       '''+@sp_maker+''' as sp_maker, '+@CR+ -- 製單人員
                       '       d1.ct_porter as sp_conv, '+@CR+ -- 貨運公司
                       '       sum(sd_stot) as sp_tot, '+@CR+ -- 小計
                       
                       -- 2013/6/10 托售回貨不需要稅金
                       '       case when sd_class =''3'' and sd_slip_fg=''7'' then 0 else sum(sd_stot)*0.05 end as sp_tax, '+@CR+ -- 營業稅(原)
                       
                       '       0 as sp_dis, '+@CR+ -- 折讓金額(原)
                       ' '+@CR+
                       '       d1.ct_payfg as sp_pay_kd, '+@CR+ -- 售價種類
                       '       ''NT'' as sp_rate_nm, '+@CR+ -- 匯率名稱
                       '       1 as sp_rate, '+@CR+ -- 匯率
                       '       sum(sd_stot) as sp_itot, '+@CR+ -- 發票金額
                       '       1 as sp_inv_kd, '+@CR+ -- 發票類別(=1  三聯式,=2  二聯式,=3  收銀機)
                       ' '+@CR+
                       '       1 as sp_tax_kd, '+@CR+ -- 稅別(=1應稅,=2零稅 )
                       '       1 as sp_invtype, '+@CR+ -- 開立方式(=1未開, =2隨單開立, =3批次開立)
                       '       '''+@Rm+'於 ''+Convert(Varchar(20), Getdate(), 120)+char(13)+char(10)+'''+@OD_CName+''' as sp_rem '+@CR+ -- 備註
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
                       '       d1.ct_no, '+@CR+ -- 客戶編號
                       '       d1.ct_name, '+@CR+ -- 客戶名稱
                       '       substring(d1.ct_addr3, 1, 255), '+@CR+ -- 送貨地址
                       '       d1.ct_sales, '+@CR+ -- 業務員
                       '       d1.ct_dept, '+@CR+ -- 部門編號
                       '       d1.ct_porter, '+@CR+ -- 貨運公司
                       '       d1.ct_payfg '

          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
          -- 2015/11/25 Rickliu 增加刪除無頭的明細資料，避免客服人員人工打出貨退回明細檔時，產生因轉檔失敗的資料出現。
          -- 2017/07/17 Rickliu 自從資料同步改為訂閱方式後，不知何故，跨DB進行 Update 會常態性失敗並且產生以下錯誤訊息，實無找不到問題所在，只好改寫為 Cursor 方式
          -- 錯誤訊息:無法從連結伺服器 "TYEIPDBS2" 的 OLE DB 提供者 "SQLNCLI10" 取得資料列的資料。
          set @Msg = '刪除 ['+@sd_date+'] 無頭之出貨退回明細檔。'
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
          set @Msg = @Msg + '...不存在，終止轉入程序。'
          Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

          -- 2013/11/28 增加失敗回傳值
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
  -- 2013/11/28 增加失敗回傳值
  Return(0)

print '34'
end
GO
