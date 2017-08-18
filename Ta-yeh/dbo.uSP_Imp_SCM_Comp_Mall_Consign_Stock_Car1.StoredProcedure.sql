USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_SCM_Comp_Mall_Consign_Stock_Car1]    Script Date: 08/18/2017 17:43:40 ******/
DROP PROCEDURE [dbo].[uSP_Imp_SCM_Comp_Mall_Consign_Stock_Car1]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_Imp_SCM_Comp_Mall_Consign_Stock_Car1]
as
begin
  /***********************************************************************************************************
     車麗屋後台拉單時，請選擇依分店抓取資料，如此 XLS 格式才會正確
     2013/11/28 增加失敗回傳值
  ***********************************************************************************************************/
  Declare @Proc Varchar(50) = 'uSP_Imp_SCM_Comp_Mall_Consign_Stock_Car1'
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
  
  Set @CompanyLikeCode = 'I0002%'
  Set @CompanyName = '車麗屋'
  Set @CompanyLikeName = '%'+@CompanyName+'%'
  Set @Rm = '系統匯入'+@CompanyName
  Set @sp_maker = 'Admin'

print '1'
  Declare Cur_Car1_Stock_Comp_DataKind cursor for
    select *
      from (select '2' as Kind, '3C' as KindName
             union
            select '3' as Kind, 'Retail' as KindName
           )m
     
print '2'
  open Cur_Car1_Stock_Comp_DataKind
  fetch next from Cur_Car1_Stock_Comp_DataKind into @Kind, @KindName

print '3'
  while @@fetch_status =0
  begin
print '4'
    Set @xls_head = 'Ori_Xls#'
    Set @TB_head = 'Comp_Mall_Consign_Stock_Car1'
    Set @TB_head_Kind = @TB_head+'_'+@KindName
    Set @TB_xls_Name = @xls_head+@TB_Head_Kind -- Ex: Ori_Xls#Comp_Mall_Consign_Stock_Car1_3C -- Excel 原檔資料
    Set @TB_tmp_Name = @TB_Head_Kind+'_tmp' -- Ex: Comp_Mall_Consign_Stock_Car1_3C_tmp -- 臨時轉入介面
    Set @TB_OD_Name = @TB_Head_Kind
    
    -- Check 匯入檔案是否存在
    IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_xls_Name+']') AND type in (N'U'))
    begin
print '5'
       --Send Message
       set @Cnt = @Errcode
       set @strSQL = ''
       Set @Msg = '外部 Excel 匯入資料表 ['+@TB_xls_Name+']不存在，終止進行轉檔作業。'
       Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
       -- 2013/11/28 增加失敗回傳值
       close Cur_Car1_Stock_Comp_DataKind
       deallocate Cur_Car1_Stock_Comp_DataKind
       Return(@Errcode)
       
--       fetch next from Cur_Car1_Stock_Comp_DataKind into @Kind, @KindName
--       Continue
    end
      
    -- 判別 Excel 列印日期 是否存在
print '6'
    set @Cnt = 0
    set @Print_Date  = ''
    
    set @strSQL = 'Select Replace(F1, ''列印日期：'', '''') as Print_date from [dbo].['+@TB_xls_Name+'] where F1 like ''列印日期%'' '
    set @Msg = '判別 Excel 列印日期 是否存在'
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
    
    delete @RowData
    insert into @RowData exec (@strSQL)
    select @Print_Date=Rtrim(isnull(aData, '')) from @RowData
    set @Print_Date = Convert(Varchar(10), CONVERT(Date, @Print_Date) , 111)
    
    --2013/11/25 林課確定每月固定以24號為轉檔日 
    set @Print_Date = Substring(Convert(Varchar(10), @Print_Date, 111), 1, 8)+@Last_Date
    
    if Rtrim(@Print_Date) = ''
    begin
print '4'
       set @Cnt = @Errcode
       set @strSQL = ''
       Set @Msg = '找不到 Excel 資料內的列印日期，終止進行轉檔作業。'
       Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
       -- 2013/11/28 增加失敗回傳值
       close Cur_Car1_Stock_Comp_DataKind
       deallocate Cur_Car1_Stock_Comp_DataKind
       Return(@Errcode)

--       fetch next from Cur_Car1_Stock_Comp_DataKind into @Kind, @KindName
--       Continue
    end
    else
    begin
/*    
print '7'
      IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_OD_Name+']') AND type in (N'U'))
      begin
         Set @strSQL = 'select Count(*) as cnt from '+@TB_OD_Name+' where print_date ='''+@Print_Date+''' '
         delete @RowCount
         insert into @RowCount Exec(@strSQL)
    
         if (select cnt from @RowCount) > 0
         begin
print '8'
          set @Msg = '最近一次列印日期相同則不進行轉檔作業!!'
          set @Cnt = @Errcode
          Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
          -- 2013/11/28 增加失敗回傳值
          fetch next from Cur_Car1_Stock_Comp_DataKind into @Kind, @KindName
          Continue
         end
      end
*/

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
                    ' where rowid <= 2 '
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    
      --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '13'
      Set @Msg = '處理臨時轉入介面資料表之群組小計店名資料 ['+@TB_tmp_Name+']。'
      Set @strSQL = 'declare cur_book cursor for '+@CR+
                    ' select rowid, f1 '+@CR+
                    '   from '+@TB_tmp_Name+@CR+
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
                    '     update '+@TB_tmp_Name+@CR+
                    '      set f1 = @vf1 '+@CR+
                    '    where rowid = @rowid '+@CR+
                    '  end '+@CR+
                    '  fetch next from cur_book into @rowid, @f1 '+@CR+
                    'end '+@CR+
                    'Close cur_book '+@CR+
                    'Deallocate cur_book '
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
        
      --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
      -- 去除臨時轉入介面資料表之店名序號


-- 2015/01/12 Rickliu 店名前兩碼為數字，一律去除不用    
/*
print '14'
      Set @str = ' 1234567890'
      set @Pos = 1
      While @Pos < Len(@str)  
      begin
print '15'
         Set @Msg = '去除 F1 欄位內含有數字['+Substring(@str, @Pos, 1)+']資料轉為空白 ['+@TB_tmp_Name+']'
    
         Set @strSQL = 'update '+@TB_tmp_Name+' set f1=replace(f1, '''+Substring(@str, @Pos, 1)+''', ''''), f7=replace(f7, '','', ''''), f17=replace(f17, '','', '''') '
         Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
        
         Set @Pos = @Pos +1
      end
*/    
      --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '16'
      Set @Msg = '去除 F1, F10~F17 欄位內含有"店" "," 字眼為空白 ['+@TB_tmp_Name+']'
      --Set @strSQL = 'update '+@TB_tmp_Name+' set f1 = substring(replace(f1, ''店'', ''''), 1, 2) '
      Set @strSQL = 'update '+@TB_tmp_Name+@CR+
                    ' set F1 = substring(replace(F1, ''店'', ''''), 3, Len(F1)), '+@CR+
                    '     F10= replace(F10, '','', ''''), '+@CR+
                    '     F11= replace(F11, '','', ''''), '+@CR+
                    '     F12= replace(F12, '','', ''''), '+@CR+
                    '     F13= replace(F13, '','', ''''), '+@CR+
                    '     F14= replace(F14, '','', ''''), '+@CR+
                    '     F15= replace(F15, '','', ''''), '+@CR+
                    '     F16= replace(F16, '','', ''''), '+@CR+
                    '     F17= replace(F17, '','', '''') '
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
         
      --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '15'
      Set @Msg = '刪除 F3, F10, F13 欄位為空白資料 ['+@TB_tmp_Name+']'
    
      Set @strSQL = 'delete '+@TB_tmp_Name+@CR+
                    ' where (isnull(f3, '''') ='''' '+@CR+
                    '   and isnull(f10, '''') ='''' '+@CR+
                    '   and isnull(f13, '''') ='''') '+@CR+
                    '    Or (F3 Like ''%商品條碼%'') '
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
         
print '17'
      IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_OD_Name+']') AND type in (N'U'))
      begin
print '18'
         --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
         Set @Msg = '重建'+@TB_OD_Name+'對帳總表 ['+@TB_head+']。'
    
         Set @strSQL = 'Create Table '+@TB_OD_Name+' ('+@CR+
                       '  	[Kind1] [varchar](1) NOT NULL, '+@CR+
                       '    [rowid] [int] NOT NULL, '+@CR+
	                   '    [ct_no] [varchar](50) NULL, '+@CR+
	                   '    [ct_sname] [varchar](50) NULL, '+@CR+
	                   '    [ct_ssname] [varchar](50) NULL, '+@CR+
	                   '    [sk_no] [varchar](50) NULL, '+@CR+
	                   '    [sk_name] [varchar](50) NULL, '+@CR+
	                   '    [F7] [varchar](50) NULL, '+@CR+
	                   '    [sk_bcode] [varchar](50) NULL, '+@CR+
	                   '    [F3] [varchar](50) NULL, '+@CR+
	                   '    [fg6_qty] [numeric](38, 6) NOT NULL, '+@CR+
	                   '    [fg7_qty] [numeric](38, 6) NOT NULL, '+@CR+
	                   '    [sum_qty] [numeric](38, 6) NOT NULL, '+@CR+
	                   '    [fg6_amt] [numeric](38, 6) NOT NULL, '+@CR+
	                   '    [fg7_amt] [numeric](38, 6) NOT NULL, '+@CR+
	                   '    [sum_amt] [numeric](38, 6) NOT NULL, '+@CR+
	                   '    [F13] [numeric](20, 6) NOT NULL, '+@CR+
	                   '    [F17] [numeric](20, 6) NOT NULL, '+@CR+
	                   '    [diff_qty] [numeric](38, 6) NOT NULL, '+@CR+
	                   '    [diff_amt] [numeric](38, 6) NOT NULL, '+@CR+
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
         
      --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
      --2015/01/13 副總指示要產出車麗屋商品匯總對帳表，此表僅依商品進行交叉比對呈現結果，Kind1: 1 為商品匯總總表, 2 為店+商品彙總總表
      
print '19'
      Set @Msg = '進行比對商品編號並寫入 ['+@TB_OD_Name+']資料表。'
      set @strSQL = 
                'Insert Into '+@TB_OD_Name+' '+@CR+
                '(Kind1, rowid, ct_no, ct_sname, ct_ssname, sk_no, '+@CR+
                ' sk_name, f7, sk_bcode, f3, fg6_qty, '+@CR+
                ' fg7_qty, sum_qty, fg6_amt, fg7_amt, sum_amt, '+@CR+
                ' F13, F17, diff_qty, diff_amt, Print_Date, '+@CR+
                ' isfound, Exec_DateTime) '+@CR+

                'select Distinct '+@CR+
                '       ''1'' as Kind1, '+@CR+
                '       0 as rowid, '+@CR+
                '       Convert(Varchar(50), '''') as ct_no, '+@CR+
                '       Convert(Varchar(50), '''') as ct_sname, '+@CR+
                '       Convert(Varchar(50), '''') as ct_ssname, '+@CR+
                '       Convert(Varchar(50), RTrim(isnull(isnull(d.co_skno, d1.sk_no), ''N/A''))) as sk_no, '+@CR+
                -- 商品名稱
                '       Convert(Varchar(50), RTrim(isnull(isnull(d.sk_name, d1.sk_name), ''N/A''))) as sk_name, '+@CR+
                -- 商品名稱(Excel)
                '       Convert(Varchar(50), RTrim(F7)) as F7, '+@CR+ 
                -- 商品條碼
                '       Convert(Varchar(50), RTrim(isnull(isnull(d.sk_bcode, d1.sk_bcode), ''N/A''))) as sk_bcode, '+@CR+
                -- 商品條碼(Excel)
                '       Convert(Varchar(50), RTrim(F3)) as F3, '+@CR+ 
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
                '       isnull(Convert(Numeric(20, 6), F13), 0) as F13, '+@CR+
                -- 寄賣金額(Excel)
                '       isnull(Convert(Numeric(20, 6), F17), 0) as F17, '+@CR+
                -- 差異數量
                '       isnull(sum_qty - Convert(Numeric(20, 6), F13), 0) as diff_qty, '+@CR+
                -- 差異金額
                '       isnull(sum_amt - Convert(Numeric(20, 6), F17), 0) as diff_amt, '+@CR+
                '       '''+@Print_Date+''' as Print_Date, '+@CR+
                -- 異常
                '       case '+@CR+
                '         when (isnull(d.co_skno, d1.sk_no) is null) then ''XLS 編號對不到凌越商品資料'' '+@CR+
                '         when (fg6_qty is null) then ''XLS 有資料，但凌越無託售回貨資料'' '+@CR+
                '         when (F3 is null and sum_qty <>0) then ''凌越有託售回貨資料，但XLS無資料'' '+@CR+
                '         when (sum_qty - F13 <>0) then ''比對數量異常'' '+@CR+
                '         else '''' '+@CR+
                '       end as isfound, '+@CR+
                '       getdate() as Exec_DateTime '+@CR+
                -- 客戶編號
                '  from (Select F3, F7, '+@CR+
                '               Sum(isnull(Convert(Numeric(20, 6), Replace(F13, '','', '''')), 0)) as F13, '+@CR+
                '               Sum(isnull(Convert(Numeric(20, 6), Replace(F17, '','', '''')), 0)) as F17 '+@CR+
                '          from '+@TB_tmp_Name+@CR+
                '         Group by F3, F7 '+@CR+
                '       )m '+@CR+
                -- 商品資料
                '       Full join '+@CR+ 
                '       (select distinct '+@CR+ 
                '               RTrim(co_ctno) as co_ctno, '+@CR+ 
                '               RTrim(co_skno) as co_skno, '+@CR+  
                '               RTrim(co_cono) as co_cono, '+@CR+  
                '               RTrim(sk_no) as sk_no, '+@CR+  
                '               RTrim(sk_name) as sk_name, '+@CR+  
                '               RTrim(sk_bcode) as sk_bcode '+@CR+ 
                '          from SYNC_TA13.dbo.sauxf m '+@CR+  
                '               left join SYNC_TA13.dbo.sstock d '+@CR+  
                '                 on m.co_skno = d.sk_no '+@CR+  
                '                and m.co_ctno like '''+@CompanyLikeCode+''' '+@CR+ 
                '         where 1=1 '+@CR+ 
                '           and co_class=''1'' '+@CR+  
                -- 只查託售回貨部分
                '           and exists '+@CR+ 
                '               (select * '+@CR+ 
                '                  from SYNC_TA13.dbo.sslpdt d1 '+@CR+ 
                '                 where 1=1 '+@CR+ 
                '                   and sd_slip_fg IN (''6'',''7'') '+@CR+   
                '                   and d.sk_no = d1.sd_skno '+@CR+ 
                '                   and d1.sd_ctno like '''+@CompanyLikeCode+''' '+@CR+ 
                '                   and d1.sd_date <= '''+@Print_Date+'''  '+@CR+
                '               ) '+@CR+
                '         ) d '+@CR+
                '         on 1=1 '+@CR+
                '        and (ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.co_skno)) collate Chinese_Taiwan_Stroke_CI_AS  '+@CR+
                '         or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.co_cono)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                '         or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.sk_no)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                '         or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.sk_bcode)) collate Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                '       left join SYNC_TA13.dbo.sstock d1 '+@CR+
                '         on 1=1 '+@CR+
                '        and (ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d1.sk_no)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                '         or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d1.sk_bcode)) collate Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                --  凌越托售回貨資料 sp_slip_fg = 6 托售, sp_slip_fg = 7 回貨
                '       left join  '+@CR+
                '         (select sd_skno,  '+@CR+
                '                 sum(case when sd_slip_fg = ''6'' then sd_qty else 0 end) as fg6_qty, '+@CR+ -- 托售數量
                '                 sum(case when sd_slip_fg = ''7'' then sd_qty else 0 end) as fg7_qty, '+@CR+ -- 回貨數量
                '                 sum(case when sd_slip_fg = ''7'' then sd_qty * -1 else sd_qty end) as sum_qty, '+@CR+ -- 合計數量<-------CHECK
                '                 sum(case when sd_slip_fg = ''6'' then sd_stot else 0 end) as fg6_amt, '+@CR+ -- 托售金額
                '                 sum(case when sd_slip_fg = ''7'' then sd_stot else 0 end) as fg7_amt, '+@CR+ -- 托售數量
                '                 sum(case when sd_slip_fg = ''7'' then sd_stot * -1 else sd_stot end) as sum_amt '+@CR+ -- 合計金額
                '            from SYNC_TA13.dbo.sslpdt '+@CR+
                '           where 1=1 '+@CR+
                '             and sd_slip_fg IN (''6'',''7'')  '+@CR+
                '             and sd_date >= ''2012/12/01''  '+@CR+
                -- 範圍截取至 2013/01/01 ~ 列印日期
                '             and sd_date <= '''+@Print_Date+'''  '+@CR+
                '             and sd_ctno like '''+@CompanyLikeCode+''' '+@CR+ 
                '           group by sd_skno '+@CR+
                '         ) d2 '+@CR+
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
      --2013/04/25 雅婷說逐筆轉入就好不用加總，比較好對
      --在此進行比對商品編號，會拿輔助編號以及商品基本資料檔進行比對作業
print '17'
      Set @Msg = '進行比對商品編號並寫入 ['+@TB_OD_Name+']資料表。'
      set @strSQL = 
                'Insert Into '+@TB_OD_Name+' '+@CR+
                '(Kind1, rowid, ct_no, ct_sname, ct_ssname, sk_no, '+@CR+
                ' sk_name, f7, sk_bcode, f3, fg6_qty, '+@CR+
                ' fg7_qty, sum_qty, fg6_amt, fg7_amt, sum_amt, '+@CR+
                ' F13, F17, diff_qty, diff_amt, Print_Date, '+@CR+
                ' isfound, Exec_DateTime) '+@CR+
                
                'select Distinct '+@CR+
                '       ''2'' as Kind1, rowid, '+@CR+
                '       Convert(Varchar(50), RTrim(ct_no)) as ct_no, '+@CR+
                '       Convert(Varchar(50), RTrim(ct_sname)) as ct_sname, '+@CR+
                '       Convert(Varchar(50), RTrim(ct_ssname)) as ct_ssname, '+@CR+
                '       Convert(Varchar(50), RTrim(isnull(isnull(d.co_skno, d1.sk_no), ''N/A''))) as sk_no, '+@CR+
                -- 商品名稱
                '       Convert(Varchar(50), RTrim(isnull(isnull(d.sk_name, d1.sk_name), ''N/A''))) as sk_name, '+@CR+
                -- 商品名稱(Excel)
                '       Convert(Varchar(50), RTrim(F7)) as F7, '+@CR+ 
                -- 商品條碼
                '       Convert(Varchar(50), RTrim(isnull(isnull(d.sk_bcode, d1.sk_bcode), ''N/A''))) as sk_bcode, '+@CR+
                -- 商品條碼(Excel)
                '       Convert(Varchar(50), RTrim(F3)) as F3, '+@CR+ 
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
                '       isnull(Convert(Numeric(20, 6), F13), 0) as F13, '+@CR+
                -- 寄賣金額(Excel)
                '       isnull(Convert(Numeric(20, 6), F17), 0) as F17, '+@CR+
                -- 差異數量
                '       isnull(sum_qty - Convert(Numeric(20, 6), F13), 0) as diff_qty, '+@CR+
                -- 差異金額
                '       isnull(sum_amt - Convert(Numeric(20, 6), F17), 0) as diff_amt, '+@CR+
                '       '''+@Print_Date+''' as Print_Date, '+@CR+
                -- 異常
                '       case '+@CR+
                '         when (isnull(d.co_skno, d1.sk_no) is null) then ''編號對不到'' '+@CR+
                '         when (fg6_qty is null) then ''無凌越銷售明細'' '+@CR+
                '         else '''' '+@CR+
                '       end as isfound, '+@CR+
                '       getdate() as Exec_DateTime '+@CR+
                -- 客戶編號
                '  from (select * '+@CR+
                '          from (select ct_no, rtrim(ct_sname) as ct_sname, rtrim(ct_name)+''#''+rtrim(ct_sname) as ct_ssname '+@CR+
                '                  from SYNC_TA13.dbo.PCUST  '+@CR+
                '                 where ct_class =''1''  '+@CR+
                '                   and ct_name like '''+@CompanyLikeName+''' '+@CR+
                '                   and substring(ct_no, 9, 1) ='''+@Kind+''' '+@CR+ -- 3C 為 2， 百貨為 3
                '                   and ct_no like '''+@CompanyLikeCode+''' '+@CR+ 
                '                ) m, '+@TB_tmp_Name+' d '+@CR+
                '         where m.ct_ssname collate Chinese_Taiwan_Stroke_CI_AS like ''%''+rtrim(d.f1)+''%'' collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                '       )m '+@CR+
                -- 商品資料
                '       left join '+@CR+ 
                '       (select distinct '+@CR+ 
                '               RTrim(co_ctno) as co_ctno, '+@CR+ 
                '               RTrim(co_skno) as co_skno, '+@CR+  
                '               RTrim(co_cono) as co_cono, '+@CR+  
                '               RTrim(sk_no) as sk_no, '+@CR+  
                '               RTrim(sk_name) as sk_name, '+@CR+  
                '               RTrim(sk_bcode) as sk_bcode '+@CR+ 
                '          from SYNC_TA13.dbo.sauxf m '+@CR+  
                '               left join SYNC_TA13.dbo.sstock d '+@CR+  
                '                 on m.co_skno = d.sk_no '+@CR+  
                '                and m.co_ctno like '''+@CompanyLikeCode+''' '+@CR+ 
                '         where 1=1 '+@CR+ 
                '           and co_class=''1'' '+@CR+  
                -- 只查託售回貨部分
                '           and exists '+@CR+ 
                '               (select * '+@CR+ 
                '                  from SYNC_TA13.dbo.sslpdt d1 '+@CR+ 
                '                 where 1=1 '+@CR+ 
                '                   and sd_slip_fg IN (''6'',''7'') '+@CR+   
                '                   and d.sk_no = d1.sd_skno '+@CR+ 
                '                   and d1.sd_ctno like '''+@CompanyLikeCode+''' '+@CR+ 
                '                   and d1.sd_date <= '''+@Print_Date+'''  '+@CR+
                '               ) '+@CR+
                '         ) d '+@CR+
                '         on 1=1 '+@CR+
                '        and m.ct_no=d.co_ctno '+@CR+
                '        and (ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.co_skno)) collate Chinese_Taiwan_Stroke_CI_AS  '+@CR+
                '         or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.co_cono)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                '         or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.sk_no)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                '         or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d.sk_bcode)) collate Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                '       left join SYNC_TA13.dbo.sstock d1 '+@CR+
                '         on 1=1 '+@CR+
                '        and (ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d1.sk_no)) collate Chinese_Taiwan_Stroke_CI_AS '+@CR+
                '         or  ltrim(rtrim(m.f3)) collate Chinese_Taiwan_Stroke_CI_AS = ltrim(rtrim(d1.sk_bcode)) collate Chinese_Taiwan_Stroke_CI_AS) '+@CR+
                --  凌越托售回貨資料 sp_slip_fg = 6 托售, sp_slip_fg = 7 回貨
                '       left join  '+@CR+
                '         (select sd_ctno, sd_skno,  '+@CR+
                '                 sum(case when sd_slip_fg = ''6'' then sd_qty else 0 end) as fg6_qty, '+@CR+ -- 托售數量
                '                 sum(case when sd_slip_fg = ''7'' then sd_qty else 0 end) as fg7_qty, '+@CR+ -- 回貨數量
                '                 sum(case when sd_slip_fg = ''7'' then sd_qty * -1 else sd_qty end) as sum_qty, '+@CR+ -- 合計數量<-------CHECK
                '                 sum(case when sd_slip_fg = ''6'' then sd_stot else 0 end) as fg6_amt, '+@CR+ -- 托售金額
                '                 sum(case when sd_slip_fg = ''7'' then sd_stot else 0 end) as fg7_amt, '+@CR+ -- 托售數量
                '                 sum(case when sd_slip_fg = ''7'' then sd_stot * -1 else sd_stot end) as sum_amt '+@CR+ -- 合計金額
                '            from SYNC_TA13.dbo.sslpdt '+@CR+
                '           where 1=1 '+@CR+
                '             and sd_slip_fg IN (''6'',''7'')  '+@CR+
                '             and sd_date >= ''2012/12/01''  '+@CR+
                -- 範圍截取至 2013/01/01 ~ 列印日期
                '             and sd_date <= '''+@Print_Date+'''  '+@CR+
                '             and sd_ctno like '''+@CompanyLikeCode+''' '+@CR+ 
                '           group by sd_ctno, sd_skno '+@CR+
                '         ) d2 '+@CR+
                '         on 1=1 '+@CR+
                '        and isnull(d.co_skno, d1.sk_no) = d2.sd_skno '+@CR+
                '        and ct_no = d2.sd_ctno '+@CR+
                ' order by 1, 2 '
                --group by ct_no, ct_sname, ct_ssname, sk_no, sk_name, sk_bcode, f3, f6
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    
      --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
      --2013/12/24將重複的資料也列入異常
print '20'
      set @Msg = '將重複的資料標註異常'
      set @Cnt = 0
      set @strSQL ='update '+@TB_OD_Name+@CR+
                   '   set isfound = isfound+''，重複比對'' '+@CR+
                   ' where 1=1 '+@CR+
                   '   and Kind1 = ''2'' '+@CR+
                   '   and rowid in '+@CR+
                   '       (select rowid '+@CR+
                   '          from '+@TB_OD_Name+@CR+
                   '         where print_date = '''+@Print_Date+''' '+@CR+
                   '         group by rowid '+@CR+
                   '        having COUNT(*) > 1) '
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
/*    
print '21'
      --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
      --2014/11/20 Rickliu 產生對帳總表
      set @Msg = '產生對帳總表'
      set @Cnt = 0
      set @strSQL ='update '+@TB_OD_Name+@CR+
                   '   set isfound = ''重複比對'' '+@CR+
                   ' where rowid in '+@CR+
                   '       (select rowid '+@CR+
                   '          from '+@TB_OD_Name+@CR+
                   '         where print_date = '''+@Print_Date+''' '+@CR+
                   '         group by rowid '+@CR+
                   '        having COUNT(*) > 1) '
*/    
      
      
      IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_head+']') AND type in (N'U'))
      begin
print '22'
         Set @Msg = '重建'+@CompanyName+'對帳總表 ['+@TB_head+']。'
    
         Set @strSQL = 'Create Table '+@TB_head+' ('+@CR+
                       '  [Kind1] [varchar](10) NULL, '+@CR+
                       '  [Kind2] [varchar](10) NULL, '+@CR+
                       '  [sd_ctno] [varchar](50) NULL, '+@CR+
                       '  [ct_sname] [nchar](50) NULL, '+@CR+
                       '  [sk_no] [varchar](30) NULL, '+@CR+
                       '  [sk_name] [nchar](50) NULL, '+@CR+
                       '  [xls_skno] [varchar](50) NOT NULL, '+@CR+
                       '  [xls_skname] [varchar](50) NOT NULL, '+@CR+
                       '  [sk_bcode] [varchar](50) NULL, '+@CR+
                       '  [xls_bcode] [varchar](50) NOT NULL, '+@CR+
                       '  [rowid] [int] NULL, '+@CR+
                       '  [Chg_sd_Qty] [numeric](18, 2) NULL, '+@CR+
                       '  [Chg_sd_stot] [numeric](18, 2) NULL, '+@CR+
                       '  [xls_qty] [numeric](18, 2) NOT NULL, '+@CR+
                       '  [xls_amt] [numeric](18, 2) NOT NULL, '+@CR+
                       '  [diff_qty] [numeric](18, 2) NULL, '+@CR+
                       '  [diff_amt] [numeric](18, 2) NULL, '+@CR+
                       '  [Print_Date] [varchar](10) NULL, '+@CR+
                       '  [isfound] [varchar](20) NOT NULL, '+@CR+
                       '  [Exec_DateTime] [DateTime] NULL '+@CR+
                       ') ON [PRIMARY] '
         Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
      end
      
      Set @Msg = '清除'+@CompanyName+'對帳總表 ['+@TB_head+']。'
      set @strSQL ='Delete '+@TB_head+' '+@CR+
                   ' where Kind2 = '''+@KindName+''' '+@CR+
                   '   and Print_Date = '''+@Print_Date+''' '                   
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
	  
      Set @Msg = '寫入'+@CompanyName+'對帳總表 ['+@TB_head+']。'
      set @strSQL ='Insert Into '+@TB_head+' '+@CR+
                   '(Kind1, Kind2, sd_ctno, ct_sname, sk_no, sk_name, '+@CR+
                   ' xls_skno, xls_skname, sk_bcode, xls_bcode, rowid, '+@CR+
                   ' Chg_sd_Qty, Chg_sd_stot, xls_qty, xls_amt, diff_qty, '+@CR+
                   ' diff_amt, print_date, isfound, Exec_DateTime) '+@CR+
                   'select distinct '+@CR+
                   '       Kind1, '+@CR+ -- 比對店家+商品
                   '       '''+@KindName+''' as Kind2, '+@CR+
                   '       m.sd_ctno, m.ct_sname, '+@CR+
                   '       m.sd_skno as sk_no, d1.sk_name as sk_name,  '+@CR+
                   '       isnull(d.f3, ''N/A'') as xls_skno, '+@CR+
                   '       isnull(d.f7, ''N/A'') as xls_skname, '+@CR+
                   '       m.sk_bcode, '+@CR+
                   '       isnull(d.f3, ''N/A'') as xls_bcode, '+@CR+
                   '       d.rowid, '+@CR+
                   '       m.Chg_sd_Qty, '+@CR+
                   '       m.Chg_sd_stot, '+@CR+
                   '       isnull(d.F13, 0) as xls_qty, '+@CR+
                   '       isnull(d.F17, 0) as xls_amt, '+@CR+
                   '       isnull(d.F13, 0) - isnull(m.Chg_sd_Qty, 0)as diff_qty, '+@CR+
                   '       isnull(d.F17, 0) - isnull(m.Chg_sd_stot, 0) as diff_amt, '+@CR+
                   '       '''+@Print_Date+''' as Print_Date, '+@CR+
                   --'       (select top 1 isnull(print_date, convert(varchar(10), getdate(), 111)) from '+@TB_head_Kind+')as Print_Date, '+@CR+
                   '       isnull(d.isfound, ''非XLS資料'') as isfound, '+@CR+
                   '       GetDate() as Exec_DateTime '+@CR+
                   '  from (select sd_ctno, ct_sname, sd_skno, sk_bcode, '+@CR+
                   '               Sum(Chg_sd_qty) as Chg_sd_Qty, Sum(Chg_sd_stot) as Chg_sd_stot '+@CR+
                   '          from SYNC_TA13.dbo.v_orders m '+@CR+
                   '               left join SYNC_TA13.dbo.pcust d '+@CR+
                   '                 on m.sp_ctno=d.ct_no and d.ct_class=''1'' '+@CR+
                   '         where sd_slip_fg IN (''6'',''7'') '+@CR+
                   '           and sd_date >= ''2012/12/01'' '+@CR+
                   '           and sd_date <= '''+@Print_Date+''' '+@CR+
                   /*
                   '               (select top 1 max(isnull(print_date, convert(varchar(10), getdate(), 111))) '+@CR+
                   '                  from '+@TB_head_Kind+' '+@CR+
                   '                 where isnull(print_date, '''') <> '''')  '+@CR+
                   */
                   '            group by sd_ctno, ct_sname, sd_skno, sk_bcode '+@CR+
                   '       ) m '+@CR+
                   '      LEFT join '+@TB_head_Kind+' d '+@CR+
                   '        on 1=1 '+@CR+
                   '       and m.sd_skno collate Chinese_Taiwan_Stroke_CI_AS = d.sk_no '+@CR+
                   '       and m.sd_ctno collate Chinese_Taiwan_Stroke_CI_AS = d.ct_no '+@CR+
                   '       and d.print_date collate Chinese_Taiwan_Stroke_CI_AS ='''+@Print_Date+''' '+@CR+
                   '      left join SYNC_TA13.dbo.sstock d1 '+@CR+
                   '        on m.sd_skno collate Chinese_Taiwan_Stroke_CI_AS = d1.sk_no '+@CR+
                   ' Where 1=1 '+@CR+
                   '   and m.ct_sname like ''%'+@CompanyName+'%'' '+@CR+
                   '   and substring(m.sd_ctno, len(m.sd_ctno), 1) = '''+@Kind+''' '+@CR+
                   ' order by d.rowid, m.sd_ctno, m.sd_skno '
      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

    end
print '23'
    fetch next from Cur_Car1_Stock_Comp_DataKind into @Kind, @KindName
  end
  close Cur_Car1_Stock_Comp_DataKind
  deallocate Cur_Car1_Stock_Comp_DataKind
  -- 2013/11/28 增加失敗回傳值
  Return(0)
end
GO
