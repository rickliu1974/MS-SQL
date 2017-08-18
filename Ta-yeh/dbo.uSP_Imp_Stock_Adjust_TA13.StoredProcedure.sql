USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_Stock_Adjust_TA13]    Script Date: 08/18/2017 17:43:40 ******/
DROP PROCEDURE [dbo].[uSP_Imp_Stock_Adjust_TA13]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_Imp_Stock_Adjust_TA13]
as
begin
  /***********************************************************************************************************
     2013/12/06 -- Rickliu
     調整單轉入，請依調整 Excel 檔案及規範進行轉入。     
  ***********************************************************************************************************/
  Declare @Proc Varchar(50) = 'uSP_Imp_Stock_Adjust_TA13'
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

  Declare @SP_date varchar(10) -- 調整日期
  Declare @SP_Sales varchar(10) -- 製單員
  Declare @SP_dpno varchar(10) -- 部門編號

  Declare @SP_cnt int
       
  Declare @CompanyName Varchar(100), @CompanyLikeName Varchar(100), @Rm Varchar(100), @SP_maker Varchar(20), @str Varchar(200)
  Declare @Pos Int
  
  Set @CompanyName = '未確認調整單' --> 請勿亂變動
  Set @CompanyLikeName = '%'+@CompanyName+'%'
  Set @Rm = '系統匯入'+@CompanyName+' 於 '+Convert(Varchar(20), getdate(), 120)
  Set @SP_maker = 'Admin'

print '1'
  Set @xls_head = 'Ori_xls#'
  Set @TB_head = 'Stock_Adjust_TA13'
  Set @TB_xls_Name = @xls_head+@TB_Head -- Ex: Ori_xls#Stock_Adjust_TA13 -- Excel 原檔資料
  Set @TB_tmp_Name = @TB_Head+'_'+'tmp' -- Ex: Stock_Adjust_TA13_tmp -- 臨時轉入介面
  Set @TB_SP_Name = @TB_Head -- Ex: Stock_Adjust_TA13
  set @SP_Class = '5' -- 調整類
  set @sp_slip_fg = 'A' -- 調整單

  -----------------------------------------------------------------------------------------------------------------------------
  -- Check 匯入檔案是否存在
  IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_xls_Name+']') AND type in (N'U'))
  begin
print '4'
     set @Cnt = @Errcode
     set @strSQL = ''
     Set @Msg = '外部 Excel 匯入資料表 ['+@TB_xls_Name+'] 不存在，終止進行轉檔作業。'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

     Return(@Errcode)
  end
    
  -----------------------------------------------------------------------------------------------------------------------------
  -- 判別 Excel 是否有資料
  set @Cnt = 0
  set @strSQL ='Select Count(1) as cnt from [dbo].['+@TB_xls_Name+']'
  set @Msg = '判別 Excel 是否有資料...'
  delete @RowData
  insert into @RowCount exec (@strSQL)
  if (Select cnt from @RowCount) > 0
  begin
     set @Msg = @Msg + '存在'
  end
  else
  begin
     set @Cnt = @Errcode
     set @Msg = @Msg + '不存在'
  end
  
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

  if @Cnt = @Errcode  Return(@Errcode)
    
  -----------------------------------------------------------------------------------------------------------------------------
  -- 判別 Excel 調整日期 必要欄位是否存在 
print '5'
  set @Cnt = 0
  set @SP_Date  = ''

  set @strSQL = 'Select F2 as SP_Date from [dbo].['+@TB_xls_Name+'] where Rowid = 1 '
  set @Msg = '判別 Excel [調整日期] 是否存在'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

  delete @RowData
  insert into @RowData exec (@strSQL)
  select @SP_Date=Rtrim(isnull(aData, '')) from @RowData

  if @SP_Date = ''
  begin
print '6'
     set @Cnt = @Errcode
     set @strSQL = ''
     Set @Msg = '找不到 Excel 資料內的 [調整日期]，終止進行轉檔作業。'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

     Return(@Errcode)
  end
  -----------------------------------------------------------------------------------------------------------------------------
  -- 判別 Excel 調整員 必要欄位是否存在 
print '5'
  set @Cnt = 0
  set @SP_Sales  = ''

  set @strSQL = 'Select F4 as sp_sales from [dbo].['+@TB_xls_Name+'] where Rowid = 1 '
  set @Msg = '判別 Excel [調整員] 是否存在'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

  delete @RowData
  insert into @RowData exec (@strSQL)
  select @SP_Sales=Rtrim(isnull(aData, '')) from @RowData
    
  if @SP_Sales = ''
  begin
print '6'
     set @Cnt = @Errcode
     set @strSQL = ''
     Set @Msg = '找不到 Excel 資料內的 [調整員]，終止進行轉檔作業。'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

     Return(@Errcode)
  end
  -----------------------------------------------------------------------------------------------------------------------------
  -- 判別 Excel 部門編號 必要欄位是否存在 
print '5'
  set @Cnt = 0
  set @SP_dpno  = ''

  set @strSQL = 'Select F6 as sp_dpno from [dbo].['+@TB_xls_Name+'] where Rowid = 1 '
  set @Msg = '判別 Excel [部門編號] 是否存在'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

  delete @RowData
  insert into @RowData exec (@strSQL)
  select @SP_dpno=Rtrim(isnull(aData, '')) from @RowData

  if @SP_dpno = ''
  begin
print '6'
     set @Cnt = @Errcode
     set @strSQL = ''
     Set @Msg = '找不到 Excel 資料內的 [部門編號]，終止進行轉檔作業。'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

     Return(@Errcode)
  end

   
  -- 增加判別 確認 及 未確認 單據 若是存在則不得進行重轉作業。
  -- 2013/10/28 增加判斷若當月24日已由未確認轉確認單則不再重轉，避免凌越未確認單據不斷刪除新增造成單號累增情況
print '7'
  select @SP_cnt = sum(cnt)
    from (select count(1) as cnt
           from SYNC_TA13.dbo.sslip m
          where SP_class = @SP_class
            and sp_slip_fg = @sp_slip_fg
            and SP_date = @SP_date
            and SP_rem like '%'+@Rm+'%'
          union
         select count(1) as cnt
           from SYNC_TA13.dbo.ssliptmp m
          where SP_class = @SP_class
            and sp_slip_fg = @sp_slip_fg
            and SP_date = @SP_date
            and SP_rem like '%'+@Rm+'%') m

  if @SP_cnt > 0
  begin
print '8'
     set @Cnt = @Errcode
     set @strSQL = ''
     set @Msg ='[ 已確認 或 未確認 單 ] 已有 ['+@SP_Date+'] 資料，若需要重轉則先行清除確認及未確認單據後重轉即可。sslip, @SP_class=['+@SP_class+'], @SP_slip_fg=['+@SP_slip_fg+'], @SP_date=['+@SP_date+'], @SP_lotno=['+@RM+']。'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
     
     Return(@Errcode)
  end
  else
  begin
print '9'
     IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_tmp_Name+']') AND type in (N'U'))
     begin
print '10'
        Set @Msg = '重整 '+@CompanyName+' 臨時轉入介面資料表 ['+@TB_tmp_Name+']。'

        Set @strSQL = 'DROP TABLE [dbo].['+@TB_tmp_Name+']'
        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
     end

     --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
     --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
     -- 刪除前兩筆表頭資料
print '11'
     Set @Msg = '刪除前兩筆表頭資料以及刪除空白欄位 ['+@TB_tmp_Name+']。'
     Set @strSQL = 'select Isnull(LTrim(Rtrim(F1)), '''') as F1, '+@CR+ -- 商品編號
                   '       Isnull(LTrim(Rtrim(F2)), '''') as F2, '+@CR+ -- 商品名稱
                   '       Isnull(LTrim(Rtrim(F3)), '''') as F3, '+@CR+ -- 倉別
                   '       Case '+@CR+
                   '         when LTrim(RTrim(F4)) = '''' then ''0'' '+@CR+
                   '         else Isnull(LTrim(Rtrim(F4)), ''0'') '+@CR+
                   '       end as F4, '+@CR+ -- 調整數量
                   '       Isnull(LTrim(RTrim(F5)), ''1'') as F5, '+@CR+ -- 成本類別
                   '       SP_date='''+@SP_date+''', '+@CR+ -- 調整日期
                   '       SP_sales='''+@SP_Sales+''', '+@CR+ -- 調整員
                   '       SP_dpno='''+@SP_dpno+''', '+@CR+ -- 部門編號
                   '       import_date=convert(date, getdate()), '+@CR+
                   '       rowid '+@CR+ -- Xls 列號
                   '       into '+@TB_tmp_Name+@CR+
                   '  from ['+@TB_xls_Name+']'
     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

     --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
     --
print '22'
     set @Cnt = 0
     Set @Msg = '檢查['+@TB_tmp_Name+']單據資料是否存在。'
     set @strSQL = 'select count(1) from ['+@TB_tmp_Name+']'
     delete @RowCount
     print @strSQL
     insert into @RowCount Exec (@strSQL)

     -- 因為使用 Memory Variable Table, 所以使用 >0 判斷
print '23'
     if (select cnt from @RowCount) > 0
     begin
print '24'
        set @Msg = @Msg + '...不存在，將進行轉入程序。'
        Exec uSP_Sys_Write_Log @Proc, @Msg, '', @Cnt
        set @new_orno = (select convert(varchar(10), Convert(Int, isnull(max(SP_no), replace(substring(@SP_date, 3, 8), '/', '') +'0000'))+1)
                           from (select distinct SP_no
                                   from SYNC_TA13.dbo.sslip m
                                  where sp_class = @SP_class
                                    and sp_slip_fg = @sp_slip_fg
                                    and sp_date = @sp_date
                                  union
                                 select distinct SP_no
                                   from SYNC_TA13.dbo.ssliptmp m
                                  where sp_class = @SP_class
                                    and sp_slip_fg = @sp_slip_fg
                                    and sp_date = @sp_date
                                ) m
                        )
        set @Msg = '取得最新調整單號(必須從已核准以及未核准的表單取號),new_orno=['+@new_orno+'], SP_class=['+@SP_class+'], SP_slip_fg=['+@SP_slip_fg+'], SP_date=['+@SP_date+'].'
        Exec uSP_Sys_Write_Log @Proc, @Msg, '', @Cnt

        --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '25'
        set @Msg = '刪除 ['+@SP_date+'] 日轉入的調整單據明細。'
        
        set @strSQL ='delete TYEIPDBS2.lytdbta13.dbo.sslpdttmp '+@CR+
                     ' where Convert(Varchar(10), sd_date, 111) = '''+@SP_date+''' '+@CR+
                     '   and sd_class = '''+@SP_class+''' '+@CR+
                     '   and sd_slip_fg = '''+@sp_slip_fg+''' '+@CR+
                     '   and sd_rem like ''%'+@Rm+'%'' '
           
        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

          --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '26'
        set @Msg = '刪除 ['+@SP_date+'] 日轉入的訂單單據主檔。'
        set @strSQL ='delete TYEIPDBS2.lytdbta13.dbo.ssliptmp '+@CR+
                     ' where Convert(Varchar(10), sp_date, 111) = '''+@SP_date+''' '+@CR+
                     '   and sp_class = '''+@SP_class+''' '+@CR+
                     '   and sp_slip_fg = '''+@sp_slip_fg+''' '+@CR+
                     '   and sp_rem like ''%'+@Rm+'%'' '
 
        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

        --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
print '27'
        set @Msg = '新增 ['+@SP_date+'] 調整單據明細。'
        set @strSQL ='insert into TYEIPDBS2.lytdbta13.dbo.sslpdttmp '+@CR+
                     '(sd_class, sd_slip_fg, sd_date, sd_no, sd_ctno, '+@CR+
                     ' sd_skno, sd_name, sd_whno, sd_whno2, sd_qty, '+@CR+
                     ' sd_price, sd_dis, sd_stot, sd_rem, sd_unit, '+@CR+
                     ' sd_unit_fg, sd_ave_p, sd_rate, sd_seqfld, sd_ordno '+@CR+
                     ') '+@CR+
                     'select '''+@sp_class+''' as sd_class, '+@CR+ -- 類別
                     '       '''+@sp_slip_fg+''' sd_slip_fg, '+@CR+ -- 單據種類
                     '       convert(datetime, '''+@sp_date+''') as sd_date, '+@CR+ -- 貨單日期
                     '       '''+@new_orno+''' as sd_no, '+@CR+ -- 貨單編號
                     '       '''' as sd_ctno, '+@CR+ -- 客戶編號
                     ''+@CR+
                     '       sk_no as sd_skno, '+@CR+ -- 貨品編號
                     '       sk_name as sd_name, '+@CR+ -- 品名規格
                     '       isnull(F3, '''') as sd_whno, '+@CR+ -- 倉庫(入)
                     '       '''' as sd_whno2, '+@CR+ -- 倉庫(出)
                     
                     '       isnull(F4, 0) as SD_qty, '+@CR+ -- 數量
                     --'       isnull(wd_qty, 0) as SD_qty, '+@CR+ -- 數量
                     
                     
                     ''+@CR+
                     
                     '       isnull(F5, 0) as SD_PRICE, '+@CR+ -- 價格
                     --'       case when isnull(s_lprice1, 0) = 0 then sk_save else isnull(s_lprice1, 0) end as SD_PRICE, '+@CR+ -- 價格
                     
                     
                     '       0 as SD_DIS, '+@CR+ -- 折讓金額
                     
                     '       Convert(int, isnull(F4, 0)) * Convert(int, isnull(F5, 0)) as sd_stot, '+@CR+ -- 小計
                     --'       Convert(int, isnull(wd_qty, 0)) * Convert(float, isnull(case when isnull(s_lprice1, 0) = 0 then sk_save else isnull(s_lprice1, 0) end, 0)) as sd_stot, '+@CR+ -- 小計
                     
                     '       '''+@Rm+''' as sd_rem, '+@CR+ -- 備註
                     '       sk_unit as SD_UNIT, '+@CR+ -- 單位
                     ' '+@CR+
                     '       0 as sd_unit_fg, '+@CR+ -- 單位旗標

                     '       F5 as sd_ave_p, '+@CR+ -- 成本金額
                     --'       isnull(case when isnull(s_lprice1, 0) = 0 then sk_save else isnull(s_lprice1, 0) end, 0) as sd_ave_p, '+@CR+ -- 成本金額
                     
                     
                     '       1 as sd_rate, '+@CR+ -- 匯率
                     '       row_number() over(partition by '''+@new_orno+''' order by sk_no) as sd_seqfld, '+@CR+ -- 明細序號, 此序號會因為凌越修改而加以變更，所以另存一份到 sd_ordno
                     '       rowid as sd_ordno'+@CR+ -- XLS 明細序號
                     
                     '  from '+@TB_tmp_Name+' m '+@CR+
                     --'       inner join TYEIPDBS2.lytdbAN13.dbo.sstock d '+@CR+
                     '       inner join SYNC_TA13.dbo.sstock d '+@CR+
                     '          on m.F1 collate Chinese_PRC_BIN = d.sk_no collate Chinese_PRC_BIN '+@CR+
                     
                     '        left join '+@CR+
                     '         (select m.wd_no, m.wd_skno, '+@CR+
                     '                 wd_qty=sum(wd_amt0+wd_amt1+wd_amt2+wd_amt3+wd_amt4+wd_amt5+wd_amt6+wd_amt7+wd_amt8+wd_amt9+wd_amt10+wd_amt11+wd_amt12) '+@CR+
                     --'            from TYEIPDBS2.lytdbAN13.dbo.swaredt m'+@CR+
                     '            from SYNC_TA13.dbo.swaredt m'+@CR+
                     '                 inner join '+@CR+
                     '                   (select wd_no, wd_skno, max(wd_yr) as wd_yr '+@CR+
                     --'                      from TYEIPDBS2.lytdbAN13.dbo.swaredt '+@CR+
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
                     
                     ' where m.rowid > 2 '

        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
        

        --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
        -- 新增出貨退回主檔 PKey: sp_no (ASC), sp_slip_fg (ASC)
print '29'
        set @Msg = '新增 ['+@SP_date+'] 調整單據主檔。'
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
                     '       '''' as sp_ctno, '+@CR+ -- 客戶編號
                     '       '''' as sp_ctname, '+@CR+ -- 客戶名稱
                     '       '''' as sp_ctadd2, '+@CR+ -- 送貨地址
                     '       '''+@SP_Sales+''' as sp_sales, '+@CR+ -- 業務員
                     '       '''+@SP_dpno+''' as sp_dpno, '+@CR+ -- 部門編號
                     ' '+@CR+
                     '       '''+@sp_maker+''' as sp_maker, '+@CR+ -- 製單人員
                     '       '''' as sp_conv, '+@CR+ -- 貨運公司
                     '       sum(sd_stot) as sp_tot, '+@CR+ -- 小計
                     
                     '       0 as sp_tax, '+@CR+ -- 營業稅(原)
                     
                     '       1 as sp_dis, '+@CR+ -- 折讓金額(原)
                     ' '+@CR+
                     '       '''' as sp_pay_kd, '+@CR+ -- 售價種類
                     '       ''NT'' as sp_rate_nm, '+@CR+ -- 匯率名稱
                     '       1 as sp_rate, '+@CR+ -- 匯率
                     '       sum(sd_stot) as sp_itot, '+@CR+ -- 發票金額
                     '       1 as sp_inv_kd, '+@CR+ -- 發票類別(=1  三聯式,=2  二聯式,=3  收銀機)
                     ' '+@CR+
                     
                     '       1 as sp_tax_kd, '+@CR+ -- 稅別(=1應稅,=2零稅 )
                     '       1 as sp_invtype, '+@CR+ -- 開立方式(=1未開, =2隨單開立, =3批次開立)
                     '       '''+@Rm+''' as sp_rem, '+@CR+ -- 備註
                     '       count(1) as sp_tal_rec '+@CR+
                     '  from TYEIPDBS2.lytdbta13.dbo.sslpdttmp m '+@CR+
                     '       left join SYNC_TA13.dbo.sstock d '+@CR+
                     '         on m.sd_skno = d.sk_no '+@CR+
                     ' where sd_class = '''+@sp_class+''' '+@CR+
                     '   and sd_slip_fg = '''+@sp_slip_fg+''' '+@CR+
                     '   and sd_date = '''+@sp_date+''' '+@CR+
                     '   and sd_rem like ''%'+@RM+'%'' '+@CR+
                     ' group by sd_class, '+@CR+
                     '       sd_slip_fg, '+@CR+
                     '       sd_date, '+@CR+
                     '       sd_no '

        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

        --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
        -- 作業完成，清除調整單原始轉入資料
print '30'
        set @Cnt = 0
        set @SP_Date  = ''

        set @strSQL = 'Delete [dbo].['+@TB_xls_Name+'] '
        set @Msg = '作業完成，清除調整單原始轉入資料!!'
        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
     end
     else
     begin
print '31'
        set @Msg = @Msg + '...已存在，終止轉入程序。'
        Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
        set @Cnt = @Errcode
     end
print '32'
  end
print '33'
  Return(@Cnt)

end
GO
