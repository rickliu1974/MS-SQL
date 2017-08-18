USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_Stock_Error_Reason]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_Imp_Stock_Error_Reason]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[uSP_Imp_Stock_Error_Reason]
as
begin
  Declare @Proc Varchar(50) = 'uSP_Imp_Stock_Error_Reason'

  Declare @Cnt Int =0
  Declare @RowCnt Int =0
  Declare @Msg Varchar(4000) =''
  Declare @Err_Code int = -1
  Declare @Get_Result Int = 0
  Declare @Result Int = 0
  Declare @xls_head Varchar(100), @TB_head Varchar(200), @TB_xls_Name Varchar(200), @TB_tmp_name Varchar(200)
  
  Declare @FirstWord Varchar(100) = 'Ori_xls#' --匯入後所產生的資料表前置檔案名稱
  Declare @TB_xls_Stock_Error_Reason Varchar(100) = '' --匯入後所產生的資料表後置檔案名稱
  Declare @strSQL Varchar(Max)
  Declare @strWhere Varchar(Max)
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @RowCount Table (cnt int)
  Declare @SendMail Int = 2

  Set @xls_head = 'Ori_xls#'
  Set @TB_head = 'Stock_Error_Reason'
  Set @TB_xls_Name = @xls_head+@TB_Head -- Ex: Ori_xls#Stock_Error_Reason -- 外部原檔匯入資料
  Set @TB_tmp_Name = @TB_Head+'_'+'tmp' -- Ex: Stock_Error_Reason_tmp -- 臨時轉入處理介面檔

  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_head+']') AND type in (N'U'))
  begin
     Set @Msg = '刪除資料表 ['+@TB_head+']'
     set @strSQL= 'DROP TABLE [dbo].['+@TB_head+']'

     Exec @Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
     if @Result <>  -1 Set @Result = 0
  end

  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_tmp_Name+']') AND type in (N'U'))
  begin
     Set @Msg = '刪除資料表 ['+@TB_tmp_Name+']'
     set @strSQL= 'DROP TABLE [dbo].['+@TB_tmp_Name+']'

     Exec @Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
     if @Result <>  -1 Set @Result = 0
  end

  begin try
    set @strSQL = 'select count(1) as cnt'+@CR+
                  '  from '+@TB_xls_Name
  
    delete @RowCount
    insert into @RowCount exec (@strSQL)
    set @Cnt =(select Cnt from @RowCount)
    set @Msg = @Msg + '...匯入筆數 ['+cast(@Cnt as varchar)+']'

print '2'
    if @RowCnt > 0
    begin  
print '3'  
/* 
       Set @Msg = '.刪除 '+@TB_tmp_Name+' 資料'
       set @strSQL = 'Exec uSP_Sys_Waiting_Table_Lock '''+@TB_tmp_Name+''''+@CR+
                     'Delete From '+@TB_tmp_Name+' M '+@CR+
                     ' Where Exists '+@CR+
                     '       (SELECT YEAR,MONTH '+@CR+
                     '          FROM '+@TB_xls_Name+' D '+@CR+
                     '         WHERE M.YEAR = D.YEAR '+@CR+
                     '           AND M.MONTH = D.MONTH '+@CR+
                     '       )'
       Exec @Get_Result =uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
       if @Get_Result = @Err_Code Set @Result = @Err_Code
*/
    end
print '4'
    set @strSQL = 'Exec uSP_Sys_Waiting_Table_Lock '''+@TB_head+''''+@CR+
                  'select year, month, '+@CR+
                  '       Convert(Varchar(4), year) + ''/'' + Substring(Convert(Varchar(3), month+100), 2, 2)  as ym, '+@CR+
                  '       sk_no, sk_name, '+@CR+
                  '       sum(qty) as qty, '+@CR+
                  '       Reverse(Substring(Reverse(IsNull( '+@CR+
                  '         (Select distinct cast(rtrim(d.error_reason) AS NVARCHAR ) + '','' '+@CR+ 
                  '            from '+@TB_xls_Name+' d '+@CR+
                  '           where 1=1 '+@CR+
                  '             and d.sk_no = m.sk_no '+@CR+
                  '             and d.year = m.year '+@CR+
                  '             FOR XML PATH('''') '+@CR+
                  '         ) '+@CR+
                  '       , '''')), 2, 1000)) as error_reason, '+@CR+
                  '       imp_date '+@CR+
                  '       into '+@TB_tmp_Name+' '+@CR+
                  '  from '+@TB_xls_Name+' m '+@CR+
                  ' group by year, month, sk_no, sk_name, imp_date '
    Exec @Get_Result =uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
    if @Get_Result = @Err_Code Set @Result = @Err_Code
    
--Return(@Result)
     
    ;With CTE_Q0 as (
      select distinct sk_no
        from Stock_Error_Reason_tmp
    ),CTE_Q1 as (
      -- 抓取調撥單的表頭備註，並擷取來源單據類別及單來源單據單號
      -- CTE_Q1 有給主要的 Query 進行 OutQuery, 以及 CTE_Q2 做 inner join, 因此寫成共用方式
      select rtrim(sd_no) as sd_no, rtrim(sd_skno) as sd_skno,
             chg_sp_date_ym as sp_ym,
             rtrim(sp_rem) as sp_rem,
             case when sp_rem like '%單:%' then substring(sp_rem, 1, charindex(':', sp_rem)-1) else '' end as kind_name,
             case when sp_rem like '%單:%' then substring(sp_rem, charindex(':', sp_rem)+1, 10) else '' end as kind_sp_no
        from fact_sslpdt m
       where 1=1
         and sd_slip_fg = 'B' 
         and exists
             (select * from CTE_Q0 where sk_no = sd_skno collate Chinese_Taiwan_Stroke_CI_AS)
    ), CTE_Q2 as (
       select m.sp_ym, code_end as sp_slip_fg, sd_no, kind_name, kind_sp_no, sd_skno
         from CTE_Q1 m
              left join Ori_Xls#sys_code d
                on d.code_class = '8'
               and m.kind_name = d.code_name collate Chinese_Taiwan_Stroke_CI_AS
               and m.sp_rem like '%單:%'
    ), CTE_Q3 as (
       -- 抓取商品調撥單編號
       select m.*,
              case when rtrim(d.ct_sname) is not null then rtrim(d.ct_sname) else '錯單:'+d.sp_no end as ct_sname
         from CTE_Q2 m
              left join fact_sslip d
                on m.sp_slip_fg = d.sp_slip_fg collate Chinese_Taiwan_Stroke_CI_AS
               -- 原調撥單內的備註若有填寫其他調撥單編號，則以該編號重新抓取調撥單並呈現錯單
               and m.kind_sp_no = d.sp_no collate Chinese_Taiwan_Stroke_CI_AS
    ), CTE_Q4 as (
      -- 各月份銷退貨單加總
      select chg_sp_date_ym as sp_ym,
             sd_skno as sk_no, sd_name as sk_name,
             Chg_skno_BKind, Chg_skno_Bkind_Name,
             Chg_skno_SKind, Chg_skno_Skind_Name,
             Chg_Wd_AA_first_Qty, Chg_Wd_AA_last_Qty,
             sum(Chg_sd_sale_qty) as sd_sale_qty, --銷售數量(無退貨)
             sum(chg_sd_qty) as sd_qty -- 實銷數量
        from fact_sslpdt 
       where 1=1 
         and sd_class = '1' 
         and exists
             (select * from CTE_Q0 where sk_no = sd_skno collate Chinese_Taiwan_Stroke_CI_AS)
         --and sd_skno ='AA020006'
       group by chg_sp_date_ym, sd_skno, sd_name,
             Chg_skno_BKind, Chg_skno_Bkind_Name,
             Chg_skno_SKind, Chg_skno_Skind_Name,
             Chg_Wd_AA_first_Qty, Chg_Wd_AA_last_Qty
    ), CTE_Q5 as (
       -- 不良品資料區間
       select min(Convert(Varchar(4), year) + '/' + Substring(Convert(Varchar(3), month+100), 2, 2)) + ' ~ ' +
              max(Convert(Varchar(4), year) + '/' + Substring(Convert(Varchar(3), month+100), 2, 2)) as Error_Reason_Rang
         from Stock_Error_Reason_tmp
    ), CTE_Q6 as (
      SELECT m.year, m.month, m.ym,
             rtrim(m.sk_no) as sk_no, rtrim(d.sk_name) as sk_name,
             d1.Error_Reason_Rang,
             -- 
             m.qty as error_cnt,
             -- 本期銷售數量
             d.sd_qty, 
             -- 本期實銷數量
             d.sd_sale_qty,
             d.Chg_skno_BKind, d.Chg_skno_Bkind_Name,
             d.Chg_skno_SKind, d.Chg_skno_Skind_Name,
             d.Chg_Wd_AA_first_Qty, d.Chg_Wd_AA_last_Qty,
             m.error_reason,
             Reverse(Substring(Reverse(IsNull(
               (Select distinct isnull(d1.ct_sname, '') + ','
                  from CTE_Q3 d1
                 where 1=1
                   and d1.sd_skno = rtrim(m.sk_no) collate Chinese_Taiwan_Stroke_CI_AS
                   and d1.sp_ym = m.ym
                   FOR XML PATH('')
               )
             , '')), 2, 1000)) as From_Cust,
             m.imp_date
        FROM Stock_Error_Reason_tmp as m
             left join CTE_Q4 as d
               on rtrim(m.sk_no) = rtrim(d.sk_no) collate Chinese_Taiwan_Stroke_CI_AS
              and m.ym = d.sp_ym
             , CTE_Q5 as d1
        where 1=1
        group by m.year, m.month, m.ym,
                 d1.Error_Reason_Rang,
                 rtrim(m.sk_no), rtrim(d.sk_name),
                 m.qty, d.sd_qty, d.sd_sale_qty,
                 d.Chg_skno_BKind, d.Chg_skno_Bkind_Name,
                 d.Chg_skno_SKind, d.Chg_skno_Skind_Name,
                 d.Chg_Wd_AA_first_Qty, d.Chg_Wd_AA_last_Qty,
                 m.error_reason, m.imp_date
    )
    -- 2017/08/10 Rickliu 以資料遞迴方式進行累計統計，PS：資料表遞迴方式，兩表欄位數一定要相同，並將KEY值做 JOIN，然後 A表年月>= B表年月，並將 B表欲累計之欄位進行加總即可。
    select m.year, m.month, m.ym,
           m.Error_Reason_Rang,
           m.sk_no, m.sk_name,
           m.Chg_skno_BKind, m.Chg_skno_Bkind_Name, m.Chg_skno_SKind, m.Chg_skno_Skind_Name,
           --AA期初存量、AA期末存量
           m.Chg_Wd_AA_first_Qty, m.Chg_Wd_AA_last_Qty,
           --不良原因、不良來源
           m.error_reason, m.From_Cust,

           --本期不良總量
           m.error_cnt, 
           --本期銷售數量
           m.sd_sale_qty as sk_sale_cnt,
           --本期銷售不良率
           case when m.sd_sale_qty =0 then 0 else m.error_cnt / m.sd_sale_qty end as error_sale_rate,
           --本期實銷數量
           m.sd_qty as sk_cnt, 
           --本期實銷不良率
           case when m.sd_qty =0 then 0 else m.error_cnt / m.sd_qty end as error_rate,

           --累計不良總量
           sum(isnull(d.error_cnt, 0)) as TOT_error_cnt, 
           --累計銷售數量
           sum(isnull(d.sd_sale_qty, 0)) as TOT_sk_sale_cnt,
           --累計銷售不良率
           case when sum(isnull(d.sd_sale_qty, 0)) =0 then 0 else sum(isnull(d.error_cnt, 0)) / sum(isnull(d.sd_sale_qty, 0)) end as TOT_error_sale_rate,
           --累計實銷數量
           sum(isnull(d.sd_qty, 0)) as TOT_sk_cnt,
           --累計實銷不良率
           case when sum(isnull(d.sd_qty, 0)) =0 then 0 else sum(isnull(d.error_cnt, 0)) / sum(isnull(d.sd_qty, 0)) end as TOT_error_rate,

           max(m.imp_date) as imp_date, exec_date=getdate()
           Into Stock_Error_Reason
      from CTE_Q6 as m
           -- 各月份銷退貨單加總
           inner join CTE_Q6 d
             on m.ym >= d.ym 
            and m.sk_no = d.sk_no collate Chinese_Taiwan_Stroke_CI_AS
            --and m.sk_no = 'AA040059'
     where 1=1
     group by m.year, m.month, m.ym,
              m.Error_Reason_Rang,
              m.sk_no, m.sk_name,
              m.error_cnt, m.sd_sale_qty, m.sd_qty,
              m.Chg_skno_BKind, m.Chg_skno_Bkind_Name,
              m.Chg_skno_SKind, m.Chg_skno_Skind_Name,
              m.Chg_Wd_AA_first_Qty, m.Chg_Wd_AA_last_Qty,
              m.error_reason, m.From_Cust
     select @Cnt = count(1) from Stock_Error_Reason
     
     if @Cnt = 0 
        Set @Result = -1
     else
        Set @Result = @Cnt
        
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  end try
  begin catch
    set @Msg = @Msg+'(錯誤訊息:'+ERROR_MESSAGE()+')'
    set @Result = -1
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Result
  end catch

  Return(@Result)

end
GO
