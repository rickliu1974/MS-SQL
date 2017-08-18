USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_Stock_Error_Reason]    Script Date: 07/24/2017 14:44:00 ******/
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
    end
  
print '4'
    set @strSQL = 'Exec uSP_Sys_Waiting_Table_Lock '''+@TB_head+''''+@CR+
                  'select year, month, '+@CR+
                  '       sk_no, sk_name, '+@CR+
                  '       qty, error_reason, '+@CR+
                  '       imp_date '+@CR+
                  '       into '+@TB_tmp_Name+' '+@CR+
                  '  from '+@TB_xls_Name 
    Exec @Get_Result =uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, @SendMail
    if @Get_Result = @Err_Code Set @Result = @Err_Code
     
    ;With CTE_Q1 as (
      -- 抓取調撥單的表頭備註，並擷取來源單據類別及單來源單據單號
      -- CTE_Q1 有給主要的 Query 進行 OutQuery, 以及 CTE_Q2 做 inner join, 因此寫成共用方式
      select rtrim(sd_no) as sd_no, rtrim(sd_skno) as sd_skno,
             Chg_sp_date_year, Chg_sp_date_month,
             rtrim(sp_rem) as sp_rem,
             case when sp_rem like '%單:%' then substring(sp_rem, 1, charindex(':', sp_rem)-1) else '' end as kind_name,
             case when sp_rem like '%單:%' then substring(sp_rem, charindex(':', sp_rem)+1, 10) else '' end as kind_sp_no
        from fact_sslpdt m
       where 1=1
         and sd_slip_fg = 'B' 
    ), CTE_Q2 as (
       select sd_no, kind_name, kind_sp_no, code_end as sp_slip_fg, m.Chg_sp_date_year, m.Chg_sp_date_month, sd_skno
         from CTE_Q1 m
              left join Ori_Xls#sys_code d
                on d.code_class = '8'
               and m.kind_name = d.code_name collate Chinese_Taiwan_Stroke_CI_AS
               and m.sp_rem like '%單:%'
    ), CTE_Q3 as (
       select m.*,
              case when rtrim(d.ct_sname) is not null then rtrim(d.ct_sname) else '錯單:'+m.sd_no end as ct_sname
         from CTE_Q2 m
              left join fact_sslpdt d
                on m.kind_sp_no = d.sd_no collate Chinese_Taiwan_Stroke_CI_AS
               and m.sp_slip_fg = d.sd_slip_fg collate Chinese_Taiwan_Stroke_CI_AS
    ), CTE_Q4 as (
      select Chg_sp_date_Year,
             Chg_sp_date_Month,
             sd_skno,
             sum(chg_sd_qty) as sk_cnt,
             sum(Chg_sd_sale_qty) as sk_sale_cnt
        from fact_sslpdt 
       where 1=1 
         and sd_class = '1' 
       group by Chg_sp_date_Year, Chg_sp_date_Month, sd_skno
    ), CTE_Q5 as (
      SELECT a.year, a.month,
             min(Convert(Varchar(4), a.year) + '/' + Substring(Convert(Varchar(3), a.month+100), 2, 2)) + ' ~ '+
             max(Convert(Varchar(4), a.year) + '/' + Substring(Convert(Varchar(3), a.month+100), 2, 2)) as Error_Reason_Rang,
             a.sk_no, c.sk_name,
             sum(a.qty) as error_cnt,
             b.sk_cnt, b.sk_sale_cnt,
             c.Chg_skno_BKind, c.Chg_skno_Bkind_Name,
             c.Chg_skno_SKind, c.Chg_skno_Skind_Name,
             Chg_Wd_AA_first_Qty, Chg_Wd_AA_last_Qty,
             Reverse(Substring(Reverse(IsNull(
               (Select distinct cast(rtrim(error_reason) AS NVARCHAR ) + ',' 
                  from Stock_Error_Reason_tmp 
                 where 1=1
                   and sk_no = a.sk_no 
                   and year = a.year 
                   and month = a.month
                   FOR XML PATH('')
                ) 
             , '')), 2, 1000)) as error_reason,
             Reverse(Substring(Reverse(IsNull(
               (Select distinct isnull(ct_sname, '') + ','
                  from CTE_Q3 m
                 where 1=1
                   and m.sd_skno = a.sk_no collate Chinese_Taiwan_Stroke_CI_AS
                   and m.Chg_sp_date_year = a.year 
                   and m.Chg_sp_date_month = a.month
                   FOR XML PATH('')
               )
             , '')), 2, 1000)) as From_Cust,
            --case when sum(b.sk_cnt) =0 then 0 else sum(a.qty) / sum(b.sk_cnt) end  as error_rate,
            --case when sum(b.sk_sale_cnt) =0 then 0 else sum(a.qty) / sum(b.sk_sale_cnt) end  as error_sale_rate,
            max(imp_date) as imp_date
        FROM Stock_Error_Reason_tmp as a 
             left join CTE_Q4 as b 
               on a.sk_no = b.sd_skno collate Chinese_Taiwan_Stroke_CI_AS
              and a.year = b.Chg_sp_date_Year 
              and a.month = b.Chg_sp_date_Month
             left join Fact_sstock as c
               on a.sk_no = c.sk_no collate Chinese_Taiwan_Stroke_CI_AS
        where 1=1
        group by a.year, a.month,
                 a.sk_no, c.sk_name,
                 b.sk_cnt, b.sk_sale_cnt,
                 c.Chg_skno_BKind, c.Chg_skno_Bkind_Name,
                 c.Chg_skno_SKind, c.Chg_skno_Skind_Name,
                 Chg_Wd_AA_first_Qty, Chg_Wd_AA_last_Qty
    ), CTE_Q6 as (
      select sk_no,
             sum(qty) as TOT_error_cnt,
             TOT_sk_sale_cnt,
             TOT_sk_cnt             
        from Stock_Error_Reason_tmp m
             left join 
               (select sd_skno, 
                       sum(sk_cnt) as TOT_sk_sale_cnt,
                       sum(sk_sale_cnt) as TOT_sk_cnt             
                  from CTE_Q4
                 group by sd_skno
               ) d
               on m.sk_no = d.sd_skno collate Chinese_Taiwan_Stroke_CI_AS
        group by sk_no, TOT_sk_sale_cnt, TOT_sk_cnt
   )
    
    select a.year, a.month,
           Convert(Varchar(4), a.year) + '/' + Substring(Convert(Varchar(3), a.month+100), 2, 2)  as year_month,
           Error_Reason_Rang,
           a.sk_no, a.sk_name,
           a.Chg_skno_BKind, a.Chg_skno_Bkind_Name, a.Chg_skno_SKind, a.Chg_skno_Skind_Name,
           --AA期初存量、AA期末存量
           a.Chg_Wd_AA_first_Qty, a.Chg_Wd_AA_last_Qty,
           --不良原因、不良來源
           a.error_reason, a.From_Cust,

           --本期不良總量
           a.error_cnt, 
           --本期銷售數量
           a.sk_sale_cnt,
           --本期銷售不良率
           case when a.sk_sale_cnt =0 then 0 else a.error_cnt / a.sk_sale_cnt end as error_sale_rate,
           --本期實銷數量
           a.sk_cnt, 
           --本期實銷不良率
           case when a.sk_cnt =0 then 0 else a.error_cnt / a.sk_cnt end as error_rate,

           --累計不良總量
           b.TOT_error_cnt, 
           --累計銷售數量
           b.TOT_sk_sale_cnt,
           --累計銷售不良率
           case when b.TOT_sk_sale_cnt =0 then 0 else b.TOT_error_cnt / b.TOT_sk_sale_cnt end as TOT_error_sale_rate,
           --累計實銷數量
           b.TOT_sk_cnt,
           --累計實銷不良率
           case when b.TOT_sk_cnt =0 then 0 else b.TOT_error_cnt / b.TOT_sk_cnt end as TOT_error_rate,

           a.imp_date, exec_date=getdate()
           Into Stock_Error_Reason
      from CTE_Q5 as a
           left join CTE_Q6 b
             on a.sk_no = b.sk_no collate Chinese_Taiwan_Stroke_CI_AS
     where 1=1
     order by a.year, a.month,
              a.sk_no, a.sk_name,
              a.error_cnt, a.sk_cnt, a.sk_sale_cnt,
              a.Chg_skno_BKind, a.Chg_skno_Bkind_Name,
              a.Chg_skno_SKind, a.Chg_skno_Skind_Name,
              a.Chg_Wd_AA_first_Qty, a.Chg_Wd_AA_last_Qty,
              a.error_reason,
              a.From_Cust,
              a.imp_date


    
    
/*
    --- 2017/08/08 Rickliu 以下此段是以遞迴方式寫作，但會因為不良品資料的年月受限，會無法真實呈現累計銷售及實銷數量因此改寫。
    select a.year, a.month,
           Convert(Varchar(4), a.year) + '/' + Substring(Convert(Varchar(3), a.month+100), 2, 2)  as year_month,
           a.sk_no, a.sk_name,
           a.Chg_skno_BKind, a.Chg_skno_Bkind_Name, a.Chg_skno_SKind, a.Chg_skno_Skind_Name,
           --AA期初存量、AA期末存量
           a.Chg_Wd_AA_first_Qty, a.Chg_Wd_AA_last_Qty,
           --不良原因、不良來源
           a.error_reason, a.From_Cust,

           --本期不良總量
           a.error_cnt, 
           --本期銷售數量
           a.sk_sale_cnt,
           --本期銷售不良率
           case when sum(a.sk_sale_cnt) =0 then 0 else sum(a.error_cnt) / sum(a.sk_sale_cnt) end as error_sale_rate,
           --本期實銷數量
           a.sk_cnt, 
           --本期實銷不良率
           case when sum(a.sk_cnt) =0 then 0 else sum(a.error_cnt) / sum(a.sk_cnt) end as error_rate,

           --累計不良總量
           sum(b.error_cnt) as TOT_error_cnt, 
           --累計銷售數量
           sum(b.sk_sale_cnt) as TOT_sk_sale_cnt,
           --累計銷售不良率
           case when sum(b.sk_sale_cnt) =0 then 0 else sum(b.error_cnt) / sum(b.sk_sale_cnt) end as TOT_error_sale_rate,
           --累計實銷數量
           sum(b.sk_cnt) as TOT_sk_cnt,
           --累計實銷不良率
           case when sum(b.sk_cnt) =0 then 0 else sum(b.error_cnt) / sum(b.sk_cnt) end as TOT_error_rate,

           a.imp_date, exec_date=getdate()
           Into Stock_Error_Reason
      from CTE_Q5 as a, CTE_Q5 as b
     where 1=1
       and Convert(DateTime, Convert(Varchar(4), b.year)+'/'+Convert(Varchar(2), b.month)+'/01') <=
           Convert(DateTime, Convert(Varchar(4), a.year)+'/'+Convert(Varchar(2), a.month)+'/01')
--       and Convert(Varchar(4), b.year) + Substring(Convert(Varchar(3), b.month+100), 2, 2) <=
--           Convert(Varchar(4), a.year) + Substring(Convert(Varchar(3), a.month+100), 2, 2)
       and a.sk_no = b.sk_no 
     group by a.year, a.month,
              a.sk_no, a.sk_name,
              a.error_cnt, a.sk_cnt, a.sk_sale_cnt,
              a.Chg_skno_BKind, a.Chg_skno_Bkind_Name,
              a.Chg_skno_SKind, a.Chg_skno_Skind_Name,
              a.Chg_Wd_AA_first_Qty, a.Chg_Wd_AA_last_Qty,
              a.error_reason,
              a.From_Cust,
              a.imp_date
*/
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
