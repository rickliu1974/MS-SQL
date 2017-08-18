USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[xuSP_ETL_Stock_Error_Reason(201708停用)]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[xuSP_ETL_Stock_Error_Reason(201708停用)]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[xuSP_ETL_Stock_Error_Reason(201708停用)]
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_ETL_Stock_Error_Reason
   Create Date: 2017/06/22
   Creator: Rickliu
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  

  Declare @Proc Varchar(50) = 'uSP_ETL_Stock_Error_Reason'
  Declare @Cnt Int =0
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @Result int = 0
  
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Stock_Error_Reason_Cal]') AND type in (N'U'))
  begin
     Set @Msg = '刪除資料表 [Stock_Error_Reason_Cal]'
     set @strSQL= 'DROP TABLE [dbo].[Stock_Error_Reason_Cal]'

     Exec @Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
     if @Result <>  -1 Set @Result = 0
  end

  set @Msg = 'Create to [Stock_Error_Reason_Cal]...'
  begin try
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
            a.sk_no, c.sk_name,
            sum(a.qty) as error_cnt,
            b.sk_cnt, b.sk_sale_cnt,
            c.Chg_skno_BKind, c.Chg_skno_Bkind_Name,
            c.Chg_skno_SKind, c.Chg_skno_Skind_Name,
            Chg_Wd_AA_first_Qty, Chg_Wd_AA_last_Qty,
            Reverse(Substring(Reverse(IsNull(
              (Select distinct cast(rtrim(error_reason) AS NVARCHAR ) + ',' 
                 from Stock_Error_Reason 
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
           case when sum(sk_cnt) =0 then 0 else sum(a.qty) / sum(sk_cnt) end  as error_rate,
           case when sum(sk_sale_cnt) =0 then 0 else sum(a.qty) / sum(sk_sale_cnt) end  as error_sale_rate
       FROM Stock_Error_Reason as a 
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
    )

    select a.year, a.month,
           Convert(Varchar(4), a.year) + '/' + Substring(Convert(Varchar(3), a.month+100), 2, 2)  as year_month,
           a.sk_no, a.sk_name,
           a.error_cnt, a.sk_cnt, a.sk_sale_cnt,
           a.Chg_skno_BKind, a.Chg_skno_Bkind_Name, a.Chg_skno_SKind, a.Chg_skno_Skind_Name,
           a.Chg_Wd_AA_first_Qty, a.Chg_Wd_AA_last_Qty,
           a.error_reason, a.From_Cust,
           case when sum(a.sk_cnt) =0 then 0 else sum(a.error_cnt) / sum(a.sk_cnt) end as error_rate,
           case when sum(a.sk_sale_cnt) =0 then 0 else sum(a.error_cnt) / sum(a.sk_sale_cnt) end as error_sale_rate,
           sum(b.error_cnt) as TOT_error_cnt, sum(b.sk_cnt) as TOT_sk_cnt,
           case when sum(b.sk_cnt) =0 then 0 else sum(b.error_cnt) / sum(b.sk_cnt) end as TOT_error_rate,
           sum(b.sk_sale_cnt) as TOT_sk_sale_cnt,
           case when sum(b.sk_sale_cnt) =0 then 0 else sum(b.error_cnt) / sum(b.sk_sale_cnt) end as TOT_error_sale_rate
           Into Stock_Error_Reason_Cal
      from CTE_Q5 as a, CTE_Q5 as b
     where 1=1
       and Convert(Varchar(4), b.year) + Substring(Convert(Varchar(3), b.month+100), 2, 2) <=
           Convert(Varchar(4), a.year) + Substring(Convert(Varchar(3), a.month+100), 2, 2)
       and a.sk_no = b.sk_no 
     group by a.year, a.month,
              a.sk_no, a.sk_name,
              a.error_cnt, a.sk_cnt, a.sk_sale_cnt,
              a.Chg_skno_BKind, a.Chg_skno_Bkind_Name,
              a.Chg_skno_SKind, a.Chg_skno_Skind_Name,
              a.Chg_Wd_AA_first_Qty, a.Chg_Wd_AA_last_Qty,
              a.error_reason,
              a.From_Cust
   
 
     select @Cnt = count(1) from Stock_Error_Reason_Cal
     
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
