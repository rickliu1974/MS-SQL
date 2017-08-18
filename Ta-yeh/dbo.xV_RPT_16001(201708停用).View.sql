USE [DW]
GO
/****** Object:  View [dbo].[xV_RPT_16001(201708°±¥Î)]    Script Date: 08/18/2017 17:43:39 ******/
DROP VIEW [dbo].[xV_RPT_16001(201708°±¥Î)]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create View [dbo].[xV_RPT_16001(201708°±¥Î)]
as


SELECT a.year
      ,a.month
      ,a.sk_no
      ,c.sk_name
      ,sum(a.qty) as error_cnt
      ,b.sk_cnt
	  ,b.sk_sale_cnt
      ,c.Chg_skno_BKind
      ,c.Chg_skno_Bkind_Name
      ,c.Chg_skno_SKind
      ,c.Chg_skno_Skind_Name
      ,Chg_Wd_AA_first_Qty
      ,Chg_Wd_AA_last_Qty
      ,(SELECT distinct cast(rtrim(error_reason) AS NVARCHAR ) + ',' 
	      from Stock_Error_Reason 
         where sk_no = a.sk_no and year = a.year and month = a.month
        FOR XML PATH('')) as error_reason
      ,case when sum(sk_cnt) =0 then 0 else  sum(a.qty) / sum(sk_cnt) end  as error_rate
      ,case when sum(sk_sale_cnt) =0 then 0 else  sum(a.qty) / sum(sk_sale_cnt) end  as error_sale_rate
  FROM Stock_Error_Reason as a left join (select sum(chg_sd_qty) as sk_cnt,sum(Chg_sd_sale_qty) as sk_sale_cnt,Chg_sp_date_Year,Chg_sp_date_Month,sd_skno 
                                            from fact_sslpdt 
										   where 1=1 
										     and sd_class = '1' 
										   group by Chg_sp_date_Year,Chg_sp_date_Month,sd_skno )as b 
                                          on a.sk_no = b.sd_skno collate Chinese_Taiwan_Stroke_CI_AS and a.year = b.Chg_sp_date_Year and a.month = b.Chg_sp_date_Month
                               left join Fact_sstock as c on a.sk_no = c.sk_no collate Chinese_Taiwan_Stroke_CI_AS
  where 1=1

  --and a.sk_no = 'AA010024'
  group by a.year 
          ,a.month 
	      ,a.sk_no 
	      ,c.sk_name 
	      ,b.sk_cnt
	      ,b.sk_sale_cnt
		  ,c.Chg_skno_BKind
          ,c.Chg_skno_Bkind_Name
          ,c.Chg_skno_SKind
          ,c.Chg_skno_Skind_Name
          ,Chg_Wd_AA_first_Qty
          ,Chg_Wd_AA_last_Qty
GO
