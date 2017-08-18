USE [DW]
GO
/****** Object:  View [dbo].[uV_Target_Sum_Group_YM_CT]    Script Date: 07/24/2017 14:43:55 ******/
DROP VIEW [dbo].[uV_Target_Sum_Group_YM_CT]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create View [dbo].[uV_Target_Sum_Group_YM_CT]
as
select year, month, ct_no, ct_name, sum(amt) as amt, getdate() as imp_date
  from Target_Sales_Cust_Bkind_Amt 
 where 1=1
   and ct_sales <> ''
 group by year, month, ct_no, ct_name
GO
