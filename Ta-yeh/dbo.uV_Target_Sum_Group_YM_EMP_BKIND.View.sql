USE [DW]
GO
/****** Object:  View [dbo].[uV_Target_Sum_Group_YM_EMP_BKIND]    Script Date: 07/24/2017 14:43:55 ******/
DROP VIEW [dbo].[uV_Target_Sum_Group_YM_EMP_BKIND]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create View [dbo].[uV_Target_Sum_Group_YM_EMP_BKIND]
as
select year, month, ct_sales, sales_name, bkind, sum(amt) as amt, getdate() as imp_date
  from Target_Sales_Cust_Bkind_Amt 
 where 1=1
   and ct_sales <> ''
 group by year, month, ct_sales, sales_name, bkind;
GO
