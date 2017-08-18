USE [DW]
GO
/****** Object:  View [dbo].[xV_Target_Sum_Group_YM_CT(201708°±¥Î)]    Script Date: 08/18/2017 17:18:53 ******/
DROP VIEW [dbo].[xV_Target_Sum_Group_YM_CT(201708°±¥Î)]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create View [dbo].[xV_Target_Sum_Group_YM_CT(201708°±¥Î)]
as
select year, month, ct_no, ct_name, sum(amt) as amt, getdate() as imp_date
  from Target_Sales_Cust_Bkind_Amt 
 where 1=1
   and ct_sales <> ''
 group by year, month, ct_no, ct_name
GO
