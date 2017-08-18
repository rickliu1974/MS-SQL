USE [DW]
GO
/****** Object:  View [dbo].[xV_Sales_Base]    Script Date: 07/24/2017 14:43:55 ******/
DROP VIEW [dbo].[xV_Sales_Base]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create View [dbo].[xV_Sales_Base] as
select *
  from (select bo_class, bo_no, bo_yr,  bo_month='1', bo_sale=bo_sale1 from SYNC_TA13.dbo.sbase
         union
        select bo_class, bo_no, bo_yr,  bo_month='2', bo_sale=bo_sale2 from SYNC_TA13.dbo.sbase
         union
        select bo_class, bo_no, bo_yr,  bo_month='3', bo_sale=bo_sale3 from SYNC_TA13.dbo.sbase
         union
        select bo_class, bo_no, bo_yr,  bo_month='4', bo_sale=bo_sale4 from SYNC_TA13.dbo.sbase
         union
        select bo_class, bo_no, bo_yr,  bo_month='5', bo_sale=bo_sale5 from SYNC_TA13.dbo.sbase
         union
        select bo_class, bo_no, bo_yr,  bo_month='6', bo_sale=bo_sale6 from SYNC_TA13.dbo.sbase
         union
        select bo_class, bo_no, bo_yr,  bo_month='7', bo_sale=bo_sale7 from SYNC_TA13.dbo.sbase
         union
        select bo_class, bo_no, bo_yr,  bo_month='8', bo_sale=bo_sale8 from SYNC_TA13.dbo.sbase
         union
        select bo_class, bo_no, bo_yr,  bo_month='9', bo_sale=bo_sale9 from SYNC_TA13.dbo.sbase
         union
        select bo_class, bo_no, bo_yr,  bo_month='10', bo_sale=bo_sale10 from SYNC_TA13.dbo.sbase
         union
        select bo_class, bo_no, bo_yr,  bo_month='11', bo_sale=bo_sale11 from SYNC_TA13.dbo.sbase
         union
        select bo_class, bo_no, bo_yr,  bo_month='12', bo_sale=bo_sale12 from SYNC_TA13.dbo.sbase
       ) m
 where bo_sale <> 0
GO
