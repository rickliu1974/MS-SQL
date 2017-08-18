USE [DW]
GO
/****** Object:  View [dbo].[xV_Master_Customers(201708停用)]    Script Date: 08/18/2017 17:43:39 ******/
DROP VIEW [dbo].[xV_Master_Customers(201708停用)]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE View [dbo].[xV_Master_Customers(201708停用)]
as
select distinct Master_Customer=ct_fld3, Cnt=Count(distinct(ct_fld4))
  from Fact_PCust
 where ct_class='1'
   and Chg_ctno_CustKind_CustCity ='大型客戶'
  group by ct_fld3
GO
