USE [DW]
GO
/****** Object:  View [dbo].[xV_Master_Customers]    Script Date: 07/24/2017 14:43:55 ******/
DROP VIEW [dbo].[xV_Master_Customers]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE View [dbo].[xV_Master_Customers]
as
select distinct Master_Customer=ct_fld3, Cnt=Count(distinct(ct_fld4))
  from Fact_PCust
 where ct_class='1'
   and Chg_ctno_CustKind_CustCity ='¤j«¬«È¤á'
  group by ct_fld3
GO
