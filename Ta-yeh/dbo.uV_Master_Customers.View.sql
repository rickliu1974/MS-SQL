USE [DW]
GO
/****** Object:  View [dbo].[uV_Master_Customers]    Script Date: 08/18/2017 17:43:39 ******/
DROP VIEW [dbo].[uV_Master_Customers]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE View [dbo].[uV_Master_Customers]
as
select distinct Master_Customer=ct_fld3, Cnt=Count(distinct(ct_fld4))
  from Fact_PCust
 where ct_class='1'
   and Chg_ctno_CustKind_CustCity ='¤j«¬«È¤á'
  group by ct_fld3
GO
