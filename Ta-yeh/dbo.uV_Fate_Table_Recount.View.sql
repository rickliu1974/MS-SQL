USE [DW]
GO
/****** Object:  View [dbo].[uV_Fate_Table_Recount]    Script Date: 08/18/2017 17:18:53 ******/
DROP VIEW [dbo].[uV_Fate_Table_Recount]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create View [dbo].[uV_Fate_Table_Recount]
as

With DB_Fact_TB_Info as
(
select 'DW' as DB_Name, 'sslip' as TB_Name, Count(1) as Cnt from Dw.dbo.Fact_sslip
union
select 'DW' as DB_Name, 'sslpdt' as TB_Name, Count(1) as Cnt from Dw.dbo.Fact_sslpdt
union
select 'DW' as DB_Name, 'sstock' as TB_Name, Count(1) as Cnt from Dw.dbo.Fact_sstock
union
select 'DW' as DB_Name, 'pemploy' as TB_Name, Count(1) as Cnt from Dw.dbo.Fact_pemploy
union
select 'DW' as DB_Name, 'pcust' as TB_Name, Count(1) as Cnt from Dw.dbo.Fact_pcust
union
select 'DW' as DB_Name, 'swaredt' as TB_Name, Count(1) as Cnt from Dw.dbo.Fact_swaredt
), DB_SYNC_TA13_TB_Info as (
select 'SYNC_TA13' as DB_Name, 'sslip' as TB_Name, Count(1) as Cnt from SYNC_TA13.dbo.sslip
union
select 'SYNC_TA13' as DB_Name, 'sslpdt' as TB_Name, Count(1) as Cnt from SYNC_TA13.dbo.sslpdt
union
select 'SYNC_TA13' as DB_Name, 'sstock' as TB_Name, Count(1) as Cnt from SYNC_TA13.dbo.sstock
union
select 'SYNC_TA13' as DB_Name, 'pemploy' as TB_Name, Count(1) as Cnt from SYNC_TA13.dbo.pemploy
union
select 'SYNC_TA13' as DB_Name, 'pcust' as TB_Name, Count(1) as Cnt from SYNC_TA13.dbo.pcust
union
select 'SYNC_TA13' as DB_Name, 'swaredt' as TB_Name, Count(1) as Cnt from SYNC_TA13.dbo.swaredt
), DB_LYTDBTA13_TB_Info as (
select 'LYTDBTA13' as DB_Name, 'sslip' as TB_Name, Count(1) as Cnt from TYEIPDBS2.LYTDBTA13.dbo.sslip
union
select 'LYTDBTA13' as DB_Name, 'sslpdt' as TB_Name, Count(1) as Cnt from TYEIPDBS2.LYTDBTA13.dbo.sslpdt
union
select 'LYTDBTA13' as DB_Name, 'sstock' as TB_Name, Count(1) as Cnt from TYEIPDBS2.LYTDBTA13.dbo.sstock
union
select 'LYTDBTA13' as DB_Name, 'pemploy' as TB_Name, Count(1) as Cnt from TYEIPDBS2.LYTDBTA13.dbo.pemploy
union
select 'LYTDBTA13' as DB_Name, 'pcust' as TB_Name, Count(1) as Cnt from TYEIPDBS2.LYTDBTA13.dbo.pcust
union
select 'LYTDBTA13' as DB_Name, 'swaredt' as TB_Name, Count(1) as Cnt from TYEIPDBS2.LYTDBTA13.dbo.swaredt
)

select A.DB_Name+'.'+A.TB_Name as TB_Name_A1, A.Cnt as A1_Cnt, 
       B.DB_Name+'.'+B.TB_Name as TB_Name_B, B.Cnt as B_Cnt,
       A.Cnt - B.Cnt as 'Diff(A-B)', '*' as s1,
       
       A.DB_Name+'.'+A.TB_Name as TB_Name_A2, A.Cnt as A2_Cnt, 
       C.DB_Name+'.'+C.TB_Name as TB_Name_C, C.Cnt as C_Cnt,
       A.Cnt - C.Cnt as 'Diff(A-C)', '*' as s2,
       getdate() as Query_Time
  from DB_LYTDBTA13_TB_Info A
       inner join DB_SYNC_TA13_TB_Info B
          on A.TB_Name = B.TB_Name
       inner join DB_Fact_TB_Info C
          on A.TB_Name = C.TB_Name
GO
