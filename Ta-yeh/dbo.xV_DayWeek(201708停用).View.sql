USE [DW]
GO
/****** Object:  View [dbo].[xV_DayWeek(201708停用)]    Script Date: 08/18/2017 17:18:53 ******/
DROP VIEW [dbo].[xV_DayWeek(201708停用)]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create View [dbo].[xV_DayWeek(201708停用)]
as
  select theday=1, Week_name='日'
   union
  select theday=2, Week_name='一'
   union
  select theday=3, Week_name='二'
   union
  select theday=4, Week_name='三'
   union
  select theday=5, Week_name='四'
   union
  select theday=6, Week_name='五'
   union
  select theday=7, Week_name='六'
GO
