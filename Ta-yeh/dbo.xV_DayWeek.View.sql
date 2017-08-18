USE [DW]
GO
/****** Object:  View [dbo].[xV_DayWeek]    Script Date: 07/24/2017 14:43:55 ******/
DROP VIEW [dbo].[xV_DayWeek]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create View [dbo].[xV_DayWeek]
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
