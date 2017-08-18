USE [DW]
GO
/****** Object:  View [dbo].[xV_Sale_Weeks]    Script Date: 07/24/2017 14:43:55 ******/
DROP VIEW [dbo].[xV_Sale_Weeks]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create View [dbo].[xV_Sale_Weeks]
as
  select distinct
         weekno=Datepart(week, sd_date), 
         Begin_YM=Replace(Convert(Varchar(7), Convert(Varchar(100), Convert(Date, Dateadd(Day, (Datepart(weekday, sd_date)-2)*-1, sd_date))), 112), '-', '/'),
         Begin_Week=Convert(Varchar(100), Convert(Date, Dateadd(Day, (Datepart(weekday, sd_date)-2)*-1, sd_date))),
         End_YM=Replace(Convert(Varchar(7), Convert(Varchar(100), Convert(Date, Dateadd(day, 6, Dateadd(day, (Datepart(weekday, sd_date)-2)*-1, sd_date)))), 112), '-', '/'),
         End_Week=Convert(Varchar(100), Convert(Date, Dateadd(day, 6, Dateadd(day, (Datepart(weekday, sd_date)-2)*-1, sd_date)))),
         week_Range=Convert(Varchar(100), Convert(Date, Dateadd(Day, (Datepart(weekday, sd_date)-2)*-1, sd_date)))+'~'+
                  Convert(Varchar(100), Convert(Date, Dateadd(day, 6, Dateadd(day, (Datepart(weekday, sd_date)-2)*-1, sd_date))))
    from ori_ta13#sslpdt 
   where Datepart(week, sd_date) < 53
GO
