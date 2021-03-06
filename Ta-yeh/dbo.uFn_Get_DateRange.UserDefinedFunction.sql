USE [DW]
GO
/****** Object:  UserDefinedFunction [dbo].[uFn_Get_DateRange]    Script Date: 08/18/2017 17:43:41 ******/
DROP FUNCTION [dbo].[uFn_Get_DateRange]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Function [dbo].[uFn_Get_DateRange](@aDate DateTime, @aKind Int)
Returns Varchar(100)
as
begin
  Declare @Begin_Week Varchar(100)
  Declare @End_Week Varchar(100)
  Declare @Week_No Varchar(100)
  Declare @Week_Range Varchar(100)
  Declare @Return Varchar(100)
  
  select @Return =
        Case @aKind
          When 1 then Convert(Varchar(1000), Dateadd(Day, (Datepart(weekday, @aDate)-2)*-1, @aDate), 111)
          When 2 then Convert(Varchar(100), Dateadd(Day, 6, Dateadd(day, (Datepart(weekday, @aDate)-2)*-1, @aDate)), 111)
          When 3 then Convert(Varchar(100), Datepart(week, @aDate))
          When 4 then Convert(Varchar(100), Dateadd(Day, (Datepart(weekday, @aDate)-2)*-1, @aDate), 111)+'~'+
                      Convert(Varchar(100), Dateadd(Day, 6, Dateadd(day, (Datepart(weekday, @aDate)-2)*-1, @aDate)), 111)
        end
  Return(@Return)
end
GO
