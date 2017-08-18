USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Get_Cust_Pdate_Range]    Script Date: 07/24/2017 14:43:59 ******/
DROP PROCEDURE [dbo].[uSP_Get_Cust_Pdate_Range]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Procedure [dbo].[uSP_Get_Cust_Pdate_Range]
 @ct_no Varchar(20), @YM Varchar(7),
 @B_pdate Varchar(10) output, @E_pdate Varchar(10) output
as

declare @ct_pdate Int
declare @First_Day Int
declare @sDate Varchar(10)

declare @B_Month_Max_Day Int
declare @E_Month_Max_Day Int

set @sDate = Convert(Varchar(10), Convert(DateTime, @YM+'/01'), 111)

-- 取得 @sDate 當月以及上個月的最大天數
select @ct_pdate = ct_pdate
  from fact_pcust 
 where ct_no = @ct_no

set @B_Month_Max_Day = day(Convert(DateTime, @sDate) -1)

set @E_Month_Max_Day =
       Case 
         when (@ct_pdate in (30, 31)) then day(DateAdd(dd, -1, DateAdd(mm, 1, Convert(DateTime, @sDate, 111))))
         else @ct_pdate
       end

set @E_pdate =
       Case 
         when @E_Month_Max_Day <= @ct_pdate then Convert(Varchar(10), Convert(DateTime, @YM+ '/'+Convert(Varchar(2), @E_Month_Max_Day)), 111)
         else Convert(Varchar(10), Convert(datetime, @sDate, 111) + @ct_pdate , 111)
       End

       
set @First_Day = Day(DateAdd(mm, -1, @E_pdate))

set @B_pdate = 
       Case
         when (@First_Day >= @B_Month_Max_Day) Or (@B_Month_Max_Day = @E_Month_Max_Day) then Convert(Varchar(10), Convert(DateTime, @sDate), 111)
         else Convert(Varchar(10), DateAdd(mm, -1, Convert(DateTime, @E_pdate) +1), 111)
       end
 
print 'ct_no:'+@ct_no+', '+
      'YM:'+@YM+', '+
      'B_Month_day:'+Convert(Varchar(2), @B_Month_Max_Day)+', '+
      'E_Month_day:'+Convert(Varchar(2), @E_Month_Max_Day)+', '+
      'first_day:'+Convert(Varchar(2), @First_Day)+','+
      'ct_pdate:'+Convert(Varchar(3), @ct_pdate)+', '+
      'sDate:'+@sDate+', '+
      'B_pdate:'+@B_pdate+', '+
      'E_pdate:'+@E_pdate
GO
