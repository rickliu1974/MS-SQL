USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_RunProcess_Info]    Script Date: 08/18/2017 17:43:41 ******/
DROP PROCEDURE [dbo].[uSP_Sys_RunProcess_Info]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Proc [dbo].[uSP_Sys_RunProcess_Info] (@Trans_Date Varchar(10) = '')
as

If IsDate(Convert(DateTime, @Trans_Date)) = 1
   Set @Trans_Date = Convert(Varchar(10), getdate(), 111)

Print @Trans_Date

;With CTE_Log_Split1 as 
(
select Rowid, process, Trans_Date,
       Replace(Replace(substring(Msg, CharIndex('<-', Msg)+2, Len(Msg)), process, ''), 'Exec ...', '') as Msg1, 
       Msg,
       Exec_Time
  from trans_log With (NoWait, NoLock)
 where 1=1
   and Exec_Time Is Not Null 
   and Trans_Date >= @Trans_Date --convert(varchar(10), getdate(), 111)
   and (process = 'uSP_Realtime_SaleData' Or process = 'uSP_Exec_ETL')
),
CTE_Log_Split2 as 
(
select Rowid, process, Trans_Date, 
       substring(Msg1, 1, case when CharIndex('¡i', Msg1) >0 then PATINDEX('%¡i%', Msg1)-1 else Len(Msg1) end) as Msg1, 
       substring(substring(Msg, PATINDEX('%...%', Msg)+3, Len(Msg)), 1, PATINDEX('%...%', substring(Msg, PATINDEX('%...%', Msg)+4, Len(Msg)))) as Msg2,
       Exec_Time
  from CTE_Log_Split1 With (NoWait)
 where Trans_Date >= @Trans_Date--convert(varchar(10), getdate(), 111)
)

select Rowid, process, Trans_Date,
       Msg1 as Msg,
       Convert(Int, Replace(Substring(Msg2, CharIndex(':', Msg2, 1)+1, Len(Msg2)), '¡j', '')) as 'Inc_Tempdb_MB',
       Exec_Time
  from CTE_Log_Split2 With (NoWait)
 where Exec_Time >= '00:00:01'
 order by Exec_Time desc
GO
