USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Release_Memory]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Release_Memory]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create PROCEDURE [dbo].[uSP_Sys_Release_Memory]
AS 
BEGIN 
  Declare @Proc Varchar(50) = 'uSP_Sys_Release_Memory'
  Declare @Proc_Name NVarchar(Max) = '釋放資料庫記憶體程序'

  Declare @CR Varchar(4) =Char(13)+Char(10)
  Declare @Cnt Int = 0;
  Declare @Release_Range Float, @Memory_Used_Range Float, @Min_Memory Float, @Max_Memory Float, @Before_Memory_Used Float, @After_Memory_Used Float
  Declare @Value NVarchar(100)
  Declare @sMsg Varchar(max)
  
  Set @Min_Memory = '1024'
  Set @Release_Range = 60 -- 記憶體使用量達 60% 以上，則才要進行記憶體釋放作業
  -- 2017/08/15 Rickliu 取得資料庫原先設定記憶體最大值，後續好還原
  Set @Max_memory = (select convert(Int, value) from sys.configurations where name like '%max server memory%')
  
  /*
  Select (physical_memory_in_use_kb/1024) as Memory_usedby_Sqlserver_MB
         (locked_page_allocations_kb/1024) as Locked_pages_used_Sqlserver_MB
         (total_virtual_address_space_kb/1024) as Total_VAS_in_MB
         process_physical_memory_low,
         process_virtual_memory_low
    FROM sys.dm_os_process_memory;
  */
 
  Set @Before_Memory_Used = (Select (physical_memory_in_use_kb/1024) as Memory_usedby_Sqlserver_MB From sys.dm_os_process_memory)
  Set @Memory_Used_Range = Round((Select @Before_Memory_Used / @Max_Memory)*100, 2)
  if @Memory_Used_Range >= @Release_Range
  begin
     DBCC FREEPROCCACHE
     DBCC FREESESSIONCACHE
     DBCC FREESYSTEMCACHE('All')
     DBCC DROPCLEANBUFFERS

     EXEC sys.sp_configure N'show advanced options', N'1'   
     RECONFIGURE WITH OVERRIDE 

     set @Value = Convert(NVarchar(100), @Min_Memory)
     EXEC sys.sp_configure N'max server memory (MB)', @Value
     RECONFIGURE WITH OVERRIDE 
     EXEC sys.sp_configure N'show advanced options', N'0' 
     RECONFIGURE WITH OVERRIDE 

     WAITFOR DELAY '00:00:30' 

     EXEC sys.sp_configure N'show advanced options', N'1'   
     RECONFIGURE WITH OVERRIDE 
     set @Value = Convert(NVarchar(100), @Max_Memory)
     EXEC sys.sp_configure N'max server memory (MB)', @Value
     RECONFIGURE WITH OVERRIDE 
     EXEC sys.sp_configure N'show advanced options', N'0' 
     RECONFIGURE WITH OVERRIDE 
  end
  
  Set @After_Memory_Used = (Select (physical_memory_in_use_kb/1024) as Memory_usedby_Sqlserver_MB From sys.dm_os_process_memory)
  
  Select getdate() as Release_Time,
         @Min_Memory as 'Set_Min_Memory(MB)',
         @Max_Memory as 'Set_Max_Memory(MB)',
         @Release_Range as 'Set_Memory_Range(%)',
         @Memory_Used_Range as 'Memory_Used_Range(%)',
         @Before_Memory_Used as 'Before_Memory_Used(MB)',
         @After_Memory_Used as 'After_Memory_Used(MB)',
         @Before_Memory_Used - @After_Memory_Used as 'Release_Memory(MB)'

print @Memory_Used_Range
  set @sMsg = '執行'+ @Proc + ' '+@Proc_Name+' 記憶體已使用 ['+Convert(Varchar(100), @Memory_Used_Range)+'%]'
  if @Memory_Used_Range >= @Release_Range
  begin
     set @sMsg = @sMsg+ 
                 ',釋放前後 ['+Convert(Varchar(100), @Before_Memory_Used)+'(MB) → '+Convert(Varchar(100), @After_Memory_Used)+'(MB)]'+
                 ',已釋放 '+Convert(Varchar(100), @Before_Memory_Used - @After_Memory_Used)+'(MB)'
     Exec uSP_Sys_Write_Log @Proc, @sMsg, '', 0, 1
  end
  else
     set @sMsg = @sMsg+ ',未達釋放標準 '+Convert(Varchar(100), @Release_Range)+'%, 目前記憶體 '+Convert(Varchar(100), @After_Memory_Used)+'(MB).'
  
print '-------------------------------------------------------------------------------------------------------------------------------------------'
print 'Mail Body:'+@sMsg
print '-------------------------------------------------------------------------------------------------------------------------------------------'
  
END
GO
