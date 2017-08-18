USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_MSSQL_Free_Memory]    Script Date: 07/24/2017 14:44:00 ******/
DROP PROCEDURE [dbo].[uSP_Sys_MSSQL_Free_Memory]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create procedure [dbo].[uSP_Sys_MSSQL_Free_Memory]
as
begin
  /*
    
    參考以下網址 
    http://fecbob.pixnet.net/blog/post/39058501-%E9%87%8B%E6%94%BEsql-server%E4%BD%94%E7%94%A8%E7%9A%84%E8%A8%98%E6%86%B6%E9%AB%94
    http://caryhsu.blogspot.tw/2011/11/sql-server-x86x64.html
    
    dbcc memorystatus;
    select * from sys.dm_os_performance_counters where object_name like '%memory%' and counter_name like '%Target%'
    
排程資訊
select
    category = jc.name,
    category_id = jc.category_id,
    job_name = j.name,
    job_enabled = j.enabled,
    last_run_time = cast(js.last_run_date as varchar(10))
        + '-' + cast(js.last_run_time as varchar(10)),
    last_run_duration = js.last_run_duration,
    last_run_status = js.last_run_outcome,
    last_run_msg = js.last_outcome_message
        + cast(nullif(js.last_run_outcome,1) as varchar(2)),
    job_created = j.date_created,
    job_modified = j.date_modified
from msdb.dbo.sysjobs j
    inner join msdb.dbo.sysjobservers js
        on j.job_id = js.job_id
    inner join msdb.dbo.syscategories jc
        on j.category_id = jc.category_id
where j.enabled = 1 
    and js.last_run_outcome in (0,1,3,5);
  */

  DBCC FREEPROCCACHE
  WAITFOR DELAY '00:00:03'
  DBCC FREESESSIONCACHE
  WAITFOR DELAY '00:00:03'
  DBCC FREESYSTEMCACHE('All')
  WAITFOR DELAY '00:00:03'
  DBCC DROPCLEANBUFFERS
  WAITFOR DELAY '00:00:03'
  
  -- 打開高級設置配置 EXEC sp_configure 'show advanced options', 1
  EXEC sp_configure 'show advanced options', 1
  RECONFIGURE WITH OVERRIDE
  -- 先設置實體記憶體上限到1G EXEC sp_configure 'max server memory (MB)', 1024
  EXEC sp_configure 'max server memory (MB)', 1024
  RECONFIGURE WITH OVERRIDE
  WAITFOR DELAY '00:00:05'
  -- 還原原先的上限 EXEC sp_configure 'max server memory (MB)', 5120
  EXEC sp_configure 'max server memory (MB)', 13500
  RECONFIGURE WITH OVERRIDE
  WAITFOR DELAY '00:00:05'
  -- 恢復預設配置 EXEC sp_configure 'show advanced options', 0
  EXEC sp_configure 'show advanced options', 0
  RECONFIGURE WITH OVERRIDE
end
GO
