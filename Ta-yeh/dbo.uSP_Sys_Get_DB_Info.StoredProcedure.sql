USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Get_DB_Info]    Script Date: 07/24/2017 14:44:00 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Get_DB_Info]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Procedure [dbo].[uSP_Sys_Get_DB_Info]
as
begin
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[DB_Info]') AND type in (N'U'))
     DROP TABLE [dbo].[DB_Info]

  Create Table DB_Info (
    Server_Name Varchar(100),
    DB_Name Varchar(100),
    File_ID int,
    File_Size float,
    Used_Size float,
    Over_Size float,
    logic_Name Varchar(100),
    file_Name Varchar(1000),
    Query_Time datetime
  )
  Begin Try
    EXEC TYEIPDBS2.master.dbo.sp_MSforeachdb 'USE ?;
         Insert into [vms\vmsop].[dw].dbo.DB_Info
         SELECT @@servername AS ''伺服器名稱'',DB_NAME() AS ''資料庫名稱'',[FileID] AS ''檔案代碼'', 
                [檔案大小_GB] = CONVERT(DECIMAL(12,2),ROUND([size]/128.000/1024,2)), 
                [空間使用大小_GB] = CONVERT(DECIMAL(12,2),ROUND(fileproperty([name],''SpaceUsed'')/128.000/1024,2)), 
                [剩餘空間大小_GB] = CONVERT(DECIMAL(12,2),ROUND(([size]-fileproperty([name],''SpaceUsed''))/128.000/1024,2)), [Name], 
                [FileName],
                CONVERT(DATETIME,GetDate(),112) AS ''資料查詢時間'' 
           FROM dbo.sysfiles;'
  end Try
  begin catch
  end catch


  Begin Try
    EXEC [vms\vmsop].master.dbo.sp_MSforeachdb 'USE ?;
         Insert into [vms\vmsop].[dw].dbo.DB_Info
         SELECT @@servername AS ''伺服器名稱'',DB_NAME() AS ''資料庫名稱'',[FileID] AS ''檔案代碼'', 
                [檔案大小_GB] = CONVERT(DECIMAL(12,2),ROUND([size]/128.000/1024,2)), 
                [空間使用大小_GB] = CONVERT(DECIMAL(12,2),ROUND(fileproperty([name],''SpaceUsed'')/128.000/1024,2)), 
                [剩餘空間大小_GB] = CONVERT(DECIMAL(12,2),ROUND(([size]-fileproperty([name],''SpaceUsed''))/128.000/1024,2)), [Name], 
                [FileName],
                CONVERT(DATETIME,GetDate(),112) AS ''資料查詢時間'' 
           FROM dbo.sysfiles;'
  end Try
  begin catch
  end catch

  Begin Try
    EXEC [vms\vmsdev].master.dbo.sp_MSforeachdb 'USE ?;
         Insert into [vms\vmsop].[dw].dbo.DB_Info
         SELECT @@servername AS ''伺服器名稱'',DB_NAME() AS ''資料庫名稱'',[FileID] AS ''檔案代碼'', 
                [檔案大小_GB] = CONVERT(DECIMAL(12,2),ROUND([size]/128.000/1024,2)), 
                [空間使用大小_GB] = CONVERT(DECIMAL(12,2),ROUND(fileproperty([name],''SpaceUsed'')/128.000/1024,2)), 
                [剩餘空間大小_GB] = CONVERT(DECIMAL(12,2),ROUND(([size]-fileproperty([name],''SpaceUsed''))/128.000/1024,2)), [Name], 
                [FileName],
                CONVERT(DATETIME,GetDate(),112) AS ''資料查詢時間'' 
           FROM dbo.sysfiles;'
  end Try
  begin catch
  end catch

  select Server_Name as N'伺服器名稱',
         DB_Name as N'資料庫名稱',
         File_ID as N'檔案代碼',
         File_Size as N'檔案大小_GB',
         Used_Size as N'空間使用大小_GB',
         Over_Size as N'剩餘空間大小_GB',
         logic_Name as N'邏輯名稱',
         file_Name as N'檔案名稱',
         Query_Time as N'查詢時間'
    from DB_Info
   order by server_name, db_name, file_id

end
GO
