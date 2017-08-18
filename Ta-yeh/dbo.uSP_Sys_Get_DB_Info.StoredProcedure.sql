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
         SELECT @@servername AS ''���A���W��'',DB_NAME() AS ''��Ʈw�W��'',[FileID] AS ''�ɮץN�X'', 
                [�ɮפj�p_GB] = CONVERT(DECIMAL(12,2),ROUND([size]/128.000/1024,2)), 
                [�Ŷ��ϥΤj�p_GB] = CONVERT(DECIMAL(12,2),ROUND(fileproperty([name],''SpaceUsed'')/128.000/1024,2)), 
                [�Ѿl�Ŷ��j�p_GB] = CONVERT(DECIMAL(12,2),ROUND(([size]-fileproperty([name],''SpaceUsed''))/128.000/1024,2)), [Name], 
                [FileName],
                CONVERT(DATETIME,GetDate(),112) AS ''��Ƭd�߮ɶ�'' 
           FROM dbo.sysfiles;'
  end Try
  begin catch
  end catch


  Begin Try
    EXEC [vms\vmsop].master.dbo.sp_MSforeachdb 'USE ?;
         Insert into [vms\vmsop].[dw].dbo.DB_Info
         SELECT @@servername AS ''���A���W��'',DB_NAME() AS ''��Ʈw�W��'',[FileID] AS ''�ɮץN�X'', 
                [�ɮפj�p_GB] = CONVERT(DECIMAL(12,2),ROUND([size]/128.000/1024,2)), 
                [�Ŷ��ϥΤj�p_GB] = CONVERT(DECIMAL(12,2),ROUND(fileproperty([name],''SpaceUsed'')/128.000/1024,2)), 
                [�Ѿl�Ŷ��j�p_GB] = CONVERT(DECIMAL(12,2),ROUND(([size]-fileproperty([name],''SpaceUsed''))/128.000/1024,2)), [Name], 
                [FileName],
                CONVERT(DATETIME,GetDate(),112) AS ''��Ƭd�߮ɶ�'' 
           FROM dbo.sysfiles;'
  end Try
  begin catch
  end catch

  Begin Try
    EXEC [vms\vmsdev].master.dbo.sp_MSforeachdb 'USE ?;
         Insert into [vms\vmsop].[dw].dbo.DB_Info
         SELECT @@servername AS ''���A���W��'',DB_NAME() AS ''��Ʈw�W��'',[FileID] AS ''�ɮץN�X'', 
                [�ɮפj�p_GB] = CONVERT(DECIMAL(12,2),ROUND([size]/128.000/1024,2)), 
                [�Ŷ��ϥΤj�p_GB] = CONVERT(DECIMAL(12,2),ROUND(fileproperty([name],''SpaceUsed'')/128.000/1024,2)), 
                [�Ѿl�Ŷ��j�p_GB] = CONVERT(DECIMAL(12,2),ROUND(([size]-fileproperty([name],''SpaceUsed''))/128.000/1024,2)), [Name], 
                [FileName],
                CONVERT(DATETIME,GetDate(),112) AS ''��Ƭd�߮ɶ�'' 
           FROM dbo.sysfiles;'
  end Try
  begin catch
  end catch

  select Server_Name as N'���A���W��',
         DB_Name as N'��Ʈw�W��',
         File_ID as N'�ɮץN�X',
         File_Size as N'�ɮפj�p_GB',
         Used_Size as N'�Ŷ��ϥΤj�p_GB',
         Over_Size as N'�Ѿl�Ŷ��j�p_GB',
         logic_Name as N'�޿�W��',
         file_Name as N'�ɮצW��',
         Query_Time as N'�d�߮ɶ�'
    from DB_Info
   order by server_name, db_name, file_id

end
GO
