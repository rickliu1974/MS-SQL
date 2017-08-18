USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Get_DB_Size]    Script Date: 08/18/2017 17:43:41 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Get_DB_Size]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Proc [dbo].[uSP_Sys_Get_DB_Size] @Kind int, @str Varchar(100)= '', @Result decimal = 0 Output
as
begin
   if @Kind in (1, 2, 3)
   begin
     Declare @Tempdb_TotSize decimal 
     Declare @Tempdb_UsedSpace decimal
     Declare @Tempdb_FreeSpace decimal 

     SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED
   
     SELECT @Tempdb_TotSize =
             SUM(user_object_reserved_page_count +
                 internal_object_reserved_page_count +
                 version_store_reserved_page_count +
                 mixed_extent_page_count +
                 unallocated_extent_page_count
                ) * (8.0/1024), -- [TotalSizeOfTempDB(MB)]
            @Tempdb_UsedSpace =
             SUM(user_object_reserved_page_count +
                internal_object_reserved_page_count + 
                version_store_reserved_page_count +
                mixed_extent_page_count
                ) * (8.0/1024), --[UsedSpace (MB)]
            @Tempdb_FreeSpace =
             SUM(unallocated_extent_page_count * (8.0/1024)) -- [FreeSpace (MB)]
      FROM sys.dm_db_file_space_usage
      if @Kind = 1 Set @Result = @Tempdb_TotSize
      if @Kind = 2 Set @Result = @Tempdb_UsedSpace
      if @Kind = 3 Set @Result = @Tempdb_FreeSpace
   end
   
   if @Kind in (4, 5, 6) And RTrim(Isnull(@str, '')) <> ''
   begin
      Declare @Table_Used_Size decimal
      Declare @Table_Reserved_Size decimal
      Declare @Table_Row_Count decimal
      
      SELECT -- o.Name table_name, 
             @Table_Used_Size = p.used_page_count * 8 , 
             @Table_Reserved_Size = p.reserved_page_count * 8 , 
             @Table_Row_Count = p.row_count
        FROM sys.dm_db_partition_stats p
             INNER JOIN sys.objects AS o  ON o.object_id = p.object_id
       WHERE o.type_desc = 'USER_TABLE' 
         AND o.is_ms_shipped = 0
         And Upper(o.Name) = Upper(@str)

      if @Kind = 4 Set @Result = @Table_Used_Size
      if @Kind = 5 Set @Result = @Table_Reserved_Size
      if @Kind = 6 Set @Result = @Table_Row_Count
   end
   RETURN @Result
end
GO
