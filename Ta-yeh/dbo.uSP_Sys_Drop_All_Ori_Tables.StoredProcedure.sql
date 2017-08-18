USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Drop_All_Ori_Tables]    Script Date: 08/18/2017 17:43:41 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Drop_All_Ori_Tables]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[uSP_Sys_Drop_All_Ori_Tables](@Ori_Table Varchar(100) = '')
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_Sys_Drop_All_Ori_Tables
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [原本使用 SP_Write_Log 改為 SP_Exec_SQL]
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  
  Declare @Proc Varchar(50) = 'uSP_Sys_Drop_All_Ori_Tables'
  Declare @strSql Varchar(Max) =''
  Declare @RowCnt Int =0
  Declare @Msg Varchar(Max) =''
  Declare @ErrMsg Varchar(Max) = ''
  
  Declare @CMD Varchar(100) =''
  Declare @TableName Varchar(100) =''
  Declare @Run_Err Int = 0
  Declare @Errcode Int = -1
  Declare @Like Varchar(1) = '%'
  --[Declare Cursor]
  if @Ori_Table = ''
     set @Like = '%'
  else
     set @Like = ''
     
  Declare Cur_OriTables Cursor local static For
    select Quotename(Name, '[') as Name
      from SysObjects 
     where 1=1
       and type = 'U'
       and name like '%#'+@Ori_Table+@Like
       and name not like '%XLS#%'
       and name not like '%TXT#%'

  --[Execute Process]
  Open Cur_OriTables
  Set @RowCnt = @@CURSOR_ROWS
  Set @Run_Err = @RowCnt
  Set @Msg = '預計刪除 ['+CONVERT(Varchar, @RowCnt)+'] 個資料表。'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @RowCnt
  
  Fetch Next From Cur_OriTables into @TableName

  While @@FETCH_STATUS = 0 
  Begin
    Set @Msg = '刪除資料表 '+@TableName+'。'

    set @strSql = 'Drop Table '+@TableName
    Exec @Run_Err = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    if @Run_Err = @Errcode
       set @ErrMsg = @Errcode + @TableName +', '
    Fetch Next From Cur_OriTables into @TableName
  end
  
  if @ErrMsg <> ''
  begin
     set @Run_Err = @Errcode
     set @Msg = @ErrMsg
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Run_Err
  end

  Close Cur_OriTables
  DEALLOCATE Cur_OriTables  
end
GO
