USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Drop_All_Ori_Tables_Constraint]    Script Date: 08/18/2017 17:43:41 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Drop_All_Ori_Tables_Constraint]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Procedure [dbo].[uSP_Sys_Drop_All_Ori_Tables_Constraint](@TableName Varchar(100) = '', @Result Int = 0)
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_Sys_Drop_All_Ori_Tables_Constraint
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [原本使用 SP_Write_Log 改為 SP_Exec_SQL]
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  
  --[Declare]
  Declare @Proc Varchar(50) = 'uSP_Sys_Drop_All_Ori_Tables_Constraint'
  Declare @strSql Varchar(Max) =''
  Declare @Cnt Int =0
  Declare @RowCnt Int =0
  Declare @Msg Varchar(Max) =''
  Declare @CMD Varchar(100) =''
  Declare @Like Varchar(1) = ''
  
  if @TableName = ''
     set @Like = '%'
  else
     set @Like = ''
     
  --[Clear Default Constraint]
  Declare Cur_Default Cursor local static for
    -- @strSQL 的SQL 命令是直接下 SQL 與法串出來的
    select ConstName=m.name+'.'+d.name, strSql='Alter table '+m.name+' drop CONSTRAINT '+d.name
      from sys.objects m
           inner join sys.objects d on m.object_id=d.parent_object_id
     where d.type = 'D'
       and m.name like @TableName+@Like
     
  Open Cur_Default
  Fetch Next From Cur_Default into @TableName, @strSQL
  Set @RowCnt = @@CURSOR_ROWS
  Set @Cnt = @RowCnt
  if @Cnt <> 0 
  begin
     Set @Msg = '查到資料表有['+CONVERT(Varchar, @Cnt)+']個預設值限制，將進行刪除限制.'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  end

  While @@FETCH_STATUS = 0
  begin
    Set @Msg = '刪除資料表預設值限制:['+@TableName+']'
    
    -- 這裡直接去執行 @strSQL 命令
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, 0
    Fetch Next From Cur_Default into @TableName, @strSQL
  end
  Close Cur_Default
  DEALLOCATE Cur_Default
end
GO
