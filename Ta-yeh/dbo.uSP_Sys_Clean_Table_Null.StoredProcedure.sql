USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Clean_Table_Null]    Script Date: 07/24/2017 14:44:00 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Clean_Table_Null]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create proc [dbo].[uSP_Sys_Clean_Table_Null] (@TableName Varchar(Max)) 
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_Sys_Clean_Table_Null
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [原本使用 SP_Write_Log 改為 SP_Exec_SQL]
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
  declare @Proc Varchar(50) = 'uSP_Sys_Clean_Table_Null' -- [Procedure Name]

  declare @Column_Name Varchar(100)
  declare @Data_Type Varchar(100)
  declare @Column_Default Varchar(100)
  declare @Is_Nullable Varchar(100)
  declare @Value Varchar(Max)
  declare @strSQL Varchar(Max)
  declare @Cnt_Table Table (cnt int)
  declare @Get_Result int =0
  Declare @Err_Code Int = -1
  Declare @Result Int = 0
  Declare @Msg Varchar(Max) =''

  declare cur_schema cursor for
     SELECT b.COLUMN_NAME               ,--as 欄位名稱     
            b.DATA_TYPE                 ,--as 資料型態 
            b.COLUMN_DEFAULT            ,--as 預設值 
            b.IS_NULLABLE                --as 允許空值 
       FROM INFORMATION_SCHEMA.TABLES  a 
            LEFT JOIN INFORMATION_SCHEMA.COLUMNS b ON ( a.TABLE_NAME=b.TABLE_NAME ) 
      WHERE TABLE_TYPE='BASE TABLE' 
        and a.TABLE_NAME=@TableName 
        and b.IS_NULLABLE = 'YES'
    
  open cur_schema
  fetch cur_schema into @Column_Name, @Data_Type, @Column_Default, @Is_Nullable

  While @@FETCH_STATUS = 0
  begin
    set @Msg='清除資料表 ['+@TableName+'.'+@Column_Name+' 型態:('+@Data_Type+')] 內所有 NULL 值.'
    set @strSQL = ''

    if (Upper(@Data_Type) like '%CHAR%') And (Upper(@Is_Nullable) = 'Yes')
    begin
       Set @Value = ''
       Set @strSQL = '   SET ['+@Column_Name+']='''+@Value+''''
    end
    else if ((Upper(@Data_Type) like '%INT%') Or (Upper(@Data_Type) in ('FLOAT', 'BIT', 'DECIMAL'))) And (Upper(@Is_Nullable) = 'Yes')
    begin
       Set @Value = '0'
       Set @strSQL = '   SET ['+@Column_Name+']='+@Value
    end

    if @strSQL <> ''
    begin
       set @strSQL = 'UPDATE '+@TableName+' '+@strSQL +'  where RTrim(IsNull(Convert(Varchar(1000), ['+@Column_Name+']), '''')) ='''' '
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       if @Get_Result = @Err_Code set @Result = @Err_Code
    end 
    fetch Next From cur_schema into @Column_Name, @Data_Type, @Column_Default, @Is_Nullable
  end
  Close cur_schema
  DEALLOCATE cur_schema
end
GO
