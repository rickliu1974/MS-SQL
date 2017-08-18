USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Get_Table_Schema]    Script Date: 07/24/2017 14:44:00 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Get_Table_Schema]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Proc [dbo].[uSP_Sys_Get_Table_Schema](@TableName varchar(50))
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_Sys_Get_Table_Schema
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [�ܧ� Trans_Log �T����ܤ��e]
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  
  Declare @Proc Varchar(50) = 'uSP_Sys_Get_Table_Schema'
  Declare @strSQL Varchar(Max)=''
  Declare @Msg Varchar(Max) =''
  Declare @ErrMsg Varchar(max) =''
  Declare @Cnt Int =0

  set @Cnt = 0
  begin try
    SELECT [���W��]=b.table_catalog+'.'+b.table_schema+'.'+a.TABLE_NAME,
           [���W��]=b.COLUMN_NAME,
           [��춶��]=b.Ordinal_position,
           [��ƫ��A]=b.DATA_TYPE,
           [�̤j����]=b.CHARACTER_MAXIMUM_LENGTH,
           [�w�]��]=b.COLUMN_DEFAULT,
           [���\�ŭ�]=b.IS_NULLABLE,
           [�w��] = b.Collation_Name,
           [���Ƶ�]=(SELECT value 
                         FROM fn_listextendedproperty (NULL, 'schema', 'dbo', 'table', a.TABLE_NAME, 'column', default) 
                        WHERE name='MS_Description'  
                          and objtype='COLUMN'  
                          and objname Collate Chinese_Taiwan_Stroke_CI_AS = b.COLUMN_NAME) 
      FROM INFORMATION_SCHEMA.TABLES  a 
           LEFT JOIN INFORMATION_SCHEMA.COLUMNS b ON ( a.TABLE_NAME=b.TABLE_NAME ) 
     WHERE TABLE_TYPE='BASE TABLE' 
       and a.TABLE_NAME=@TableName 
     Order by b.table_catalog, b.table_schema, b.TABLE_NAME, b.Ordinal_position
  end try
  begin catch 
    set @Cnt = -1
    Set @ErrMsg = @ErrMsg+'.(���~�T��:'+ERROR_MESSAGE()+')'
  end catch
  Exec uSP_Sys_Write_Log @Proc, @Msg, @ErrMsg, @Cnt

end
GO
