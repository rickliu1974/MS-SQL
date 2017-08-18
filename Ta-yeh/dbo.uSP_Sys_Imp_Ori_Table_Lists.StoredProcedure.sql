USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Imp_Ori_Table_Lists]    Script Date: 08/18/2017 17:43:41 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Imp_Ori_Table_Lists]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[uSP_Sys_Imp_Ori_Table_Lists](@Application_Name Varchar(100) = '')
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_Sys_Imp_Ori_Table_Lists
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [�ܧ� Trans_Log �T����ܤ��e]
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  

  --[Create Temp Table]
  IF OBJECT_ID('tempdb..#Msg') Is Not Null DROP TABLE #Msg
  IF OBJECT_ID('tempdb..#Cnt') Is Not Null DROP TABLE #Cnt
  
  Create Table #Msg(Msg Varchar(Max))
  Create Table #Cnt (Cnt Int)
  --[Declare]
  Declare @DataBase_Name Varchar(100) = ''
  Declare @ServerName Varchar(100) = ''
  Declare @DataBaseName Varchar(100) = ''
  Declare @CompanyName Varchar(100) = ''
  declare @strSQL Varchar(Max) = ''
  declare @CR Varchar(4) = Char(13)+Char(10)
  declare @Src_DbName Varchar(100) = ''
  declare @Src_TableName Varchar(100) = ''
  declare @Dsc_TableName Varchar(100) = ''

  Declare @Proc Varchar(50) = 'uSP_Sys_Imp_Ori_Table_Lists'
  Declare @Cnt Int = 0
  Declare @RowCnt Int = 0
  Declare @Msg Varchar(Max) = ''
  Declare @CMD Varchar(100) = ''
  
  IF OBJECT_ID('Sys_Ori_Tables') Is Null
  begin
     Create Table Sys_Ori_Tables
     (Application_Name Varchar(100),
      Company_Name Varchar(100), 
      Server_Name Varchar(100),
      DataBase_Name Varchar(100), 
      Table_Name Varchar(200),
      Enabled Int)
  end

  Declare Cur_DataBases Cursor local static For
    Select Company_Name, Server_Name, DataBase_Name
      from Sys_Ori_DataBases
     where Enabled = 1 -- True
       and Application_Name=''+@Application_Name+''
     order by DataBase_Name Desc
     
  Set @RowCnt = @@CURSOR_ROWS
  Set @Cnt = @RowCnt
  set @strSQL =''
  Set @Msg = 'Ū����@��['+CONVERT(Varchar, @Cnt)+'] �Ӹ�Ʈw�i����.'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  
  Truncate Table Sys_Ori_Tables

  Open Cur_DataBases 
  Fetch Next From Cur_DataBases into @CompanyName, @ServerName, @DataBaseName

  While @@FETCH_STATUS = 0 
  Begin
    --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
    -- ���q�]���n���ƶq�ҥH���� Trans_Log ���y�{
    begin try
      Set @strSQL = 'select Cnt = COUNT(1) '+@CR+
                    '  from ['+@ServerName+'].['+@DataBaseName+'].[dbo].sysobjects M '+@CR+
                    ' where xtype=''u''  '+@CR+
                    '   and Not Exists '+@CR+
                    '       (select *  '+@CR+
                    '          from Sys_Ori_Tables '+@CR+
                    '         where Application_Name='''+@Application_Name+''' '+@CR+
                    '           and DataBase_Name='''+@DataBaseName+''' '+@CR+
                    '           and Table_name=M.Name) '

      Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
      
      Insert Into #Cnt Exec (@strSQL)
      Select @Cnt = Cnt From #Cnt
      
      Set @strSQL =''
      Set @Msg = '['+@CompanyName+'-'+@ServerName+'-'+@DataBaseName+']...�w�p�פJ['+CONVERT(varchar(100), @Cnt)+'�Ӹ�ƪ�].'
    end try
    begin catch
       Set @Cnt = -1
       Set @Msg = @Msg +'.(���~�T��:'+ERROR_MESSAGE()+')'
    end catch
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

    --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
    Set @Msg = '�פJ[CompanyName:'+@CompanyName+', ServerName:'+@ServerName+', DataBaseName:'+@DataBaseName+']�C'
    Set @strSQL = 'insert into Sys_Ori_Tables '+@CR+
                  'select '''+@Application_Name+''', '''+@CompanyName+''', '''+@ServerName+''', '''+@DataBaseName+''', Upper(name), 1 '+@CR+
                  '  from  ['+@ServerName+'].['+@DataBaseName+'].[dbo].sysobjects M '+@CR+
                  ' where xtype=''u''  '+@CR+
                  '   and Not Exists '+@CR+
                  '       (select *  '+@CR+
                  '          from Sys_Ori_Tables D '+@CR+
                  '         where D.Application_Name='''+@Application_Name+''' '+@CR+
                  '           and DataBase_Name='''+@DataBaseName+''' '+@CR+
                  '           and D.Table_name=M.Name) '+@CR+
                  ' order by name '
 
     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

     Fetch Next From Cur_DataBases into @CompanyName, @ServerName, @DataBaseName
  end
  Close Cur_DataBases
  DEALLOCATE Cur_DataBases
end
GO
