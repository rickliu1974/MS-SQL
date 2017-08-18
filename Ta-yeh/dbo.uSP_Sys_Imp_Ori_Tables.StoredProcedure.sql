USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Imp_Ori_Tables]    Script Date: 08/18/2017 17:43:41 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Imp_Ori_Tables]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[uSP_Sys_Imp_Ori_Tables](@Application_Name Varchar(100) = '', @Server_Name Varchar(100) = '', @DataBase_Name Varchar(100) = '', @Table_Name Varchar(100)='')
as
begin  
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_Sys_Imp_Ori_Tables
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
  Declare @Company_Name Varchar(100) = ''
  declare @strSQL Varchar(Max) = ''
  declare @CR Varchar(4) = Char(13)+Char(10)
  declare @Src_DbName Varchar(100) = ''
  declare @Src_TableName Varchar(100) = ''
  declare @Dsc_TableName Varchar(100) = ''
  declare @First_Name Varchar(100) = 'Ori_'
  declare @Split_Name char(1) = '#'
  
  Declare @Proc Varchar(50) = 'uSP_Sys_Imp_Ori_Tables'
  Declare @Cnt Int= 0, @Cnt1 Int = 0
  Declare @RowCnt Int = 0
  Declare @Result Int = 0
  Declare @Get_Result Int = 0
  Declare @Err_Code Int = -1

  Declare @Msg Varchar(Max) = ''
  Declare @CMD Varchar(100) = ''
  Declare @Like Varchar(1) = '%'
  
  if @Table_Name = ''
     set @Like = '%'
  else
     set @Like = ''

  --[Declare Cursor]
  Declare Cur_Tables Cursor local static For
    Select D.Company_Name, D.Server_Name, 
           M.DataBase_Name, M.Table_Name
      from Sys_Ori_Tables M
           inner join Sys_Ori_DataBases D 
              on M.Application_Name=D.Application_Name
             and M.DataBase_Name=D.DataBase_Name
     where D.[Enabled] = 1 -- True
       and D.Application_Name=@Application_Name
       and D.Server_Name Like '%'+@Server_Name+'%'
       and D.DataBase_Name Like '%'+@DataBase_Name+'%'
       and M.Table_Name Like @Table_Name+@Like
       

  Exec uSP_Sys_Drop_All_Ori_Tables_Constraint
  
  --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
  --[Execute Process]
  Set @Msg = 'Ū����@��['+CONVERT(Varchar, @Cnt)+'] �Ӹ�Ʈw�i����.'
  Set @RowCnt = @@CURSOR_ROWS
  Set @Cnt = @RowCnt
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

  Open Cur_Tables 
  Fetch Next From Cur_Tables into @Company_Name, @Server_Name, @DataBase_Name, @Table_Name

  While @@FETCH_STATUS = 0 
  Begin
    --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
    Set @strSQL =''
    Set @Dsc_TableName = @First_Name+@Company_Name+@Split_Name+@Table_Name
    Set @Src_DbName = '['+@Server_Name+'].['+@DataBase_Name+']' 
    Set @Src_TableName = @Src_DbName+'.[dbo].['+@Table_Name+']'
    Set @Msg = '�}�l����פJ'+@Src_TableName+' �� ['+@Dsc_TableName+'].'
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
   
    --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
    -- ���q�]���n���ƶq�ҥH���� Trans_Log ���y�{
    Set @Msg = 'Ū���ӷ���ƪ�'+@Src_TableName+'...'
    Begin Try
      Set @strSQL = 'select Cnt=Count(1) '+@CR+
                    '  from ['+@Server_Name+'].['+@DataBase_Name+'].[dbo].[sysobjects] '+@CR+
                    ' where name = '''+@Table_Name+''''
      Insert into #Cnt Exec (@strSQL)
      Select @Cnt=Cnt from #Cnt
      Set @Msg = @Msg+'���\'
    end Try
    Begin Catch
      Set @Get_Result = @Err_Code
      Set @Msg = @Msg+'���ѡA�N������פJ�@�~ .(���~�T��:'+ERROR_MESSAGE()+')'
    end Catch
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
    
    --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
    -- �զX�ӷ���ƪ����
    Declare @FieldNames Varchar(Max) = ''
    exec @Cnt1 = uSP_Sys_Get_Columns @Src_DbName, @Table_Name, '@Col_Name', ', ', @FieldNames output
   
    If (@Cnt <> 0) Or (@Cnt1 <> 0)
    begin
      -- [Check Table & Trncate Table]
      If OBJECT_ID(@Dsc_TableName) Is Not Null
      begin
         --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
         set @Msg='�M�� ['+@Dsc_TableName+'].'
         Set @strSQL = 'Truncate Table '+@Dsc_TableName
         Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, 0
         if @Get_Result = @Err_Code set @Result = @Err_Code
/*         
         Begin Try
           Set @Msg =''
           
           Insert into #Msg Exec (@strSQL)
           Set @Msg = '�����M�� ['+@Dsc_TableName+'].'
         end Try
         begin Catch
           Set @Cnt = -1
           Set @Msg = '�L�k�M�� ['+@Dsc_TableName+'].(���~�T��:'+ERROR_MESSAGE()+')'
         end Catch
         Exec uSP_Sys_Write_Log @Proc, @Msg, @Cnt
         Set @Msg = 'Sql Code:'+@strSql
         Exec uSP_Sys_Write_Log @Proc, @Msg, @Cnt
*/     
         --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
         -- [Read Table]
         -- ���q�]���n���ƶq�ҥH���� Trans_Log ���y�{

         Set @Msg = 'Ū�������浧�� ['+@Src_TableName+']...'
         Begin Try
           Set @strSQL = 'select Cnt=Count(1) '+@CR+
                         '  from '+@Src_TableName
           Insert into #Cnt Exec (@strSQL)
           Select @RowCnt = Cnt from #Cnt
           Set @Get_Result = @RowCnt
           Set @Msg = @Msg+'���\'
         end Try
         Begin Catch
           Set @Get_Result = @Err_Code
           Set @Msg = @Msg+'����.(���~�T��:'+ERROR_MESSAGE()+')'
         end Catch
         Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Get_Result
      end
      else
      begin
        --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        -- [Insert Table] �إߪŸ�ƪ�
        Set @Msg = '�إ߸�ƪ� ['+@Dsc_TableName+']'
        Set @strSQL = 'select '+@FieldNames+', Getdate() as UPD_'+@Dsc_TableName+@CR+
                      '  into '+@Dsc_TableName+' '+@CR+
                      '  from '+@Src_TableName+' '+@CR+
                      ' where 1=0' 
        Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, 0
      end
/*
      Begin Try
         Begin Try
            --[Create Ori_Table]
            Set @strSQL = 'select '+@FieldNames+' into '+@Dsc_TableName+' '+@CR+
                          '  from '+@Src_TableName+' '+@CR+
                          ' where 1=0'
            Insert into #Msg Exec (@strSQL)
         end Try
         Begin Catch
            Set @Msg = ERROR_MESSAGE()
            Set @Cnt = -1
            Set @Msg = '�L�k�إ߸�ƪ� ['+@Dsc_TableName+'].(���~�T��1:'+@Msg+')'
            Exec uSP_Sys_Write_Log @Proc, @Msg, @Cnt
            Set @Msg = 'Sql Code:'+@strSql
            Exec uSP_Sys_Write_Log @Proc, @Msg, @Cnt
         end Catch
*/            
         --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
         --[Drop Constraint]
         Exec uSP_Sys_Drop_All_Ori_Tables_Constraint @Dsc_TableName, @Get_Result
         if @Get_Result = @Err_Code set @Result = @Err_Code

         --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
         --[The table have IDENTITY Columns]
         --[IDENTITY Column Cannot Drop, So use Set IDENTITY_INSERT Table ON.]
         --if exists (select * from sys.columns where object_id = OBJECT_ID(@Dsc_TableName, 'U') and is_identity = 1)
         if OBJECTPROPERTY(OBJECT_ID(@Dsc_TableName), 'TableHasIdentity') = 1
         begin
            Set @Msg = '�]�w���۰ʼW�C�}��'
            Set @strSQL = 'SET IDENTITY_INSERT '+@Dsc_TableName+' ON '
            Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, 0
            if @Get_Result = @Err_Code set @Result = @Err_Code
         end

         --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
         --[Insert data to Ori_Table]
         Set @Msg = '�פJ��� '+@Src_TableName+' �� ['+@Dsc_TableName+']'
         Set @strSQL = 'Insert into '+@Dsc_TableName+' '+@CR+
                       '('+@FieldNames+', UPD_'+@Dsc_TableName+')'+@CR+
                       'select '+@FieldNames+', GetDate() as UPD_'+@Dsc_TableName+@CR+
                       '  from '+@Src_TableName
         Exec @Get_Result = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, 0
         if @Get_Result = @Err_Code set @Result = @Err_Code


/*         
         Begin Try
            Set @strSQL = 'Insert into '+@Dsc_TableName+' '+@CR+
                          '('+@FieldNames+')'+@CR+
                          'select '+@FieldNames+' '+@CR+
                          '  from '+@Src_TableName
            Insert into #Msg Exec (@strSQL)
         end Try
         Begin Catch
            Set @Msg = ERROR_MESSAGE()
            Set @Cnt = -1
            Set @Msg = '�L�k�פJ��� '+@Src_TableName+' �� ['+@Dsc_TableName+']���զA���פJ.(���~�T��1:'+@Msg+')'
            Exec uSP_Sys_Write_Log @Proc, @Msg, @Cnt
            Set @Msg = 'Sql Code:'+@strSql
            Exec uSP_Sys_Write_Log @Proc, @Msg, @Cnt
            --[The table have IDENTITY Columns]
            --[IDENTITY Column Cannot Drop, So use Set IDENTITY_INSERT Table ON.]
            Set @strSQL = 'SET IDENTITY_INSERT '+@Dsc_TableName+' ON '+@CR+
                          ' '+@CR+
                          'Insert into '+@Dsc_TableName+' '+@CR+
                          '('+@FieldNames+')'+@CR+
                          'select '+@FieldNames+' '+@CR+
                          '  from '+@Src_TableName
            Insert into #Msg Exec (@strSQL)
            Set @Cnt = @RowCnt
            Set @Msg = '���\�פJ���� ['+@Dsc_TableName+'].'
         end Catch
      end Try
      Begin Catch
         Set @Cnt = -1
         Set @Msg = '�L�k�פJ��� '+@Src_TableName+' �� ['+@Dsc_TableName+'].(���~�T��2:'+ERROR_MESSAGE()+')'
         Exec uSP_Sys_Write_Log @Proc, @Msg, @Cnt
         Set @Msg = 'Sql Code:'+@strSql
      end Catch      
      Exec uSP_Sys_Write_Log @Proc, @Msg, @Cnt
*/
    end

    set @strSQL =''
    set @Cnt = 0
    Set @Msg = '��������פJ'+@Src_TableName+' �� ['+@Dsc_TableName+'].'
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

    Fetch Next From Cur_Tables into @Company_Name, @Server_Name, @DataBase_Name, @Table_Name
  end
  Close Cur_Tables
  DEALLOCATE Cur_Tables
end
GO
