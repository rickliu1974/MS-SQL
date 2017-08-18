USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Exec_IMP]    Script Date: 08/18/2017 17:43:40 ******/
DROP PROCEDURE [dbo].[uSP_Exec_IMP]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[uSP_Exec_IMP]
as
begin
  Declare @Proc Varchar(50) = 'uSP_Exec_IMP'
  Declare @CR Varchar(4) =' '+Char(13)+Char(10)
  Declare @Body_Msg NVarchar(Max) = '', @Run_Proc_Msg NVarchar(Max) = '', @Run_Msg NVarchar(1000) = ''
  Declare @Cnt_xls int = 0, @Cnt_txt int = 0, @Run_Err int = 0, @Errcode int = -1
  Declare @Cnt_Err int =0,  @Cnt_Suss int = 0
  Declare @Cnt int =0
  Declare @Result int = 0
  Declare @Proce_SubName NVarchar(Max) = '�~���ɮ׶פJ�� DW ��Ʈw�{��'
  Declare @BTime DateTime, @ETime DateTime
  Declare @Hour Varchar(2) = ''
  
  declare @TB_tmp_Name Varchar(1000) ='Import_XLS_Schedule'
  declare @strSQL Varchar(max) = ''
  declare @Msg Varchar(max) = '' 
  declare @CurrentDirName Varchar(1000) =''

  set @BTime = GetDate()
  
  set @Hour =(select left(convert(varchar(20), @BTime, 114), 2))

/*  
  if (@Hour in ('12', '00')) --< �]�w�T��B��ɶ� (24 �p�ɨ�)
  Begin
     Set @body_Msg = @body_Msg + '['+@Proc+'] �w�Q�]�w '+@Hour+' �ɸT��B��!!'+@CR
     Goto End_Exit
  end
*/
   
  --[Create Xls Table]
  --�P�_�O�_�s�b��ƪ�A���s�b�h�s�W�@��
    IF not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_tmp_Name+']') AND type in (N'U'))
    begin
       Set @Msg = '�s�WTable ['+@TB_tmp_Name+']�C'
       
       Set @strSQL = 'CREATE TABLE [dbo].['+@TB_tmp_Name+']('+@CR+
                     '             [UniqueID] [int] IDENTITY(1,1) NOT NULL,'+@CR+
                     '             [Folder_Name] [varchar](100) NULL,'+@CR+
                     '             [xlsFileName] [varchar](100) NULL,'+@CR+
                     '             [Flag] [varchar](1) NULL,'+@CR+
                     '             [import_date] [datetime] NULL,'+@CR+
                     '             [Update_date] [datetime] NULL'+@CR+
                     ') ON [PRIMARY]' 
        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    end

  Begin Try
    Exec @Cnt_xls = DW.dbo.uSP_Sys_Imp_xls_to_db
    Exec @Cnt_txt = DW.dbo.uSP_Sys_Imp_txt_to_db
    set @Cnt_xls = (select count(1) from Import_XLS_Schedule where flag=1)

    set @Cnt_Err = @Cnt_xls + @Cnt_txt
    if (@Cnt_Err in (0, -2))
    begin
       Set @body_Msg = @body_Msg + '�L�����Ƥ��i������!!'+@CR
       Goto End_Exit
    end
    
    if @Cnt_xls <> @Errcode
    begin
       set @Body_Msg = @Body_Msg + REPLICATE('#', 100)+@CR
       set @Body_Msg = @Body_Msg + '���� uSP_Imp_xls_to_db�i�@�פJ Excel �ɮ׼ơj...('+CONVERT(Varchar(100), @Cnt_xls)+'��)�C'+@CR
       set @Body_Msg = @Body_Msg + REPLICATE('#', 100)+@CR
 
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Order_Books'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = 'EC01.���� uSP_Imp_SCM_EC_Order_Books�i�իȨ� EC��J�j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_Books
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '����'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Order_CrazyMike'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = 'EC02.���� uSP_Imp_SCM_EC_Order_CrazyMike�i�ƨg��� EC��J�j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_CrazyMike
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '����'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1      

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Order_ETmall'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = 'EC03.���� uSP_Imp_SCM_EC_Order_ETmall�i�F���ʪ� EC��J�j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_ETmall
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '����'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
        end
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Order_ForMall'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = 'EC04.���� uSP_Imp_SCM_EC_Order_ForMall�i��MALL EC��J�j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_ForMall
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '����'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
        end
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Order_Friday'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = 'EC05.���� uSP_Imp_SCM_EC_Order_Friday�i�ɶ��b EC��J�j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_Friday
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '����'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name + @CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
        end
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Order_GoHappy'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = 'EC06.���� uSP_Imp_SCM_EC_Order_GoHappy�i�ּ��� EC��J�j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_GoHappy
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '����'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name + @CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
        end
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Order_GoMy'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = 'EC07.���� uSP_Imp_SCM_EC_Order_GoMy�iGOMY EC��J�j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_GoMy
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '����'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
        end
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Order_MOMO'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = 'EC08.���� uSP_Imp_SCM_EC_Order_MOMO�i�I��MOMO EC��J�j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_MOMO
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '����'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
        end
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Order_MyFone'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = 'EC09.���� uSP_Imp_SCM_EC_Order_MyFone�i�x�W�j���j EC��J�j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_MyFone
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '����'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
        end
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Order_PayEasy'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = 'EC10.���� uSP_Imp_SCM_EC_Order_PayEasy�iPayEasy EC��J�j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_PayEasy
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '����'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
        end
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Order_PCShop'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = 'EC11.���� uSP_Imp_SCM_EC_Order_PCShop�iPCHome�ʪ� EC��J�j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_PCShop
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '����'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
        end
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Order_PCStore'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = 'EC12.���� uSP_Imp_SCM_EC_Order_PCStore�iPCHome�ө��� EC��J�j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_PCStore
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '����'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Order_RELO'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = 'EC13.���� uSP_Imp_SCM_EC_Order_RELO�i�Q�� EC��J�j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_RELO
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '����'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Order_UDN'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = 'EC14.���� uSP_Imp_SCM_EC_Order_UDN�iUDN EC��J�j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_UDN
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '����'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Order_Umall'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = 'EC15.���� uSP_Imp_SCM_EC_Order_Umall�i�˴��ʪ� EC��J�j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_Umall
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '����'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Order_YahooBuy'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = 'EC16.���� uSP_Imp_SCM_EC_Order_YahooBuy�iYahoo�ʪ� EC��J�j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_YahooBuy
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '����'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1    

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''', '+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Order_SuperMarket'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = 'EC17.���� uSP_Imp_SCM_EC_Order_SuperMarket�iYahoo�W�Űӫ� EC��J�j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_SuperMarket
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '����'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1     

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Order_SharedFormat'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = 'EC18.���� uSP_Imp_SCM_EC_Order_SharedFormat�i�@�ή榡��J�j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_SharedFormat
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '����'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1    

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name + @CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'Mall_Consign_Order_Car1'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name in (@CurrentDirName+'_3C',@CurrentDirName+'_Retail'))
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = '01.���� uSP_Imp_SCM_Mall_Consign_Order_Car1�i���R��-�H�ܰU��^�f���ɡj...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Mall_Consign_Order_Car1
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '����'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name like ''%' + @CurrentDirName + '%'' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end
       
        
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Consign_Order_PChome'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
             set @Run_Proc_Msg = '02.���� uSP_Imp_SCM_EC_Consign_Order_PChome�iPCHome-�H�ܰU��^�f���ɡj...'
             Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Consign_Order_PChome
             if @Run_Err = @Errcode
             begin
                set @Run_Msg = '����'
                set @Cnt_Err = @Cnt_Err +1
             end
             else
             begin
               set @Run_Msg = '���\'
               set @Cnt_Suss = @Cnt_Suss +1
             end
           End Try
           Begin Catch
                set @Run_Msg = '����'
                set @Cnt_Err = @Cnt_Err +1
           End Catch
       
           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end
  
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Consign_Order_Yahoo'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
             set @Run_Proc_Msg = '03.���� uSP_Imp_SCM_EC_Consign_Order_Yahoo�iYahoo-�U��^�f���ɡj...'
             Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Consign_Order_Yahoo
             if @Run_Err = @Errcode
             begin
                set @Run_Msg = '����'
                set @Cnt_Err = @Cnt_Err +1
             end
             else
             begin
               set @Run_Msg = '���\'
               set @Cnt_Suss = @Cnt_Suss +1
             end
           End Try
           Begin Catch
                set @Run_Msg = '����'
                set @Cnt_Err = @Cnt_Err +1
           End Catch
         
           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'Comp_Mall_Consign_Stock_Car1'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name in (@CurrentDirName+'_3C',@CurrentDirName+'_Retail'))
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = '04.���� uSP_Imp_SCM_Comp_Mall_Consign_Stock_Car1�i���R�� �H�ܮw�s��b���ɡj...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Comp_Mall_Consign_Stock_Car1
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '����'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name like ''%' + @CurrentDirName + '%'' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end
  
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'Comp_EC_Consign_Order_Yahoo'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = '05.���� uSP_Imp_SCM_Comp_EC_Consign_Order_Yahoo�iYahoo �H�ܮw�s�q���b���ɡj...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Comp_EC_Consign_Order_Yahoo
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '����'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name + @CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'Comp_EC_Consign_Order_MyFone'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = '06.���� uSP_Imp_SCM_Comp_EC_Consign_Order_MyFone�i�x�W�j���j �H�ܭq���b���ɡj...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Comp_EC_Consign_Order_MyFone
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '����'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end
       
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'Comp_EC_Order_SuperMarket'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = '07.���� uSP_Imp_SCM_Comp_EC_Order_SuperMarket�i�W�Űӫ� �H�ܹ�b��b���ɡj...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Comp_EC_Order_SuperMarket
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '����'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end
           
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       --set @CurrentDirName = 'Sale_Order_Momo' --< Rickliu 2017/04/13 �]���P Comp �@�Τ@�� XLS �]���A�R�W�P SCM_Consignment ���P
       set @CurrentDirName = 'Comp_EC_Consign_Order_Momo'
       
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = '08.���� uSP_Imp_SCM_Comp_EC_Consign_Order_Momo�i�I�� MOMO �H�ܭq���b���ɡj...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Comp_EC_Consign_Order_Momo
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '����'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           --begin    
           --   Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
           --   Set @strSQL = ' Update '+ @TB_tmp_Name + ' '+@CR+
           --                 '    set xlsFileName = '''', '+@CR+
           --                 '        Flag = ''0'', '+@CR+
           --                 '        Update_date = getdate() '+@CR+
           --                 '  where Folder_Name = ''' + @CurrentDirName + ''''
           --   Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           --end
       end

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       --set @CurrentDirName = 'Sale_Order_Momo' --< Rickliu 2017/04/13 �]���P Comp �@�Τ@�� XLS �]���A�R�W�P SCM_Consignment ���P
       set @CurrentDirName = 'EC_Consign_Order_Momo'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = '09.���� uSP_Imp_SCM_EC_Consign_Order_Momo�i�I�� MOMO �H�����ɡj...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Consign_Order_Momo
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '����'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end
       
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'Sales_Invoice_To_Receipt_Voucher'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = '13.���� uSP_Imp_Sales_Invoice_To_Receipt_Voucher�i�P���o������b�ǲ��j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_Sales_Invoice_To_Receipt_Voucher
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '����'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name + @CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'Mall_Normal_Order_Car1'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name in (@CurrentDirName+'_3C',@CurrentDirName+'_Retail'))
       if @Cnt > 0
       begin
	       Begin Try
              set @Run_Proc_Msg = '14.���� uSP_Imp_SCM_Mall_Normal_Order_Car1�i���R��-�z�f��q��j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Mall_Normal_Order_Car1
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '����'
                set @Cnt_Err = @Cnt_Err +1
           End Catch


           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'', '+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name like ''%' + @CurrentDirName + '%'' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'Mall_Normal_Order_SeCar'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name in (@CurrentDirName+'_3C',@CurrentDirName+'_Retail'))
       if @Cnt > 0
       begin
	       Begin Try
              set @Run_Proc_Msg = '15.���� uSP_Imp_SCM_Mall_Normal_Order_SeCar�i���q-�z�f��q��j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Mall_Normal_Order_SeCar
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '����'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

	   
           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name like ''%' + @CurrentDirName + '%'' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end

       
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'Stock_Error_Reason'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = '16.���� uSP_Imp_Stock_Error_Reason�i���}�~�e�ת�j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_Stock_Error_Reason
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '����'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       --set @CurrentDirName = 'Stock_Error_Reason'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name in ('Hunderd_NonSales_Stock','Master_Stock','New_Stock_Lists','Sys_Code','Target_Stock_BKind_Business','Target_Stock_BKind_Personal'))
       if @Cnt > 0
       begin
           --Begin Try
           --   set @Run_Proc_Msg = '17.���� SP_Imp_Stock_Error_Reason�i���}�~�e�ת�j...'
           --   Exec @Run_Err = DW.dbo.uSP_Imp_Stock_Error_Reason
           --   if @Run_Err = @Errcode
           --   begin
           --      set @Run_Msg = '����'
           --      set @Cnt_Err = @Cnt_Err +1
           --   end
           --   else
           --   begin
           --     set @Run_Msg = '���\'
           --     set @Cnt_Suss = @Cnt_Suss +1
           --   end
           --End Try
           --Begin Catch
           --     set @Run_Msg = '����'
           --     set @Cnt_Err = @Cnt_Err +1
           --End Catch

           --set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           --set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           --Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [�@�θ�ƪ�] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name in (''Hunderd_NonSales_Stock'',''Master_Stock'',''New_Stock_Lists'',''Target_Stock_BKind_Business'',''Target_Stock_BKind_Personal'',''Sys_Code'')'
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end
    end
    else
       set @Body_Msg = '�L Excel ��ƱN���i�����ɡC'
       
---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---

    if @Cnt_txt <> @Errcode
    begin
       set @Body_Msg = @Body_Msg + REPLICATE('#', 100)+@CR
       set @Body_Msg = @Body_Msg + '���� uSP_Sys_Imp_txt_to_db�i�@�פJ Text �ɮ׼ơj...('+CONVERT(Varchar(100), @Cnt_txt)+'��)�C'+@CR
       set @Body_Msg = @Body_Msg + REPLICATE('#', 100)+@CR
  
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'Mall_Normal_Order_800yaoya'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name in (@CurrentDirName+'_3C',@CurrentDirName+'_Retail'))
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = '01.���� uSP_Imp_SCM_Mall_Normal_Order_800yaoya�i�K�ʫ�-�z�f��q��j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Mall_Normal_Order_800yaoya
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '����'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name like ''%' + @CurrentDirName + '%'' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end
        
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'Mall_Normal_Order_Carno1'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = '02.���� uSP_Imp_SCM_Mall_Normal_Order_Carno1�i������-�z�f��q��j...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Mall_Normal_Order_Carno1
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '����'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '���\'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '����'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @CurrentDirName +'] ���\'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end
    end
    else
       set @Body_Msg = '�L Text ��ƱN���i�����ɡC'

    Set @Run_Msg = '���� '+@Proce_SubName+'...���\ ('+CONVERT(Varchar(2), @Cnt_Suss)+'), ���� ('+CONVERT(Varchar(2), @Cnt_Err)+')'
  end try
  begin catch
    Set @Run_Msg = '���� '+@Run_Proc_Msg+'...�o���Y�����~!!'
    Set @Body_Msg = @Run_Proc_Msg+'���~�T���G'+@CR+ERROR_MESSAGE()
    Exec uSP_Sys_Write_Log @Proc, @Run_Msg, @Body_Msg, @ErrCode, 1
  end catch

  End_Exit:
  Set @ETime = GetDate()
  Set @Run_Msg =@Proc+' '+@Proce_SubName+'...����ɶ��G[ '+CONVERT(Varchar(100), @ETime - @BTime, 114)+' ].'+@CR
  Set @Body_Msg = @Body_Msg + @Run_Msg
  Exec uSP_Sys_Write_Log @Proc, @Run_Msg, @Body_Msg, 0, 1

End
GO
