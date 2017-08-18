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
  Declare @Proce_SubName NVarchar(Max) = '外部檔案匯入至 DW 資料庫程序'
  Declare @BTime DateTime, @ETime DateTime
  Declare @Hour Varchar(2) = ''
  
  declare @TB_tmp_Name Varchar(1000) ='Import_XLS_Schedule'
  declare @strSQL Varchar(max) = ''
  declare @Msg Varchar(max) = '' 
  declare @CurrentDirName Varchar(1000) =''

  set @BTime = GetDate()
  
  set @Hour =(select left(convert(varchar(20), @BTime, 114), 2))

/*  
  if (@Hour in ('12', '00')) --< 設定禁止運行時間 (24 小時制)
  Begin
     Set @body_Msg = @body_Msg + '['+@Proc+'] 已被設定 '+@Hour+' 時禁止運行!!'+@CR
     Goto End_Exit
  end
*/
   
  --[Create Xls Table]
  --判斷是否存在資料表，不存在則新增一個
    IF not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_tmp_Name+']') AND type in (N'U'))
    begin
       Set @Msg = '新增Table ['+@TB_tmp_Name+']。'
       
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
       Set @body_Msg = @body_Msg + '無任何資料不進行轉檔!!'+@CR
       Goto End_Exit
    end
    
    if @Cnt_xls <> @Errcode
    begin
       set @Body_Msg = @Body_Msg + REPLICATE('#', 100)+@CR
       set @Body_Msg = @Body_Msg + '執行 uSP_Imp_xls_to_db【共匯入 Excel 檔案數】...('+CONVERT(Varchar(100), @Cnt_xls)+'個)。'+@CR
       set @Body_Msg = @Body_Msg + REPLICATE('#', 100)+@CR
 
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'EC_Order_Books'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = 'EC01.執行 uSP_Imp_SCM_EC_Order_Books【博客來 EC轉入】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_Books
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '失敗'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1 

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = 'EC02.執行 uSP_Imp_SCM_EC_Order_CrazyMike【瘋狂賣客 EC轉入】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_CrazyMike
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '失敗'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1      

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = 'EC03.執行 uSP_Imp_SCM_EC_Order_ETmall【東森購物 EC轉入】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_ETmall
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '失敗'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = 'EC04.執行 uSP_Imp_SCM_EC_Order_ForMall【瘋MALL EC轉入】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_ForMall
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '失敗'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = 'EC05.執行 uSP_Imp_SCM_EC_Order_Friday【時間軸 EC轉入】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_Friday
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '失敗'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = 'EC06.執行 uSP_Imp_SCM_EC_Order_GoHappy【快樂購 EC轉入】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_GoHappy
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '失敗'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = 'EC07.執行 uSP_Imp_SCM_EC_Order_GoMy【GOMY EC轉入】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_GoMy
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '失敗'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = 'EC08.執行 uSP_Imp_SCM_EC_Order_MOMO【富邦MOMO EC轉入】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_MOMO
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '失敗'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = 'EC09.執行 uSP_Imp_SCM_EC_Order_MyFone【台灣大哥大 EC轉入】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_MyFone
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '失敗'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = 'EC10.執行 uSP_Imp_SCM_EC_Order_PayEasy【PayEasy EC轉入】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_PayEasy
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '失敗'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = 'EC11.執行 uSP_Imp_SCM_EC_Order_PCShop【PCHome購物 EC轉入】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_PCShop
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '失敗'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = 'EC12.執行 uSP_Imp_SCM_EC_Order_PCStore【PCHome商店街 EC轉入】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_PCStore
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '失敗'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = 'EC13.執行 uSP_Imp_SCM_EC_Order_RELO【利樂 EC轉入】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_RELO
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '失敗'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = 'EC14.執行 uSP_Imp_SCM_EC_Order_UDN【UDN EC轉入】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_UDN
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '失敗'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = 'EC15.執行 uSP_Imp_SCM_EC_Order_Umall【森森購物 EC轉入】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_Umall
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '失敗'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1  

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = 'EC16.執行 uSP_Imp_SCM_EC_Order_YahooBuy【Yahoo購物 EC轉入】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_YahooBuy
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '失敗'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1    

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = 'EC17.執行 uSP_Imp_SCM_EC_Order_SuperMarket【Yahoo超級商城 EC轉入】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_SuperMarket
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '失敗'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1     

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = 'EC18.執行 uSP_Imp_SCM_EC_Order_SharedFormat【共用格式轉入】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Order_SharedFormat
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '失敗'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1    

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = '01.執行 uSP_Imp_SCM_Mall_Consign_Order_Car1【車麗屋-寄倉託售回貨轉檔】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Mall_Consign_Order_Car1
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch 
              set @Run_Msg = '失敗'
              set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
             set @Run_Proc_Msg = '02.執行 uSP_Imp_SCM_EC_Consign_Order_PChome【PCHome-寄倉託售回貨轉檔】...'
             Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Consign_Order_PChome
             if @Run_Err = @Errcode
             begin
                set @Run_Msg = '失敗'
                set @Cnt_Err = @Cnt_Err +1
             end
             else
             begin
               set @Run_Msg = '成功'
               set @Cnt_Suss = @Cnt_Suss +1
             end
           End Try
           Begin Catch
                set @Run_Msg = '失敗'
                set @Cnt_Err = @Cnt_Err +1
           End Catch
       
           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
             set @Run_Proc_Msg = '03.執行 uSP_Imp_SCM_EC_Consign_Order_Yahoo【Yahoo-託售回貨轉檔】...'
             Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Consign_Order_Yahoo
             if @Run_Err = @Errcode
             begin
                set @Run_Msg = '失敗'
                set @Cnt_Err = @Cnt_Err +1
             end
             else
             begin
               set @Run_Msg = '成功'
               set @Cnt_Suss = @Cnt_Suss +1
             end
           End Try
           Begin Catch
                set @Run_Msg = '失敗'
                set @Cnt_Err = @Cnt_Err +1
           End Catch
         
           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = '04.執行 uSP_Imp_SCM_Comp_Mall_Consign_Stock_Car1【車麗屋 寄倉庫存對帳轉檔】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Comp_Mall_Consign_Stock_Car1
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '失敗'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = '05.執行 uSP_Imp_SCM_Comp_EC_Consign_Order_Yahoo【Yahoo 寄倉庫存訂單對帳轉檔】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Comp_EC_Consign_Order_Yahoo
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '失敗'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = '06.執行 uSP_Imp_SCM_Comp_EC_Consign_Order_MyFone【台灣大哥大 寄倉訂單對帳轉檔】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Comp_EC_Consign_Order_MyFone
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '失敗'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = '07.執行 uSP_Imp_SCM_Comp_EC_Order_SuperMarket【超級商城 寄倉對帳對帳轉檔】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Comp_EC_Order_SuperMarket
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '失敗'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
              Set @strSQL = ' Update '+ @TB_tmp_Name+@CR+
                            '    set xlsFileName = '''','+@CR+
                            '        Flag = ''0'','+@CR+
                            '        Update_date = getdate()'+@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''' '
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           end
       end
           
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       --set @CurrentDirName = 'Sale_Order_Momo' --< Rickliu 2017/04/13 因為與 Comp 共用一個 XLS 因此，命名與 SCM_Consignment 不同
       set @CurrentDirName = 'Comp_EC_Consign_Order_Momo'
       
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = '08.執行 uSP_Imp_SCM_Comp_EC_Consign_Order_Momo【富邦 MOMO 寄倉訂單對帳轉檔】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Comp_EC_Consign_Order_Momo
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '失敗'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           --begin    
           --   Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
           --   Set @strSQL = ' Update '+ @TB_tmp_Name + ' '+@CR+
           --                 '    set xlsFileName = '''', '+@CR+
           --                 '        Flag = ''0'', '+@CR+
           --                 '        Update_date = getdate() '+@CR+
           --                 '  where Folder_Name = ''' + @CurrentDirName + ''''
           --   Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
           --end
       end

       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       --set @CurrentDirName = 'Sale_Order_Momo' --< Rickliu 2017/04/13 因為與 Comp 共用一個 XLS 因此，命名與 SCM_Consignment 不同
       set @CurrentDirName = 'EC_Consign_Order_Momo'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name=@CurrentDirName)
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = '09.執行 uSP_Imp_SCM_EC_Consign_Order_Momo【富邦 MOMO 寄倉轉檔】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_EC_Consign_Order_Momo
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '失敗'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = '13.執行 uSP_Imp_Sales_Invoice_To_Receipt_Voucher【銷項發票轉轉帳傳票】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_Sales_Invoice_To_Receipt_Voucher
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '失敗'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = '14.執行 uSP_Imp_SCM_Mall_Normal_Order_Car1【車麗屋-揀貨轉訂單】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Mall_Normal_Order_Car1
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '失敗'
                set @Cnt_Err = @Cnt_Err +1
           End Catch


           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = '15.執行 uSP_Imp_SCM_Mall_Normal_Order_SeCar【旭益-揀貨轉訂單】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Mall_Normal_Order_SeCar
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '失敗'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

	   
           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = '16.執行 uSP_Imp_Stock_Error_Reason【不良品送修表】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_Stock_Error_Reason
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '失敗'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
           --   set @Run_Proc_Msg = '17.執行 SP_Imp_Stock_Error_Reason【不良品送修表】...'
           --   Exec @Run_Err = DW.dbo.uSP_Imp_Stock_Error_Reason
           --   if @Run_Err = @Errcode
           --   begin
           --      set @Run_Msg = '失敗'
           --      set @Cnt_Err = @Cnt_Err +1
           --   end
           --   else
           --   begin
           --     set @Run_Msg = '成功'
           --     set @Cnt_Suss = @Cnt_Suss +1
           --   end
           --End Try
           --Begin Catch
           --     set @Run_Msg = '失敗'
           --     set @Cnt_Err = @Cnt_Err +1
           --End Catch

           --set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           --set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           --Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [共用資料表] 成功'
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
       set @Body_Msg = '無 Excel 資料將不進行轉檔。'
       
---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---

    if @Cnt_txt <> @Errcode
    begin
       set @Body_Msg = @Body_Msg + REPLICATE('#', 100)+@CR
       set @Body_Msg = @Body_Msg + '執行 uSP_Sys_Imp_txt_to_db【共匯入 Text 檔案數】...('+CONVERT(Varchar(100), @Cnt_txt)+'個)。'+@CR
       set @Body_Msg = @Body_Msg + REPLICATE('#', 100)+@CR
  
       ---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---***---
       set @CurrentDirName = 'Mall_Normal_Order_800yaoya'
       SET @Cnt = (select count(1) from Import_XLS_Schedule where flag=1 and Folder_Name in (@CurrentDirName+'_3C',@CurrentDirName+'_Retail'))
       if @Cnt > 0
       begin
           Begin Try
              set @Run_Proc_Msg = '01.執行 uSP_Imp_SCM_Mall_Normal_Order_800yaoya【八百屋-揀貨轉訂單】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Mall_Normal_Order_800yaoya
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '失敗'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
              set @Run_Proc_Msg = '02.執行 uSP_Imp_SCM_Mall_Normal_Order_Carno1【車之輪-揀貨轉訂單】...'
              Exec @Run_Err = DW.dbo.uSP_Imp_SCM_Mall_Normal_Order_Carno1
              if @Run_Err = @Errcode
              begin
                 set @Run_Msg = '失敗'
                 set @Cnt_Err = @Cnt_Err +1
              end
              else
              begin
                set @Run_Msg = '成功'
                set @Cnt_Suss = @Cnt_Suss +1
              end
           End Try
           Begin Catch
                set @Run_Msg = '失敗'
                set @Cnt_Err = @Cnt_Err +1
           End Catch

           set @Run_Proc_Msg = @Run_Proc_Msg + @Run_Msg+'!!'
           set @Body_Msg = @Body_Msg + @Run_Proc_Msg+@CR
           Exec uSP_Sys_Write_Log @Proc, @Run_Proc_Msg, @Run_Proc_Msg, @Run_Err, 1

           begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @CurrentDirName +'] 成功'
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
       set @Body_Msg = '無 Text 資料將不進行轉檔。'

    Set @Run_Msg = '執行 '+@Proce_SubName+'...成功 ('+CONVERT(Varchar(2), @Cnt_Suss)+'), 失敗 ('+CONVERT(Varchar(2), @Cnt_Err)+')'
  end try
  begin catch
    Set @Run_Msg = '執行 '+@Run_Proc_Msg+'...發生嚴重錯誤!!'
    Set @Body_Msg = @Run_Proc_Msg+'錯誤訊息：'+@CR+ERROR_MESSAGE()
    Exec uSP_Sys_Write_Log @Proc, @Run_Msg, @Body_Msg, @ErrCode, 1
  end catch

  End_Exit:
  Set @ETime = GetDate()
  Set @Run_Msg =@Proc+' '+@Proce_SubName+'...執行時間：[ '+CONVERT(Varchar(100), @ETime - @BTime, 114)+' ].'+@CR
  Set @Body_Msg = @Body_Msg + @Run_Msg
  Exec uSP_Sys_Write_Log @Proc, @Run_Msg, @Body_Msg, 0, 1

End
GO
