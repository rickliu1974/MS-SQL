USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Imp_Txt_to_db]    Script Date: 07/24/2017 14:44:00 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Imp_Txt_to_db]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[uSP_Sys_Imp_Txt_to_db]
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_Sys_Imp_Txt_to_db
   Create Date: 2013/11/24
   Creator: Rickliu
   
   -- 目錄名稱末字若有以下符號則代表檔案處理方式：
   
   PS: 由於使用 BULK 方式匯入，所以統一開一個欄位，並將資料全部匯入。
   @: 目錄內的每一個 Text 都轉成一個 Table 並以檔案名稱底線前方為 TableName, Ex: FileName MOMO_20141218.xls => TableName 為 Ori_TXT#MOMO.

  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  
  --[Enabled advance options]
  Exec uSP_Sys_Advanced_Options 1 
  
  --[Declare]
  declare @Proc Varchar(50) = 'uSP_Sys_Imp_Txt_to_db' -- [Procedure Name]

  declare @ImportRoot Varchar (1000) = 'D:\Transform_Data'
  declare @txtRoot Varchar (1000) = @ImportRoot + '\Import_Txt_To_DB' -- [Txt root path]
  declare @HistoryRoot Varchar(1000) = @ImportRoot + '\History_Import' -- [History Import path]
  declare @CurrentPath Varchar(1000) = @TxtRoot
  Declare @Errcode int = -1

  declare @txtSubDir table (SubDirName varchar(255)) --目錄列表
  declare @SubDirName Varchar(255) = '' --目錄名稱
  
  declare @txtFiles table (txtFileName varchar(255)) --檔案列表
  declare @txtFileName Varchar(255) = '' --檔案名稱
  declare @FileName Varchar(255) --檔案名稱不含附檔案名稱
  declare @SplitFileName Varchar(255) --檔案名稱分隔符號後名之名稱
  declare @RowCount table (cnt int)
  
  declare @FirstWord Varchar(100) = 'Ori_Txt#' --匯入後所產生的資料表前置檔案名稱
  declare @LastWord Varchar(100) = '' --匯入後所產生的資料表後置檔案名稱
  declare @CurrentTable Varchar(100) =''
  declare @CurrentDirName Varchar(100) = '' -- 取得當前的目錄名稱
  declare @HistoryTable Varchar(100) =''

  declare @OldbP1 Varchar(100) = 'WITH (FIELDTERMINATOR='' '', ROWTERMINATOR=''\n'') '
  
  declare @HaveFileName Varchar(100) = '' --把目錄名稱當作匯入後的其中一個欄位
  
  declare @strSQL Varchar(max) = ''
  declare @LastSQL Varchar(max) = ''
  declare @Cnt Int = 0
  declare @Cnt_Suss Int = 0
  declare @Msg Varchar(max) = '' 
  declare @Cmd Varchar(2000) = ''
  declare @CR Varchar(4) = Char(13)+Char(10)
  declare @Fields Varchar(max) = ''
  declare @ErrCode_Msg Varchar(max) = ''
  declare @imp_date Varchar(20)
  declare @Columns Varchar(1000)
  
  declare @TB_tmp_Name Varchar(1000) ='Import_XLS_Schedule'
  
  --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
  --[Test Driver can work.]
  -- 此處有重大訊息顯示，所以不另外修改成 Exec uSP_Exec_SQL 語法
  
  Set @Msg = 'Text 引擎啟動測試.'
  Set @Cnt = 0

  begin try
    set @CurrentTable ='Test_tmp'
    set @txtFileName ='測試檔案(勿刪).txt'
    
    if @CurrentPath like '%@%'
       set @FileName = Substring(@txtFileName, 1, CHARINDEX('.', @txtFileName)-1)
    else
       set @FileName = ''

    IF OBJECT_ID(@CurrentTable) Is Not Null 
    begin
       set @Msg ='刪除測試資料表 ['+@CurrentTable+']'
       set @strSQL='Drop TABLE '+@CurrentTable
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    end
 
    set @Msg ='建立測試資料表 ['+@CurrentTable+']'
    set @strSQL='CREATE TABLE '+@CurrentTable+' (F1 VARCHAR(1000)) '
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

    set @Msg = '匯入測試資料表 ['+@CurrentTable+']'
    Set @strSQL = 'BULK INSERT '+@CurrentTable+' FROM '''+@ImportRoot+'\'+@txtFileName+''' '+ @OldbP1

    set @ErrCode_Msg = 'SqlCode-1.1:'+@strSQL

    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

    
    Set @Cnt = 0
    Set @Msg = @Msg+'成功.'
    
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  end try
  begin catch
    Set @Cnt = @Errcode
    Set @Msg = @Msg+'失敗，終止所有轉入動作，請重新啟動 SQL Server.(錯誤訊息：'+ERROR_MESSAGE()+')'
    Exec uSP_Sys_Write_Log @Proc, @Msg, @Cnt
    
    goto Exception
  end catch
  
  
  --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

  
  --[Create Xls Table]
  --判斷是否存在資料表，不存在則新增一個
    IF not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_tmp_Name+']') AND type in (N'U'))
    begin
       Set @Msg = '新增Table ['+@TB_tmp_Name+']。'
       
       Set @strSQL = 'CREATE TABLE [dbo].['+@TB_tmp_Name+'](             ' +@CR+
                     '             [UniqueID] [int] IDENTITY(1,1) NOT NULL,' +@CR+
                     '             [Folder_Name] [varchar](100) NULL,         ' +@CR+
                     '             [xlsFileName] [varchar](100) NULL,            ' +@CR+
                     '             [Flag] [varchar](1) NULL,            ' +@CR+
                     '             [import_date] [datetime] NULL,              ' +@CR+
                     '             [Update_date] [datetime] NULL              ' +@CR+
                     ') ON [PRIMARY]                                       ' 
        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL


    end





  --[Set Values]
  Set @strSQL =''
  Set @Msg = '匯入開始...'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  
  Set @Msg = '取得外部 Text 目錄路徑 ['+@HistoryRoot+']'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  
  insert into @txtSubDir execute xp_subdirs @txtRoot
  delete @txtSubDir where SubDirName is null
  
  select @Cnt = COUNT(1) from @txtSubDir
  Set @Msg = '共取得 ['+Convert(Varchar(100), @Cnt)+'] 目錄.'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
    
  declare Cur_SubDirs Cursor local static For
    select * From @txtSubDir
  -- 子目錄循序讀取
  Open Cur_SubDirs -- 子目錄循序讀取
  Fetch Next From Cur_SubDirs into @SubDirName
  
  While @@FETCH_STATUS = 0 
  Begin
    set @CurrentPath = @txtRoot +'\'+@SubDirName
    set @CurrentDirName = Replace(Replace(@SubDirName, '@', ''), '#', '')
    set @HistoryTable = @CurrentDirName+'_History'
    set @CurrentTable = @FirstWord+@CurrentDirName
    
    set @cnt = (select count(1) from Import_XLS_Schedule where Folder_Name = @SubDirName)
    if @cnt = 0
    begin
        Set @strSQL = 'Insert into '+ @TB_tmp_Name + '(Folder_Name,xlsFileName,Flag,import_date,Update_date) ' +@CR+
                    '     values (''' + @CurrentDirName +''','''',''0'',getdate() ,'''')'
        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    end

    set @Cmd = 'dir '+@CurrentPath + ' *.txt *.prn /b'
    -- Clear Variable
    delete @txtFiles
    -- Get Import files
    insert into @txtFiles Execute xp_cmdshell @Cmd
 
    -- Clear Error Message
    delete @txtFiles where txtFileName is null or txtFileName like '%找不到%' or txtFileName like '%Found%'
    select @Cnt = COUNT(1) from @txtFiles
    Set @Msg = '讀取 ['+@CurrentPath+'] 目錄, 共取得 ['+CONVERT(Varchar(100), @Cnt)+'] 個檔案.'
    set @strSQL = ''
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

    -- Found Files
    if @Cnt <> 0
    begin
       if @SubDirName not like '%@%'
       begin
          set @Msg = '清除資料表 ['+@CurrentTable+'].'
          set @strSQL='DROP TABLE ['+@CurrentTable+']'
          IF OBJECT_ID(@CurrentTable) Is Not Null 
             -- Exec (@strSQL)
             Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       end

       Declare Cur_txtFiles Cursor local static For
         Select * from @txtFiles
       -- 目錄內的檔案循序讀取
       Open Cur_txtFiles
       Fetch Next From Cur_txtFiles into @txtFileName
       While @@FETCH_STATUS = 0
       begin
          begin try
            --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
            set @imp_date = getdate()
            set @FileName = Convert(nVarchar(255), +Substring(@txtFileName, 1, CHARINDEX('.', @txtFileName)-1))
            set @SplitFileName = Convert(nVarchar(255), +Substring(@FileName, CHARINDEX('_', @FileName)+1, Len(@FileName)))

            set @Columns = ', xlsFileName=Convert(Varchar(255), '''+@FileName+'''), SplitFileName =Convert(Varchar(255), '''+@SplitFileName+'''), imp_date=getdate()'


            -- 2014/12/18 Rickliu 目錄名稱若有 @ 符號則代表，每一個 Text 檔案都另行產生一個 Table。
            if @SubDirName like '%@%'
            begin
               set @CurrentTable = @FirstWord+
                                   @CurrentDirName+'_'+
                                   Convert(nVarchar(255), +Substring(@FileName, 1, CHARINDEX('_', @FileName)-1))
                                   
               set @Msg = '清除資料表 ['+@CurrentTable+'].'
               set @strSQL='DROP TABLE '+@CurrentTable
               IF OBJECT_ID(@CurrentTable) Is Not Null 
                  --Exec (@strSQL)
                  Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
            end

            IF OBJECT_ID(@CurrentTable) Is Null 
            begin
               set @Msg ='建立資料表 ['+@CurrentTable+']'
               set @strSQL='CREATE TABLE ['+@CurrentTable+'] (F1 VARCHAR(1000)) '

               Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
            end
    
            Set @strSQL = 'BULK INSERT ['+@CurrentTable+'] FROM '''+@CurrentPath+'\'+@txtFileName+''' '+ @OldbP1
    
            Set @Msg = '準備匯入 ['+@CurrentPath+'\'+@txtFileName+'] 至 ['+@CurrentTable+']資料表.'
            Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

            if @SubDirName like '%@%'
            begin
               -- 增加 rowid 建立 Text 內容每一筆的唯一鍵值
               Set @Msg = '產生 Rowid 於 ['+@CurrentTable+'] 資料表.'
               set @Cnt =0
               Set @strSQL ='alter table ['+@CurrentTable+'] add rowid int identity(1, 1)'
               Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

               --set @strSQL = 'alter table ['+@HistoryTable+'] add imp_date datetime '
               --Exec uSP_Exec_SQL @Proc, @Msg, @strSQL

               --set @strSQL = 'update ['+@HistoryTable+'] set imp_date=getdate() '
               --Exec uSP_Exec_SQL @Proc, @Msg, @strSQL

               --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
               -- 清除所有欄位皆含 NULL 的資料
               set @Msg='清除所有欄位皆含 NULL 的資料'
               exec uSP_Sys_Get_Columns '', @CurrentTable, '(@Col_Name is null)', ' and ', @Fields output 
               Set @Fields = REPLACE(@Fields, 'AND ([TXTFILENAME] IS NULL)', '')
               --2013/5/3 排除 RowID
               --Set @Fields = REPLACE(@Fields, 'AND ([ROWID] IS NULL)', '')

               Set @strSQL = 'Delete ['+@CurrentTable+'] where '+@Fields
               Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
            end
            
            Exec uSP_Sys_MoveFile @CurrentPath, @txtFileName

            --[Reset advance options]
            Exec uSP_Sys_Advanced_Options 1
          end try
          begin catch
               Set @Cnt = -1
               Set @Msg = '匯入 ['+@CurrentPath+'\'+@txtFileName+'] 檔案失敗.(錯誤訊息：'+ERROR_MESSAGE()+')'
               Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
          end catch
          Fetch Next From Cur_txtFiles into @txtFileName
       end

       --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
       -- 這裡為了取筆數，所以不另外修改成 uSP_Exec_SQL
       -- 查詢匯入筆數
       delete @RowCount
       set @strSQL = 'select count(1) as Cnt from '+@CurrentTable
       set @ErrCode_Msg = 'SqlCode-1.4:'+@strSQL
       insert into @RowCount Exec(@strSQL)

       Select @Cnt = Cnt from @RowCount
       Set @Msg = '共匯入 ['+CONVERT(Varchar(100), @Cnt)+'] 筆資料 至 ['+@CurrentTable+']資料表.'
       Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

       --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
       if @SubDirName not like '%@%'
       begin
          -- 增加 rowid 建立 Text 內容每一筆的唯一鍵值
          Set @Msg = '產生 Rowid 於 ['+@CurrentTable+'] 資料表.'
          set @Cnt =0
          Set @strSQL ='alter table ['+@CurrentTable+'] add rowid int identity(1, 1)'
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

          --set @strSQL = 'alter table ['+@HistoryTable+'] add imp_date datetime '
          --Exec uSP_Exec_SQL @Proc, @Msg, @strSQL

          --set @strSQL = 'update ['+@HistoryTable+'] set imp_date=getdate() '
          --Exec uSP_Exec_SQL @Proc, @Msg, @strSQL

          --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
          -- 清除所有欄位皆含 NULL 的資料
          set @Msg='清除所有欄位皆含 NULL 的資料'
          exec uSP_Sys_Get_Columns '', @CurrentTable, '(@Col_Name is null)', ' and ', @Fields output 
          Set @Fields = REPLACE(@Fields, 'AND ([TXTFILENAME] IS NULL)', '')
          --2013/5/3 排除 RowID
          --Set @Fields = REPLACE(@Fields, 'AND ([ROWID] IS NULL)', '')

          Set @strSQL = 'Delete ['+@CurrentTable+'] where '+@Fields
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       end

       --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
       -- 清除資料表內所有 NULL 值
       exec uSP_Sys_Clean_Table_Null @CurrentTable

       begin try
         -- 匯入完後搬移檔案至歷史目錄
         -- exec uSP_Get_Columns '', @CurrentTable, 
--                              '(M.@Col_Name = D.@Col_Name)', ' And ', @Fields output 
         exec uSP_Sys_Get_Columns '', @CurrentTable, '@Col_Name', ', ', @Fields output 

         If (OBJECT_ID(@HistoryTable) Is Not Null)
         begin
            set @strSQL = 'Insert into '+@HistoryTable+' '+@CR+
                          '('+@Fields+')'+@CR+
                          'select '+@Fields+' '+@CR+
                          '  from '+@CurrentTable+' M ' --+@CR+
--                           ' where not exists '+@CR+
--                           '       (select * from '+@HistoryTable+' D '+@CR+
--                           '         where '+@Fields+')'
                      

         end
         else
         begin
            -- Table Not Exist
            -- 去除 Rowid 的自動增量
            set @strSQL = 'select '+@Fields+' into '+@HistoryTable+@CR+
                           '  from '+@CurrentTable+' M '
                
         end
         --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
         Set @Msg = '匯入歷史資料表 ['+@HistoryTable+']'
         Set @Cnt =0
         Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
                     
         --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
         Select @Cnt = Cnt from @RowCount
         Set @Msg = '共匯入 ['+CONVERT(Varchar(100), @Cnt)+'] 筆資料 至 ['+@HistoryTable+'] 資料表.'
         Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

         --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
         Set @Msg = '將刪除 ['+@HistoryTable+'] 資料表含有增量條件欄位.'
         set @Cnt = 0
             
         set @strSQL = 'exec SP_rename ''['+@HistoryTable+'].rowid'', ''rowid_tmp'', ''column'' '
         Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

         set @strSQL = 'alter table ['+@HistoryTable+'] add rowid int '
         Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
         
         set @strSQL = 'alter table ['+@HistoryTable+'] drop column rowid_tmp '
         Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

         -- 2013/11/28 增加回傳匯入成功檔案數
         Set @Cnt_Suss = @Cnt_Suss + 1




            begin    
              Set @Msg = '更新資料表 ['+ @TB_tmp_Name +'] 中 [' + @SubDirName +'] 成功'
              Set @strSQL = ' Update '+ @TB_tmp_Name + ' ' +@CR+
                            '    set xlsFileName = '''+@txtFileName+''', ' +@CR+
                            '        Flag = ''1'', ' +@CR+
                            '        import_date = getdate(),' +@CR+
                            '        Update_date = '''' ' +@CR+
                            '  where Folder_Name = ''' + @CurrentDirName + ''''
              Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
            end



       end try
       begin catch
          Set @Cnt = -1
          Set @Msg = '匯入歷史資料表 ['+@HistoryTable+'] 失敗.(錯誤訊息：'+ERROR_MESSAGE()+')'
          Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
       end catch
       
       Close Cur_txtFiles
       DEALLOCATE Cur_txtFiles
    end
    Fetch Next From Cur_SubDirs into @SubDirName
  end
  Close Cur_SubDirs
  DEALLOCATE Cur_SubDirs
  --[Disabled advance options]
  Exec uSP_Sys_Advanced_Options 0  
  
  set @strSQL =''
  set @Cnt = 0
  Set @Msg = '匯入結束...'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  Return(@Cnt_Suss)
  Goto End_Exit
  
Exception:  
  Return(@Errcode)

End_Exit:

end
GO
