USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Imp_xls_to_db]    Script Date: 07/24/2017 14:44:00 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Imp_xls_to_db]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[uSP_Sys_Imp_xls_to_db]
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_Sys_Imp_xls_to_db
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [變更 Trans_Log 訊息顯示內容]
   Return Code: -1 失敗, 正整數 代表成功讀取 N 個 Excel 檔案
   
   HDR ( HeaDer Row )
   若指定值為 Yes，代表 Excel 檔中的工作表第一列是欄位名稱
   若指定值為 No，代表 Excel 檔中的工作表第一列就是資料了，沒有欄位名稱
   
   IMEX ( IMport EXport mode )
   當 IMEX=0 時為「Export mode 匯出模式」，這個模式開啟的 Excel 檔案只能用來做「寫入」用途。
   當 IMEX=1 時為「Import mode 匯入模式」，這個模式開啟的 Excel 檔案只能用來做「讀取」用途。
   當 IMEX=2 時為「Linked mode (full update capabilities) 連結模式」，這個模式開啟的 Excel 檔案可同時支援「讀取」與「寫入」用途。   
   
   資料參考
   http://blog.miniasp.com/post/2008/08/05/How-to-read-Excel-file-using-OleDb-correctly.aspx

   -- 目錄名稱末字若有以下符號則代表檔案處理方式：
   
   #: Excel 檔案第一行不為欄位名稱。
   @: 目錄內的每一個 Excel 都轉成一個 Table 並以檔案名稱底線前方為 TableName, Ex: FileName MOMO_20141218.xls => TableName 為 Ori_XLS#MOMO.

  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  
  --[Enabled advance options]
  Exec uSP_Sys_Advanced_Options 1 
  
  --[Declare]
  declare @Proc Varchar(50) = 'uSP_Sys_Imp_xls_to_db' -- [Procedure Name]

  declare @ImportRoot Varchar (1000) = 'D:\Transform_Data'
  declare @xlsRoot Varchar (1000) = @ImportRoot + '\Import_Xls_To_DB' -- [Excel root path]
  declare @HistoryRoot Varchar(1000) = @ImportRoot + '\History_Import' -- [History Import path]
  declare @CurrentPath Varchar(1000)=  @xlsRoot
  declare @xlsSheet Varchar(100) = '[Sheet1$]'
  Declare @Errcode int = -1

  declare @xlsSubDir table (SubDirName varchar(255)) --目錄列表
  declare @SubDirName Varchar(255) --目錄名稱
  
  declare @xlsFiles table (xlsFileName varchar(255)) --檔案列表
  declare @xlsFileName Varchar(255) --檔案名稱
  declare @FileName Varchar(255) --檔案名稱不含附檔案名稱
  declare @SplitFileName Varchar(255) --檔案名稱分隔符號後名之名稱
  declare @RowCount table (cnt int)
  
  declare @FirstWord Varchar(100) = 'Ori_xls#' --匯入後所產生的資料表前置檔案名稱
  declare @LastWord Varchar(100) = '' --匯入後所產生的資料表後置檔案名稱
  declare @CurrentTable Varchar(100) =''
  declare @CurrentDirName Varchar(100) = '' -- 取得當前的目錄名稱
  declare @HistoryTable Varchar(100) =''

  declare @OLE_dbP1 Varchar(100) = 'Microsoft.ACE.OLEDB.12.0'
  declare @OLE_dbP2 Varchar(100) = 'Excel 12.0; HDR=Yes; IMEX=1; DataBase='
  
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
  -- 此處有重大訊息顯示，所以不另外修改成 Exec uSP_Sys_Exec_SQL 語法
  
  Set @Msg = 'Excel 引擎啟動測試.'
  Set @Cnt = 0

  begin try
    set @xlsFileName ='測試檔案(勿刪).xls'
    set @FileName = Substring(@xlsFileName, 1, CHARINDEX('.', @xlsFileName)-1)
    set @OLE_dbP2 = Replace(@OLE_dbP2 , 'HDR=No', 'HDR=Yes')
    Set @LastSQL = ' From OpenRowSet ('''+@OLE_dbP1+''', '''+@OLE_dbP2+@ImportRoot+'\'+@xlsFileName+''', ''Select * From '+@xlsSheet+''')'
    Set @strSQL = 'Select * '+ @LastSQL
    set @Errcode_Msg = 'SqlCode-1.1:'+@strSQL

    Exec (@strSQL)
    
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
  
  
  --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
  --[Create Xls Table]
  --判斷是否存在資料表，不存在則新增一個
  IF not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_tmp_Name+']') AND type in (N'U'))
  begin
     Set @Msg = '新增Table ['+@TB_tmp_Name+']。'
     Set @strSQL = 'CREATE TABLE [dbo].['+@TB_tmp_Name+']( ' +@CR+
                   '     [UniqueID] [int] IDENTITY(1,1) NOT NULL, ' +@CR+
                   '     [Folder_Name] [varchar](100) NULL, ' +@CR+
                   '     [XlsFileName] [varchar](100) NULL, ' +@CR+
                   '     [Flag] [varchar](1) NULL, ' +@CR+
                   '     [Import_date] [datetime] NULL, ' +@CR+
                   '     [Update_date] [datetime] NULL ' +@CR+
                   ') ON [PRIMARY] ' 
     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end
  --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
  
  --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
  --[Set Values]
  Set @strSQL =''
  Set @Msg = '匯入開始...'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  
  Set @Msg = '取得外部 Excel 目錄路徑 ['+@HistoryRoot+']'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  
  insert into @xlsSubDir execute xp_subdirs @xlsRoot
  
  --Set @strSQL = 'Insert into '+ @TB_tmp_Name + '(Folder_Name,xlsFileName,Flag,import_date,Update_date) ' +@CR+
  --              ' select SubDirName as Folder_Name,'' as xlsFileName ,''0'' as Flag,import_date=convert(date, getdate()),Update_date=convert(date, getdate()))' +@CR+
  --              '   from ' + @xlsSubDir + ' where not exists (select Folder_Name from '+@TB_tmp_Name+' )'
  
  delete @xlsSubDir where SubDirName is null
  
  select @Cnt = COUNT(*) from @xlsSubDir
  Set @Msg = '共取得 ['+Convert(Varchar(100), @Cnt)+'] 目錄.'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
    
  declare Cur_SubDirs Cursor local static For
    select * From @xlsSubDir

  -- 子目錄循序讀取
  Open Cur_SubDirs 
  Fetch Next From Cur_SubDirs into @SubDirName
  
  While @@FETCH_STATUS = 0 
  Begin
    set @CurrentPath = @xlsRoot +'\'+@SubDirName
    set @CurrentDirName = Replace(Replace(@SubDirName, '@', ''), '#', '')
    set @HistoryTable = @CurrentDirName+'_History'
    set @CurrentTable = @FirstWord+@CurrentDirName
    
    set @cnt = (select count(1) from Import_XLS_Schedule where Folder_Name = @CurrentDirName)
    if @cnt = 0
    begin
        Set @strSQL = 'Insert into '+ @TB_tmp_Name +@CR+
                      '(Folder_Name, XlsFileName, Flag, Import_date, Update_date) ' +@CR+
                      'Values (''' + @CurrentDirName +''', '''', ''0'', getdate() , '''')'
        Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    end

    set @Cmd = 'dir '+@CurrentPath + '\*.xl* /b'
    -- Clear Variable
    delete @xlsFiles
    -- Get Import files
    insert into @xlsFiles Execute xp_cmdshell @Cmd
    
    -- Clear Error Message
    delete @xlsFiles where xlsFileName is null or xlsFileName like '%找不到%' or xlsFileName like '%Found%'
    select @Cnt = COUNT(*) from @xlsFiles
    Set @Msg = '讀取 ['+@CurrentPath+'] 目錄, 共取得 ['+CONVERT(Varchar(100), @Cnt)+']個檔案.'
    set @strSQL = ''
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

    -- Found Files
    if @Cnt <> 0
    begin
       if @SubDirName not like '%@%'
       begin
          set @Msg = '清除資料表 ['+@CurrentTable+'].'
          set @strSQL='Drop Table ['+@CurrentTable+']'
          IF OBJECT_ID(@CurrentTable) Is Not Null 
             -- Exec (@strSQL)
             Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       end

       Declare Cur_xlsFiles Cursor local static For
         Select * from @xlsFiles

       -- 目錄內的檔案循序讀取
       Open Cur_xlsFiles 
       Fetch Next From Cur_xlsFiles into @xlsFileName
       While @@FETCH_STATUS = 0 
       begin
          begin try
            --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
            --set @imp_date = getdate()
            set @FileName = Convert(nVarchar(255), +Substring(@xlsFileName, 1, CHARINDEX('.', @xlsFileName)-1))
            set @SplitFileName = Convert(nVarchar(255), +Substring(@FileName, CHARINDEX('_', @FileName)+1, Len(@FileName)))

            set @Columns = ', XlsFileName=Convert(Varchar(255), '''+@FileName+'''), SplitFileName=Convert(Varchar(255), '''+@SplitFileName+'''), Imp_date=getdate()'

            -- 2013/11/18 Rickliu 目錄名稱若有 # 符號則代表，Excel 檔案第一行不為欄位名稱。
            if @SubDirName like '%#%'
               set @OLE_dbP2 = Replace(@OLE_dbP2 , 'HDR=Yes', 'HDR=No')
            else
               set @OLE_dbP2 = Replace(@OLE_dbP2 , 'HDR=No', 'HDR=Yes')
               
            Set @LastSQL = ' From OpenRowSet ('''+@OLE_dbP1+''', '''+@OLE_dbP2+@CurrentPath+'\'+@xlsFileName+''', ''Select * From '+@xlsSheet+''')'

            -- 2014/12/18 Rickliu 目錄名稱若有 @ 符號則代表，每一個 Excel 檔案都另行產生一個 Table。
            if @SubDirName like '%@%'
            begin
               set @CurrentTable = @FirstWord+
                                   @CurrentDirName+'_'+
                                   Convert(nVarchar(255), +Substring(@FileName, 1, CHARINDEX('_', @FileName)-1))
                                   
               IF OBJECT_ID(@CurrentTable) Is Not Null 
               begin
                  set @Msg = '清除資料表 ['+@CurrentTable+'].'
                  set @strSQL='Drop Table '+@CurrentTable

                  Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
               end
            end

            IF OBJECT_ID(@CurrentTable) Is Null
               Set @strSQL = 'Select *'+@Columns+' into '+@CurrentTable 
            else
               Set @strSQL = 'Insert into '+@CurrentTable+' Select *'+@Columns
            set @strSQL = @strSQL + @LastSQL
            
            Set @Msg = '準備匯入['+@CurrentPath+'\'+@xlsFileName+'] 至 ['+@CurrentTable+']資料表.'
            Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
            if @Cnt = -1 RaisError(@Msg, 16, 1)
            
            --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
            --2013/5/3 增加 rowid 建立 Excle 內容每一筆的唯一鍵值
            
            --2014/08/12 Rickliu 增加判斷是否有增量條件欄位
            If (OBJECT_ID(@CurrentTable) Is Not Null) and
               (IDENT_SEED(@CurrentTable) is null)
            begin
             --*-*-*
               Set @Msg = '產生 Rowid 於 ['+@CurrentTable+']資料表.'
               set @Cnt =0
               Set @strSQL ='Alter Table '+@CurrentTable+' add rowid int identity(1, 1)'
               Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
               if @Cnt = -1 RaisError(@Msg, 16, 1)
            end
            
            --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
            --清除所有欄位皆含 NULL 的資料
            set @Msg='清除所有欄位皆含 NULL 的資料'
            Exec uSP_Sys_Get_Columns '', @CurrentTable, '(@Col_Name is null)', ' and ', @Fields output 
            Set @Fields = REPLACE(@Fields, 'AND ([XLSFILENAME] IS NULL)', '')
            --2013/5/3 排除 RowID
            Set @Fields = REPLACE(@Fields, 'AND ([ROWID] IS NULL)', '')

            Set @strSQL = 'Delete '+@CurrentTable+' where '+@Fields
            Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
            if @Cnt = -1 RaisError(@Msg, 16, 1)
               
            --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
            -- 這裡為了取筆數，所以不另外修改成 SP_Exec_SQL
            --查詢匯入筆數
            delete @RowCount
            set @strSQL = 'Select count(*) '
            set @strSQL = @strSQL + @LastSQL
            set @Errcode_Msg = 'SqlCode-1.4:'+@strSQL
            insert into @RowCount Exec(@strSQL)

            If OBJECT_ID(@CurrentTable) Is Not Null 
            begin
               Select @Cnt = Cnt from @RowCount
               Set @Msg = '共匯入 ['+CONVERT(Varchar(100), @Cnt)+'] 筆資料 至 ['+@CurrentTable+']資料表.'
               Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
            end            
            
            Exec @Cnt = uSP_Sys_MoveFile @CurrentPath, @xlsFileName
            if @Cnt = -1 RaisError(@Msg, 16, 1)               
            
            --[Reset advance options]
            Exec uSP_Sys_Advanced_Options 1
            
            -- 2013/11/28 增加回傳匯入成功檔案數
            Set @Cnt_Suss = @Cnt_Suss + 1

            Set @strSQL = ' Update '+ @TB_tmp_Name + ' ' +@CR+
                          '    set XlsFileName = '''+@xlsFileName+''', ' +@CR+
                          '        Flag = ''1'', ' +@CR+
                          '        Import_date = getdate(),' +@CR+
                          '        Update_date = '''' ' +@CR+
                          '  where Folder_Name = ''' + @CurrentDirName + ''''
            Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
          end try
          begin catch
            Set @Cnt = @Errcode
            Set @Msg = '匯入 ['+@CurrentPath+'\'+@xlsFileName+']檔案失敗.(錯誤訊息：'+ERROR_MESSAGE()+')'
            Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
          end catch
          Fetch Next From Cur_xlsFiles into @xlsFileName
       end
  
       --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
       --清除資料表內所有 NULL 值
       Exec uSP_Sys_Clean_Table_Null @CurrentTable
       Close Cur_xlsFiles
       DEALLOCATE Cur_xlsFiles
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
  
Exception:  
  Return(@Errcode)
  

end
GO
