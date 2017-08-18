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
   
   -- �ؿ��W�٥��r�Y���H�U�Ÿ��h�N���ɮ׳B�z�覡�G
   
   PS: �ѩ�ϥ� BULK �覡�פJ�A�ҥH�Τ@�}�@�����A�ñN��ƥ����פJ�C
   @: �ؿ������C�@�� Text ���ন�@�� Table �åH�ɮצW�٩��u�e�謰 TableName, Ex: FileName MOMO_20141218.xls => TableName �� Ori_TXT#MOMO.

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

  declare @txtSubDir table (SubDirName varchar(255)) --�ؿ��C��
  declare @SubDirName Varchar(255) = '' --�ؿ��W��
  
  declare @txtFiles table (txtFileName varchar(255)) --�ɮצC��
  declare @txtFileName Varchar(255) = '' --�ɮצW��
  declare @FileName Varchar(255) --�ɮצW�٤��t���ɮצW��
  declare @SplitFileName Varchar(255) --�ɮצW�٤��j�Ÿ���W���W��
  declare @RowCount table (cnt int)
  
  declare @FirstWord Varchar(100) = 'Ori_Txt#' --�פJ��Ҳ��ͪ���ƪ�e�m�ɮצW��
  declare @LastWord Varchar(100) = '' --�פJ��Ҳ��ͪ���ƪ��m�ɮצW��
  declare @CurrentTable Varchar(100) =''
  declare @CurrentDirName Varchar(100) = '' -- ���o��e���ؿ��W��
  declare @HistoryTable Varchar(100) =''

  declare @OldbP1 Varchar(100) = 'WITH (FIELDTERMINATOR='' '', ROWTERMINATOR=''\n'') '
  
  declare @HaveFileName Varchar(100) = '' --��ؿ��W�ٷ�@�פJ�᪺�䤤�@�����
  
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
  -- ���B�����j�T����ܡA�ҥH���t�~�ק令 Exec uSP_Exec_SQL �y�k
  
  Set @Msg = 'Text �����Ұʴ���.'
  Set @Cnt = 0

  begin try
    set @CurrentTable ='Test_tmp'
    set @txtFileName ='�����ɮ�(�ŧR).txt'
    
    if @CurrentPath like '%@%'
       set @FileName = Substring(@txtFileName, 1, CHARINDEX('.', @txtFileName)-1)
    else
       set @FileName = ''

    IF OBJECT_ID(@CurrentTable) Is Not Null 
    begin
       set @Msg ='�R�����ո�ƪ� ['+@CurrentTable+']'
       set @strSQL='Drop TABLE '+@CurrentTable
       Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
    end
 
    set @Msg ='�إߴ��ո�ƪ� ['+@CurrentTable+']'
    set @strSQL='CREATE TABLE '+@CurrentTable+' (F1 VARCHAR(1000)) '
    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

    set @Msg = '�פJ���ո�ƪ� ['+@CurrentTable+']'
    Set @strSQL = 'BULK INSERT '+@CurrentTable+' FROM '''+@ImportRoot+'\'+@txtFileName+''' '+ @OldbP1

    set @ErrCode_Msg = 'SqlCode-1.1:'+@strSQL

    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

    
    Set @Cnt = 0
    Set @Msg = @Msg+'���\.'
    
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  end try
  begin catch
    Set @Cnt = @Errcode
    Set @Msg = @Msg+'���ѡA�פ�Ҧ���J�ʧ@�A�Э��s�Ұ� SQL Server.(���~�T���G'+ERROR_MESSAGE()+')'
    Exec uSP_Sys_Write_Log @Proc, @Msg, @Cnt
    
    goto Exception
  end catch
  
  
  --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

  
  --[Create Xls Table]
  --�P�_�O�_�s�b��ƪ�A���s�b�h�s�W�@��
    IF not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_tmp_Name+']') AND type in (N'U'))
    begin
       Set @Msg = '�s�WTable ['+@TB_tmp_Name+']�C'
       
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
  Set @Msg = '�פJ�}�l...'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  
  Set @Msg = '���o�~�� Text �ؿ����| ['+@HistoryRoot+']'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  
  insert into @txtSubDir execute xp_subdirs @txtRoot
  delete @txtSubDir where SubDirName is null
  
  select @Cnt = COUNT(1) from @txtSubDir
  Set @Msg = '�@���o ['+Convert(Varchar(100), @Cnt)+'] �ؿ�.'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
    
  declare Cur_SubDirs Cursor local static For
    select * From @txtSubDir
  -- �l�ؿ��`��Ū��
  Open Cur_SubDirs -- �l�ؿ��`��Ū��
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
    delete @txtFiles where txtFileName is null or txtFileName like '%�䤣��%' or txtFileName like '%Found%'
    select @Cnt = COUNT(1) from @txtFiles
    Set @Msg = 'Ū�� ['+@CurrentPath+'] �ؿ�, �@���o ['+CONVERT(Varchar(100), @Cnt)+'] ���ɮ�.'
    set @strSQL = ''
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

    -- Found Files
    if @Cnt <> 0
    begin
       if @SubDirName not like '%@%'
       begin
          set @Msg = '�M����ƪ� ['+@CurrentTable+'].'
          set @strSQL='DROP TABLE ['+@CurrentTable+']'
          IF OBJECT_ID(@CurrentTable) Is Not Null 
             -- Exec (@strSQL)
             Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       end

       Declare Cur_txtFiles Cursor local static For
         Select * from @txtFiles
       -- �ؿ������ɮ״`��Ū��
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


            -- 2014/12/18 Rickliu �ؿ��W�٭Y�� @ �Ÿ��h�N��A�C�@�� Text �ɮ׳��t�沣�ͤ@�� Table�C
            if @SubDirName like '%@%'
            begin
               set @CurrentTable = @FirstWord+
                                   @CurrentDirName+'_'+
                                   Convert(nVarchar(255), +Substring(@FileName, 1, CHARINDEX('_', @FileName)-1))
                                   
               set @Msg = '�M����ƪ� ['+@CurrentTable+'].'
               set @strSQL='DROP TABLE '+@CurrentTable
               IF OBJECT_ID(@CurrentTable) Is Not Null 
                  --Exec (@strSQL)
                  Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
            end

            IF OBJECT_ID(@CurrentTable) Is Null 
            begin
               set @Msg ='�إ߸�ƪ� ['+@CurrentTable+']'
               set @strSQL='CREATE TABLE ['+@CurrentTable+'] (F1 VARCHAR(1000)) '

               Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
            end
    
            Set @strSQL = 'BULK INSERT ['+@CurrentTable+'] FROM '''+@CurrentPath+'\'+@txtFileName+''' '+ @OldbP1
    
            Set @Msg = '�ǳƶפJ ['+@CurrentPath+'\'+@txtFileName+'] �� ['+@CurrentTable+']��ƪ�.'
            Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

            if @SubDirName like '%@%'
            begin
               -- �W�[ rowid �إ� Text ���e�C�@�����ߤ@���
               Set @Msg = '���� Rowid �� ['+@CurrentTable+'] ��ƪ�.'
               set @Cnt =0
               Set @strSQL ='alter table ['+@CurrentTable+'] add rowid int identity(1, 1)'
               Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

               --set @strSQL = 'alter table ['+@HistoryTable+'] add imp_date datetime '
               --Exec uSP_Exec_SQL @Proc, @Msg, @strSQL

               --set @strSQL = 'update ['+@HistoryTable+'] set imp_date=getdate() '
               --Exec uSP_Exec_SQL @Proc, @Msg, @strSQL

               --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
               -- �M���Ҧ����ҧt NULL �����
               set @Msg='�M���Ҧ����ҧt NULL �����'
               exec uSP_Sys_Get_Columns '', @CurrentTable, '(@Col_Name is null)', ' and ', @Fields output 
               Set @Fields = REPLACE(@Fields, 'AND ([TXTFILENAME] IS NULL)', '')
               --2013/5/3 �ư� RowID
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
               Set @Msg = '�פJ ['+@CurrentPath+'\'+@txtFileName+'] �ɮץ���.(���~�T���G'+ERROR_MESSAGE()+')'
               Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
          end catch
          Fetch Next From Cur_txtFiles into @txtFileName
       end

       --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
       -- �o�̬��F�����ơA�ҥH���t�~�ק令 uSP_Exec_SQL
       -- �d�߶פJ����
       delete @RowCount
       set @strSQL = 'select count(1) as Cnt from '+@CurrentTable
       set @ErrCode_Msg = 'SqlCode-1.4:'+@strSQL
       insert into @RowCount Exec(@strSQL)

       Select @Cnt = Cnt from @RowCount
       Set @Msg = '�@�פJ ['+CONVERT(Varchar(100), @Cnt)+'] ����� �� ['+@CurrentTable+']��ƪ�.'
       Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

       --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
       if @SubDirName not like '%@%'
       begin
          -- �W�[ rowid �إ� Text ���e�C�@�����ߤ@���
          Set @Msg = '���� Rowid �� ['+@CurrentTable+'] ��ƪ�.'
          set @Cnt =0
          Set @strSQL ='alter table ['+@CurrentTable+'] add rowid int identity(1, 1)'
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

          --set @strSQL = 'alter table ['+@HistoryTable+'] add imp_date datetime '
          --Exec uSP_Exec_SQL @Proc, @Msg, @strSQL

          --set @strSQL = 'update ['+@HistoryTable+'] set imp_date=getdate() '
          --Exec uSP_Exec_SQL @Proc, @Msg, @strSQL

          --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
          -- �M���Ҧ����ҧt NULL �����
          set @Msg='�M���Ҧ����ҧt NULL �����'
          exec uSP_Sys_Get_Columns '', @CurrentTable, '(@Col_Name is null)', ' and ', @Fields output 
          Set @Fields = REPLACE(@Fields, 'AND ([TXTFILENAME] IS NULL)', '')
          --2013/5/3 �ư� RowID
          --Set @Fields = REPLACE(@Fields, 'AND ([ROWID] IS NULL)', '')

          Set @strSQL = 'Delete ['+@CurrentTable+'] where '+@Fields
          Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       end

       --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
       -- �M����ƪ��Ҧ� NULL ��
       exec uSP_Sys_Clean_Table_Null @CurrentTable

       begin try
         -- �פJ����h���ɮצܾ��v�ؿ�
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
            -- �h�� Rowid ���۰ʼW�q
            set @strSQL = 'select '+@Fields+' into '+@HistoryTable+@CR+
                           '  from '+@CurrentTable+' M '
                
         end
         --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
         Set @Msg = '�פJ���v��ƪ� ['+@HistoryTable+']'
         Set @Cnt =0
         Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
                     
         --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
         Select @Cnt = Cnt from @RowCount
         Set @Msg = '�@�פJ ['+CONVERT(Varchar(100), @Cnt)+'] ����� �� ['+@HistoryTable+'] ��ƪ�.'
         Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

         --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
         Set @Msg = '�N�R�� ['+@HistoryTable+'] ��ƪ�t���W�q�������.'
         set @Cnt = 0
             
         set @strSQL = 'exec SP_rename ''['+@HistoryTable+'].rowid'', ''rowid_tmp'', ''column'' '
         Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

         set @strSQL = 'alter table ['+@HistoryTable+'] add rowid int '
         Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
         
         set @strSQL = 'alter table ['+@HistoryTable+'] drop column rowid_tmp '
         Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

         -- 2013/11/28 �W�[�^�ǶפJ���\�ɮ׼�
         Set @Cnt_Suss = @Cnt_Suss + 1




            begin    
              Set @Msg = '��s��ƪ� ['+ @TB_tmp_Name +'] �� [' + @SubDirName +'] ���\'
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
          Set @Msg = '�פJ���v��ƪ� ['+@HistoryTable+'] ����.(���~�T���G'+ERROR_MESSAGE()+')'
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
  Set @Msg = '�פJ����...'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  Return(@Cnt_Suss)
  Goto End_Exit
  
Exception:  
  Return(@Errcode)

End_Exit:

end
GO
