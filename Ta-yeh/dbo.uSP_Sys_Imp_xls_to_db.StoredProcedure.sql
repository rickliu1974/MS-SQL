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
   Updated Date: 2013/09/06 [�ܧ� Trans_Log �T����ܤ��e]
   Return Code: -1 ����, ����� �N���\Ū�� N �� Excel �ɮ�
   
   HDR ( HeaDer Row )
   �Y���w�Ȭ� Yes�A�N�� Excel �ɤ����u�@��Ĥ@�C�O���W��
   �Y���w�Ȭ� No�A�N�� Excel �ɤ����u�@��Ĥ@�C�N�O��ƤF�A�S�����W��
   
   IMEX ( IMport EXport mode )
   �� IMEX=0 �ɬ��uExport mode �ץX�Ҧ��v�A�o�ӼҦ��}�Ҫ� Excel �ɮץu��ΨӰ��u�g�J�v�γ~�C
   �� IMEX=1 �ɬ��uImport mode �פJ�Ҧ��v�A�o�ӼҦ��}�Ҫ� Excel �ɮץu��ΨӰ��uŪ���v�γ~�C
   �� IMEX=2 �ɬ��uLinked mode (full update capabilities) �s���Ҧ��v�A�o�ӼҦ��}�Ҫ� Excel �ɮץi�P�ɤ䴩�uŪ���v�P�u�g�J�v�γ~�C   
   
   ��ưѦ�
   http://blog.miniasp.com/post/2008/08/05/How-to-read-Excel-file-using-OleDb-correctly.aspx

   -- �ؿ��W�٥��r�Y���H�U�Ÿ��h�N���ɮ׳B�z�覡�G
   
   #: Excel �ɮײĤ@�椣�����W�١C
   @: �ؿ������C�@�� Excel ���ন�@�� Table �åH�ɮצW�٩��u�e�謰 TableName, Ex: FileName MOMO_20141218.xls => TableName �� Ori_XLS#MOMO.

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

  declare @xlsSubDir table (SubDirName varchar(255)) --�ؿ��C��
  declare @SubDirName Varchar(255) --�ؿ��W��
  
  declare @xlsFiles table (xlsFileName varchar(255)) --�ɮצC��
  declare @xlsFileName Varchar(255) --�ɮצW��
  declare @FileName Varchar(255) --�ɮצW�٤��t���ɮצW��
  declare @SplitFileName Varchar(255) --�ɮצW�٤��j�Ÿ���W���W��
  declare @RowCount table (cnt int)
  
  declare @FirstWord Varchar(100) = 'Ori_xls#' --�פJ��Ҳ��ͪ���ƪ�e�m�ɮצW��
  declare @LastWord Varchar(100) = '' --�פJ��Ҳ��ͪ���ƪ��m�ɮצW��
  declare @CurrentTable Varchar(100) =''
  declare @CurrentDirName Varchar(100) = '' -- ���o��e���ؿ��W��
  declare @HistoryTable Varchar(100) =''

  declare @OLE_dbP1 Varchar(100) = 'Microsoft.ACE.OLEDB.12.0'
  declare @OLE_dbP2 Varchar(100) = 'Excel 12.0; HDR=Yes; IMEX=1; DataBase='
  
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
  -- ���B�����j�T����ܡA�ҥH���t�~�ק令 Exec uSP_Sys_Exec_SQL �y�k
  
  Set @Msg = 'Excel �����Ұʴ���.'
  Set @Cnt = 0

  begin try
    set @xlsFileName ='�����ɮ�(�ŧR).xls'
    set @FileName = Substring(@xlsFileName, 1, CHARINDEX('.', @xlsFileName)-1)
    set @OLE_dbP2 = Replace(@OLE_dbP2 , 'HDR=No', 'HDR=Yes')
    Set @LastSQL = ' From OpenRowSet ('''+@OLE_dbP1+''', '''+@OLE_dbP2+@ImportRoot+'\'+@xlsFileName+''', ''Select * From '+@xlsSheet+''')'
    Set @strSQL = 'Select * '+ @LastSQL
    set @Errcode_Msg = 'SqlCode-1.1:'+@strSQL

    Exec (@strSQL)
    
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
  
  
  --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
  --[Create Xls Table]
  --�P�_�O�_�s�b��ƪ�A���s�b�h�s�W�@��
  IF not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@TB_tmp_Name+']') AND type in (N'U'))
  begin
     Set @Msg = '�s�WTable ['+@TB_tmp_Name+']�C'
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
  Set @Msg = '�פJ�}�l...'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  
  Set @Msg = '���o�~�� Excel �ؿ����| ['+@HistoryRoot+']'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  
  insert into @xlsSubDir execute xp_subdirs @xlsRoot
  
  --Set @strSQL = 'Insert into '+ @TB_tmp_Name + '(Folder_Name,xlsFileName,Flag,import_date,Update_date) ' +@CR+
  --              ' select SubDirName as Folder_Name,'' as xlsFileName ,''0'' as Flag,import_date=convert(date, getdate()),Update_date=convert(date, getdate()))' +@CR+
  --              '   from ' + @xlsSubDir + ' where not exists (select Folder_Name from '+@TB_tmp_Name+' )'
  
  delete @xlsSubDir where SubDirName is null
  
  select @Cnt = COUNT(*) from @xlsSubDir
  Set @Msg = '�@���o ['+Convert(Varchar(100), @Cnt)+'] �ؿ�.'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
    
  declare Cur_SubDirs Cursor local static For
    select * From @xlsSubDir

  -- �l�ؿ��`��Ū��
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
    delete @xlsFiles where xlsFileName is null or xlsFileName like '%�䤣��%' or xlsFileName like '%Found%'
    select @Cnt = COUNT(*) from @xlsFiles
    Set @Msg = 'Ū�� ['+@CurrentPath+'] �ؿ�, �@���o ['+CONVERT(Varchar(100), @Cnt)+']���ɮ�.'
    set @strSQL = ''
    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt

    -- Found Files
    if @Cnt <> 0
    begin
       if @SubDirName not like '%@%'
       begin
          set @Msg = '�M����ƪ� ['+@CurrentTable+'].'
          set @strSQL='Drop Table ['+@CurrentTable+']'
          IF OBJECT_ID(@CurrentTable) Is Not Null 
             -- Exec (@strSQL)
             Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
       end

       Declare Cur_xlsFiles Cursor local static For
         Select * from @xlsFiles

       -- �ؿ������ɮ״`��Ū��
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

            -- 2013/11/18 Rickliu �ؿ��W�٭Y�� # �Ÿ��h�N��AExcel �ɮײĤ@�椣�����W�١C
            if @SubDirName like '%#%'
               set @OLE_dbP2 = Replace(@OLE_dbP2 , 'HDR=Yes', 'HDR=No')
            else
               set @OLE_dbP2 = Replace(@OLE_dbP2 , 'HDR=No', 'HDR=Yes')
               
            Set @LastSQL = ' From OpenRowSet ('''+@OLE_dbP1+''', '''+@OLE_dbP2+@CurrentPath+'\'+@xlsFileName+''', ''Select * From '+@xlsSheet+''')'

            -- 2014/12/18 Rickliu �ؿ��W�٭Y�� @ �Ÿ��h�N��A�C�@�� Excel �ɮ׳��t�沣�ͤ@�� Table�C
            if @SubDirName like '%@%'
            begin
               set @CurrentTable = @FirstWord+
                                   @CurrentDirName+'_'+
                                   Convert(nVarchar(255), +Substring(@FileName, 1, CHARINDEX('_', @FileName)-1))
                                   
               IF OBJECT_ID(@CurrentTable) Is Not Null 
               begin
                  set @Msg = '�M����ƪ� ['+@CurrentTable+'].'
                  set @strSQL='Drop Table '+@CurrentTable

                  Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
               end
            end

            IF OBJECT_ID(@CurrentTable) Is Null
               Set @strSQL = 'Select *'+@Columns+' into '+@CurrentTable 
            else
               Set @strSQL = 'Insert into '+@CurrentTable+' Select *'+@Columns
            set @strSQL = @strSQL + @LastSQL
            
            Set @Msg = '�ǳƶפJ['+@CurrentPath+'\'+@xlsFileName+'] �� ['+@CurrentTable+']��ƪ�.'
            Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
            if @Cnt = -1 RaisError(@Msg, 16, 1)
            
            --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
            --2013/5/3 �W�[ rowid �إ� Excle ���e�C�@�����ߤ@���
            
            --2014/08/12 Rickliu �W�[�P�_�O�_���W�q�������
            If (OBJECT_ID(@CurrentTable) Is Not Null) and
               (IDENT_SEED(@CurrentTable) is null)
            begin
             --*-*-*
               Set @Msg = '���� Rowid �� ['+@CurrentTable+']��ƪ�.'
               set @Cnt =0
               Set @strSQL ='Alter Table '+@CurrentTable+' add rowid int identity(1, 1)'
               Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
               if @Cnt = -1 RaisError(@Msg, 16, 1)
            end
            
            --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
            --�M���Ҧ����ҧt NULL �����
            set @Msg='�M���Ҧ����ҧt NULL �����'
            Exec uSP_Sys_Get_Columns '', @CurrentTable, '(@Col_Name is null)', ' and ', @Fields output 
            Set @Fields = REPLACE(@Fields, 'AND ([XLSFILENAME] IS NULL)', '')
            --2013/5/3 �ư� RowID
            Set @Fields = REPLACE(@Fields, 'AND ([ROWID] IS NULL)', '')

            Set @strSQL = 'Delete '+@CurrentTable+' where '+@Fields
            Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
            if @Cnt = -1 RaisError(@Msg, 16, 1)
               
            --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
            -- �o�̬��F�����ơA�ҥH���t�~�ק令 SP_Exec_SQL
            --�d�߶פJ����
            delete @RowCount
            set @strSQL = 'Select count(*) '
            set @strSQL = @strSQL + @LastSQL
            set @Errcode_Msg = 'SqlCode-1.4:'+@strSQL
            insert into @RowCount Exec(@strSQL)

            If OBJECT_ID(@CurrentTable) Is Not Null 
            begin
               Select @Cnt = Cnt from @RowCount
               Set @Msg = '�@�פJ ['+CONVERT(Varchar(100), @Cnt)+'] ����� �� ['+@CurrentTable+']��ƪ�.'
               Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
            end            
            
            Exec @Cnt = uSP_Sys_MoveFile @CurrentPath, @xlsFileName
            if @Cnt = -1 RaisError(@Msg, 16, 1)               
            
            --[Reset advance options]
            Exec uSP_Sys_Advanced_Options 1
            
            -- 2013/11/28 �W�[�^�ǶפJ���\�ɮ׼�
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
            Set @Msg = '�פJ ['+@CurrentPath+'\'+@xlsFileName+']�ɮץ���.(���~�T���G'+ERROR_MESSAGE()+')'
            Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
          end catch
          Fetch Next From Cur_xlsFiles into @xlsFileName
       end
  
       --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
       --�M����ƪ��Ҧ� NULL ��
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
  Set @Msg = '�פJ����...'
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  Return(@Cnt_Suss)
  
Exception:  
  Return(@Errcode)
  

end
GO
