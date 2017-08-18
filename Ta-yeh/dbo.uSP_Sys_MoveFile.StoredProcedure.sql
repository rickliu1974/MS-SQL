USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_MoveFile]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_Sys_MoveFile]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[uSP_Sys_MoveFile](@Src_Dir Varchar(1000), @Src_FileName Varchar(1000))
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_Sys_MoveFile
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [�ܧ� Trans_Log �T����ܤ��e]
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  
  --[Declare]
  declare @FileName table (File_Exists int, File_Is_Directory int, Parent_Directory_Exists int)
  declare @Proc Varchar(50) = 'uSP_Sys_MoveFile'
  declare @ErrMsg table (Msg Varchar(1000))
  declare @SubDir  table (SubDir varchar(1000))
  declare @vStr Varchar(max)
  
  declare @ImportRoot Varchar (1000) = 'D:\Transform_Data'
  declare @xlsRoot Varchar (1000) = @ImportRoot + '\Import_Xls_To_DB' -- [Excel root path]
  declare @txtRoot Varchar (1000) = @ImportRoot + '\Import_Txt_To_DB' -- [Text root path]
  declare @HistoryRoot Varchar(1000) = @ImportRoot + '\History_Import' -- [History Import path]

  Declare @File_FullName Varchar(1000)
  Declare @Cnt Int = 0
  Declare @CR Varchar(4) = Char(13)+Char(10)
  Declare @Msg Varchar(Max) = ''
  Declare @strSQL Varchar(Max) = ''
  Declare @CMD Varchar(1000)
  Declare @ErrCode int
  Declare @BK_Dir Varchar(1000)
  Declare @BK_File Varchar(1000)
  Declare @Exec_DateTime Varchar(50)=''

  --[Enabled Command Shell]
  Exec uSP_Sys_Advanced_Options 1
  
  --[Set Values]
  Set @BK_Dir = @HistoryRoot+'\'+Convert(Varchar(20), GETDATE(), 112)
  Set @BK_File = @BK_Dir+'\'+@Src_FileName
 
  --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
  --[Check File Exists]
  Set @Msg='����h���R�O...'
  Set @CMD = @Src_Dir+'\'+@Src_FileName
  Set @File_FullName = @Src_Dir+'\'+@Src_FileName

  set @strSQL = 'Exec master..xp_fileexist '+@CMD
  begin try
    set @Cnt = 0
    insert into @FileName Exec master..xp_fileexist @CMD
    Select @Cnt=File_Exists from @FileName
  end try
  begin catch
     Set @Cnt=-1
     Set @Msg='����.(���~�T��:'+ERROR_MESSAGE()+')'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  end catch
  
  if @Cnt = 0
  begin
     Set @strSQL = @CMD
     Set @Cnt=-1
     Set @Msg='�䤣���ɮ׵L�k�h��['+Isnull(@CMD, '')+'].'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  end 
  else
  begin
     --*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
     --[Check Dir Exists]
     Set @Msg='�إ߳ƥ��ؿ� ['+Isnull(@Bk_Dir, '')+']'
     set @CMD = @Bk_Dir
     set @strSQL ='Exec master..xp_cmdshell '+@CMD
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
     
     insert into @FileName Exec master..xp_fileexist @CMD
     Select @Cnt=File_Is_Directory from @FileName
     if @Cnt = 0
     begin
        -- �إߥؿ�
        if @Cnt = 0 
        begin
          begin try
            Set @CMD = 'MD "'+@BK_Dir+'"'
            Insert into @ErrMsg Exec master..xp_cmdshell @CMD
            delete @ErrMsg where Msg is null
            
            --# Get Table ErrMsg------------------------#
            declare Cur_ErrMsg Cursor for
              select Msg from @ErrMsg
              
            open Cur_ErrMsg
            set @Msg = ''
            fetch next from Cur_ErrMsg into @vStr
            While @@FETCH_STATUS=0
            begin
              set @vStr = Rtrim(isnull(@vStr, ''))
              if @vStr <> ''
                 set @Msg = isnull(@Msg, '')+' ### '+@vStr
              fetch next from Cur_ErrMsg into @vStr
            end
            close Cur_ErrMsg
            Deallocate Cur_ErrMsg
            --#-----------------------------------------#
            if @Msg <> ''
            begin
              Set @Msg = Isnull(@Msg, '') +'(�T���^��:'+Isnull(@Msg, '')+')'
              Set @Cnt=-1
              Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
            end
          end try
          begin catch
            Set @Cnt = -1
            Set @Msg = isnull(@Msg, '')+'(���~�T��:'+ERROR_MESSAGE()+')'
            Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
          end catch
        end
     end

     -- �h���ɮ�
     begin try
       Set @Exec_DateTime = Replace(Replace(Replace(Convert(Varchar(50), getdate(), 21), '-', ''), ':', '.'), ' ', '-')
       
       Set @CMD = 'Move /Y "'+isnull(@File_FullName, '')+'" "'+isnull(@BK_Dir, '')+
                  '\�ӷ��ؿ�['+Replace(Replace(@Src_Dir, @xlsRoot+'\', ''), @txtRoot+'\', '')+']-�h���ɶ�['+@Exec_DateTime+']-���ɦW��['+Replace(@Src_FileName, '.', '].')+'"'
       Print @CMD
       Set @Msg=''
       Delete @ErrMsg
       insert into @ErrMsg Exec master..xp_cmdshell @CMD
       delete @ErrMsg where Msg is null

       --# Get Table ErrMsg------------------------#
       declare Cur_ErrMsg Cursor for
       select Msg from @ErrMsg
              
       open Cur_ErrMsg
       set @Msg = ''
       fetch next from Cur_ErrMsg into @vStr
       While @@FETCH_STATUS=0
       begin
         set @vStr = Rtrim(isnull(@vStr, ''))
         if @vStr <> ''
            set @Msg = isnull(@Msg, '')+' ### '+@vStr
         fetch next from Cur_ErrMsg into @vStr
       end
       close Cur_ErrMsg
       Deallocate Cur_ErrMsg
       --#-----------------------------------------#
       if @Msg <> ''
       begin
          Set @Msg='�h���ɮ�['+Isnull(@File_FullName, '')+'].(�T���^��:'+Isnull(@Msg, '')+')'
          set @Cnt=-1
          Exec uSP_Sys_Write_Log @Proc, @Msg, @Cnt
       end
     end try
     begin catch
       Set @Msg = ERROR_MESSAGE()
       Set @Cnt = -1
       Set @Msg = 'Dos Command:'+Isnull('Exec master..xp_cmdshell '+isnull(@Cmd, ''), '')+@CR+
                  '�L�k�h���ɮ� ['+isnull(@BK_Dir, '')+'\'+isnull(@File_FullName, '')+'].(���~�T��:'+isnull(@Msg, '')+')'
       Exec uSP_Sys_Write_Log @Proc, @Msg, @Cnt
     end catch
  end

  --[Disabled Command Shell]
  Exec uSP_Sys_Advanced_Options 0
end
GO
