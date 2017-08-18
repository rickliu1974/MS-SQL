USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Imp_Ori_DataBases]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Imp_Ori_DataBases]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[uSP_Sys_Imp_Ori_DataBases](@Application_Name Varchar(100)='')
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_Sys_Imp_Ori_DataBases
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [�ܧ� Trans_Log �T����ܤ��e]
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  
  Declare @Proc Varchar(50) = 'uSP_Sys_Imp_Ori_DataBases'
  Declare @strSQL Varchar(Max) =''
  Declare @Cnt Int =0
  Declare @Msg Varchar(Max) =''
  Declare @CMD Varchar(100) =''
  Declare @Server_Name Varchar(100) =''
  Declare @DataBase_Name Varchar(100) =''
  
  Declare Cur_Ori_DataBases Cursor local static For
    Select Server_Name, Database_Name
      from Sys_Ori_Databases
     where [Enabled] = 1 -- True
       and Application_Name=@Application_Name
  
  Open Cur_Ori_DataBases 
  Fetch Next From Cur_Ori_DataBases into @Server_Name, @DataBase_Name

  Set @Cnt=@@CURSOR_ROWS
  Set @strSQL =''
  
  if @Cnt = 0
  begin
     Set @strSQL=''
     Set @Msg='��l��ƪ��Ƥ���, ��������פJ�@�~.'
     Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
  end
  else
  begin
     While @@FETCH_STATUS = 0 
     Begin 
        Set @Msg='�ǳư���פJ ['+@Server_Name+'.'+@DataBase_Name+'] �{�ǡC'
        Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
        Begin Try 
          --2013/6/25 �אּ SP_Imp_Ori_Tables2 �D�n�Y�u�P�B�ɶ�
          --Exec uSP_Imp_Ori_Tables @Application_Name, @Server_Name, @DataBase_Name
          Exec uSP_Sys_Imp_Ori_Tables @Application_Name, @Server_Name, @DataBase_Name
        end Try
        begin catch
          set @Cnt = -1
          Set @Msg='����פJ ['+@Server_Name+'.'+@DataBase_Name+'] �{�ǥ���...(���~�T��:'+ERROR_MESSAGE()+')'
          Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
        end catch

        Set @Cnt =0
        Set @Msg = '[Process End]'+REPLICATE('=', 50)
        Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Cnt
        Fetch Next From Cur_Ori_DataBases into @Server_Name, @DataBase_Name
     end
     Close Cur_Ori_DataBases
     DEALLOCATE Cur_Ori_DataBases
  end
end
GO
