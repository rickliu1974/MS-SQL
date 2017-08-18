USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Advanced_Options]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Advanced_Options]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[uSP_Sys_Advanced_Options](@Option bit)
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_Sys_Advanced_Options
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [變更訊息顯示內容]
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
  Declare @Proc Varchar(50) = 'uSP_Sys_Advanced_Options'
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max) =''
  Declare @Err_Code Int = -1

  Set @Msg = 'advanced Option.'
  if @Option = 1
  begin
     Set @Msg = '[Enabled] '+@Msg
     begin try
       --[Enabled Option]
       Exec master..sp_configure 'show advanced options', 1
       Reconfigure with override

       Exec master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'AllowInProcess', 1
       Reconfigure with override

       Exec master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'DynamicParameters', 1
       Reconfigure with override
     
       Exec master..sp_configure 'xp_cmdshell', 1
       Reconfigure with override
  
       Exec master..sp_configure 'Ad Hoc Distributed Queries', 1
       Reconfigure with override

       Exec master..sp_configure 'Ole Automation Procedures', 1
       Reconfigure with override
     end try
     begin catch
       Set @Msg = @Msg+'.(錯誤訊息:'+ERROR_MESSAGE()+')' 
       Return @Err_Code
     end catch
  end
  else
  begin
     Set @Msg = '[Disabled] '+@Msg
     begin try
       --[Disabled Option]
       Exec master..sp_configure 'show advanced options', 1
       Reconfigure with override
     
       Exec master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'AllowInProcess', 0
       Reconfigure with override

       Exec master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'DynamicParameters', 0
       Reconfigure with override

       Exec master..sp_configure 'xp_cmdshell', 0
       Reconfigure with override

       Exec master..sp_configure 'Ad Hoc Distributed Queries', 0
       Reconfigure with override

       Exec master..sp_configure 'show advanced options', 0
       Reconfigure with override

       Exec master..sp_configure 'Ole Automation Procedures', 0
       Reconfigure with override
     end try
     begin catch
       Set @Msg = @Msg+'.(錯誤訊息:'+ERROR_MESSAGE()+')' 
       Return @Err_Code
     end catch

  end
  Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL

end
GO
