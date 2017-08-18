USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Exec_SQL]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Exec_SQL]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[uSP_Sys_Exec_SQL](@ProcName Varchar(50), @Msg Varchar(max)='', @strSQL Varchar(max)='', @SendMail Int= 0)
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_Sys_Exec_SQL
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/03 [因為有太多 Procedure 都使用外部呼叫 SQL 語法方式，所以新增此 Store Procedure 
                             來共同處理，並加以處理 Exception 機制避免 Log 遺漏。................RickLiu.
                 2015/02/27 這支程式是算底層程式，千萬不可以使用 Cursor，否則其他程式會造成無法使用 Cursor 的情況。
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  
  Declare @Proc Varchar(50) = 'uSP_Sys_Exec_SQL'
  Declare @Mstr Varchar(Max)
  Declare @Errcode Int = -1
  Declare @Sender Varchar(100) 
  Declare @subject Varchar(500)
  Declare @Error Varchar(50) = '【錯誤警告】'
  Declare @ErrMsg Varchar(Max)
  Declare @CR Varchar(4) =Char(13)+Char(10)
  Declare @BTime DateTime, @ETime DateTime
  Declare @Tempdb_Space_B Decimal=0, @Tempdb_Space_A Decimal=0
  Declare @Exec_Time Time
  Declare @Result Int = 0

  if Rtrim(Isnull(@ProcName, '')) = ''
     set @ProcName = 'Script'
print 'A:'+@Msg
  if Rtrim(Isnull(@Msg, '')) <> ''
     set @Msg = '...'+@Msg
  else
     set @Msg = ''

print 'B:'+@Msg

  if @SendMail = 1
     exec @Tempdb_Space_B = uSP_sys_Get_DB_Size 1

  Set @BTime = GetDate()
  
  if object_id('tempdb..#Msg') is not null 
     DROP TABLE #Msg

  Create Table #Msg(Msg Varchar(Max))

  if (RTrim(IsNull(@strSQL, '')) = '')
  begin
     Set @Result = @Errcode
     Set @ErrMsg = '空字串命令語法無法執行!!'
  end
  else
  begin 
     begin try
       print 'Exec SQL:'+@CR+@strSQL
       insert into #Msg exec (@strSQL)

       set @ErrMsg = @ProcName+' '+@Msg + @ErrMsg + '...'+Isnull((select Isnull(Msg, '')+', '+@CR from #Msg for xml path('')), '')
       set @Result = 0
       Set @Error = ''
     end try
     begin catch
     print 'catch:'+ERROR_MESSAGE()
        set @Result = @Errcode
        Set @ErrMsg = @Error+@ProcName+' 命令失敗!!'+Isnull(@Msg, '')+'...(錯誤訊息:'+Isnull(ERROR_MESSAGE(), '...無法取得 SQL Error_Message() 訊息')+')...(錯誤列:'+Convert(Varchar(10), ERROR_LINE())+')...(SQL 長度:'+Convert(Varchar(10), Len(@strSQL))+').'
     print @ErrMsg
     end catch
  end
 
  if @SendMail = 1 -- 若參數有設定強制發送則增加以下訊息
  begin
     exec @Tempdb_Space_A = uSP_sys_Get_DB_Size 1
     Set @ErrMsg = Isnull(@ErrMsg, '') + '..【Tempdb增長(MB):'+Convert(Varchar(10), isnull(@Tempdb_Space_A, 0) - isnull(@Tempdb_Space_B, 0))+'】'
  end

  Set @ETime = GetDate()
  Set @Exec_Time = Isnull(@ETime, getdate()) - Isnull(@BTime, getdate())

  If @Result = @Errcode
  begin
     Set @SendMail  = 1 -- 若發生執行錯誤則強制發送訊息
     set @ErrMsg = Isnull(@ErrMsg, '')+'...【執行時間：'+Convert(Varchar(20), @Exec_Time, 114)+'】'
  end
  else
     set @ErrMsg = 'Exec '+isnull(@ProcName, '')+Isnull(@Msg, '')+Isnull(@ErrMsg, '')+'...【執行時間：'+Convert(Varchar(20), @Exec_Time, 114)+'】'

  Exec uSP_Sys_Write_Log @ProcName, @ErrMsg, @strSQL, @Result, @SendMail, @Exec_Time
  
  Print @ErrMsg
  Return(@Result)
end
GO
