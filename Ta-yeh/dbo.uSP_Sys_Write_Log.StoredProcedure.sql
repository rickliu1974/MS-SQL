USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Write_Log]    Script Date: 08/18/2017 17:43:41 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Write_Log]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[uSP_Sys_Write_Log](@ProcName Varchar(50), @Msg Varchar(max)=Null, @SqlCmd Varchar(max)=Null, @Cnt Int=0, @SendMail Int=0, @Exec_Time Time = null)
as
Begin
  Declare @Proc Varchar(100) ='uSP_Sys_Write_Log'
  Declare @Errcode Int = -1
  Declare @CR Varchar(4) =Char(13)+Char(10)

  --print '[Show '+@ProcName+' Message Begin] '+Replicate('#', 100)
  if @Cnt = @Errcode
     set @Msg = @Msg + ' ...失敗'
  else
     set @Msg = @Msg + ' ...成功'

  -- 2015/01/22 Rickliu 增加寫入LOG時，可以轉發MAIL
  -- 時機點: @SendMail = 1 或者 @Cnt = -1 且 @SendMail 為 0, 1 時 才會發送郵件
  -- 當要寫入的LOG @Cnt 為 -1且又不要發送郵件時，只要將 @SendMail 設定不為 0, 1 的值即可。
  
  If @SendMail = 1 Or (@Cnt = @Errcode And @SendMail in (0, 1))
  begin
     Print 'Send Mail:'+@CR+@Msg
     If Isnull(@SqlCmd, '') = ''  
        Set @SqlCmd = @Msg
     Exec uSP_Sys_Send_Mail @ProcName, @Msg, '', @Msg, @SqlCmd
  end
     
  set @Msg = 'Host_Name:['+Host_Name()+'],'+
             'Host_AP:['+APP_NAME()+'],'+
             @Proc+'<-'+Isnull(@Msg, '無訊息傳入')+@CR
  --print 'Message:'+Replicate('#', 20)+@CR+@Msg
  --print @CR+'SQL Command:'+Replicate('#', 20)+@CR+@SqlCmd
  print @CR+'SQL Command:'+@CR+@SqlCmd
  --print '[Show '+@ProcName+' Message End  ] '+Replicate('#', 100)
  
  Insert Into Trans_Log
  Values(GETDATE(), @ProcName, @Msg, @SqlCmd, @Cnt, @Exec_Time)
  
End
GO
