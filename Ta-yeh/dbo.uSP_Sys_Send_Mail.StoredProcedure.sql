USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Send_Mail]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Send_Mail]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Proc [dbo].[uSP_Sys_Send_Mail]
  @Run_ProcName Varchar(500)='',
  @subject Varchar(500)='',
  @Sender Varchar(500)='',
  @Msg Varchar(Max) ='',
  @strSQL Varchar(Max) =''
as
begin
  Declare @Proc Varchar(50) = 'uSP_Sys_Send_Mail'
  Declare @Now Varchar(100) =''
  Declare @CR Varchar(4) =Char(13)+Char(10)
  
  Set @Now = Convert(Varchar(30), Getdate(), 121)

  Begin Try
    -- 2015/01/22 Rickliu �o�{�l��D�������\�_��Ÿ��A�_�h�|�X�{�l��i�H�o�X�h���O�|�����쪺���p�C
    Set @Subject = Replace(Replace(Isnull(@Subject, ''), Char(13), ''), Char(10), '')
    --set @Sender = 'it@ta-yeh.com.tw;rickliu.1974@gmail.com;nanliao1982@gmail.com;'+Isnull(@Sender, '')
    set @Sender = 'it@ta-yeh.com.tw;'+Isnull(@Sender, '')
    set @Msg = Isnull(@Msg, '')
    
    Set @Msg = N'<H2>�{���I�s�G'+@Run_ProcName+'</H2>'+
               N'<H2>�D�����e�G'+@Subject+'</H2>'+
               N'<H2>�o�e�ɶ��G'+@Now+'</H2>'+
               N'<H2>�o�e��H�G'+@Sender+'</H2>'+
               N'<H2>�T�����e�G</H2>'+
               Convert(Varchar(Max), Replace(@Msg, Char(13)+Char(10), '<BR>'))+'<BR><BR>'
             
    If @strSQL <> ''
       Set @Msg = @Msg +
                  N'<H2>SQL Command / RowData�G</H2>'+
                  Case
                    when @strSQL like '%<td%' then
                      N'<style type="text/css">'+
                      N' table {'+
                      N'   width="640";'+
                      N'   margin: 10px;'+
                      N'   border-collapse: collapse;'+
                      N' }'+
                      N' table td {'+
                      N'	border-width: 1px;'+
                      N'	padding: 8px;'+
                      N'	border-style: solid;'+
                      N'	border-color: red;'+
                      N' }'+
--                      N' table.horizontal td:first-child {'+
--                      N'	background-color: #000000!important;'+
--                      N'	font-weight: bold;'+
--                      N'	color: red;'+
--                      N' }'+
--                      N' table.vertical tr>td:first-child {'+
--                      N'	background-color: Gray!important;'+
--                      N'	font-weight: bold;'+
--                      N'	color: #fff;'+
--                      N' }'+
                      N' tr {'+
                      N'   background-color:#000000;'+
                      N'   font-style:normal;'+
                      N' }'+
                      N'</style>'+
                      N'<table class="horizontal">'
                    else ''
                  end+
                  '<Pre>'+@strSQL+'</Pre>'
                  --'<Pre>'+Convert(Varchar(Max), Replace(@strSQL, Char(13), '<BR>'))+'</Pre>'

    Print @CR+@CR+
          'EXECUTE msdb.dbo.sp_send_dbmail '+@CR+
          ' @profile_name=''VMSOP_Alert'','+@CR+
          ' @recipients='''+@Sender+''','+@CR+
          ' @subject='''+@subject+''','+@CR+
          ' @body='''+@Msg+''','+@CR+
          ' @body_format =''HTML'' '+@CR+@CR

    EXECUTE msdb.dbo.sp_send_dbmail
            @recipients=@Sender,
            @subject= @subject,
            @body=@Msg,
            @body_format = 'HTML'

    Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
  end try
  begin catch
    Set @Msg = N'<H2>�o�e�ɶ��G'+@Now+'</H2><BR><BR>'+
               N'<H2>�ѼƩR�O�G</H2><BR>'+
               N'<H2>@Run_ProcName = '+Convert(Varchar(1000), @Run_ProcName)+'</H2><BR>'+
               N'<H2>@Subject = '+Convert(Varchar(1000), @subject)+'</H2><BR>'+
               N'<H2>@Sender = '+Convert(Varchar(1000), @Sender)+'</H2><BR>'+
               N'<H2>@Msg = '+Convert(Varchar(Max), Replace(@Msg, Char(13), '<BR>'))+'</H2><BR>'+
               N'<H2>@strSQL = <Pre>'+@strSQL+'</Pre></H2><BR><BR>'+
               N'<H2>���~�T���G</H2><BR>'+
               ERROR_MESSAGE()

    EXECUTE msdb.dbo.sp_send_dbmail
            @recipients=@Sender,
            @subject= @subject,
            @body=@Msg,
            @body_format = 'HTML'
    Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
  end catch
end
GO
