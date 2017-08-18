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
   Updated Date: 2013/09/03 [�]�����Ӧh Procedure ���ϥΥ~���I�s SQL �y�k�覡�A�ҥH�s�W�� Store Procedure 
                             �Ӧ@�P�B�z�A�å[�H�B�z Exception �����קK Log ��|�C................RickLiu.
                 2015/02/27 �o��{���O�⩳�h�{���A�d�U���i�H�ϥ� Cursor�A�_�h��L�{���|�y���L�k�ϥ� Cursor �����p�C
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  
  Declare @Proc Varchar(50) = 'uSP_Sys_Exec_SQL'
  Declare @Mstr Varchar(Max)
  Declare @Errcode Int = -1
  Declare @Sender Varchar(100) 
  Declare @subject Varchar(500)
  Declare @Error Varchar(50) = '�i���~ĵ�i�j'
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
     Set @ErrMsg = '�Ŧr��R�O�y�k�L�k����!!'
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
        Set @ErrMsg = @Error+@ProcName+' �R�O����!!'+Isnull(@Msg, '')+'...(���~�T��:'+Isnull(ERROR_MESSAGE(), '...�L�k���o SQL Error_Message() �T��')+')...(���~�C:'+Convert(Varchar(10), ERROR_LINE())+')...(SQL ����:'+Convert(Varchar(10), Len(@strSQL))+').'
     print @ErrMsg
     end catch
  end
 
  if @SendMail = 1 -- �Y�ѼƦ��]�w�j��o�e�h�W�[�H�U�T��
  begin
     exec @Tempdb_Space_A = uSP_sys_Get_DB_Size 1
     Set @ErrMsg = Isnull(@ErrMsg, '') + '..�iTempdb�W��(MB):'+Convert(Varchar(10), isnull(@Tempdb_Space_A, 0) - isnull(@Tempdb_Space_B, 0))+'�j'
  end

  Set @ETime = GetDate()
  Set @Exec_Time = Isnull(@ETime, getdate()) - Isnull(@BTime, getdate())

  If @Result = @Errcode
  begin
     Set @SendMail  = 1 -- �Y�o�Ͱ�����~�h�j��o�e�T��
     set @ErrMsg = Isnull(@ErrMsg, '')+'...�i����ɶ��G'+Convert(Varchar(20), @Exec_Time, 114)+'�j'
  end
  else
     set @ErrMsg = 'Exec '+isnull(@ProcName, '')+Isnull(@Msg, '')+Isnull(@ErrMsg, '')+'...�i����ɶ��G'+Convert(Varchar(20), @Exec_Time, 114)+'�j'

  Exec uSP_Sys_Write_Log @ProcName, @ErrMsg, @strSQL, @Result, @SendMail, @Exec_Time
  
  Print @ErrMsg
  Return(@Result)
end
GO
