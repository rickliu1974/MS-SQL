USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Clean_Log]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Clean_Log]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_Sys_Clean_Log](@Kind Varchar(1) = '1', @MM Int = 3)
as
begin
  Declare @Proc Varchar(50) = 'uSP_Sys_Clean_Log'
  Declare @Cnt Int = 0
  Declare @RowCnt Int = 0
  Declare @Msg Varchar(4000) =''
  Declare @Errcode int = -1

  Declare @strSQL Varchar(Max)
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @RowCount Table (cnt int)
  Declare @RowData Table (aData Varchar(255))
  
  
  Declare @RowID Int
  Declare @ProcName Varchar(50) = '�M����Ʈw�{�ǰO����'

  Declare @YM Varchar(10) = '', @delYM Varchar(7) = Convert(Varchar(7), Dateadd(month, -@MM, Getdate()), 111)
  Declare @Kind_Name Varchar(50)= '', @Table_Name Varchar(100) = '', @Column_Name Varchar(50) = ''
  Declare @YM_Cnt Int
  
  Declare @TempTable Table (
    Kind Varchar(1),
    Kind_Name Varchar(50),
    Table_Name Varchar(50),
    Column_Name Varchar(50),
    YM varchar(7),
    YM_Cnt Int
  )
  
  Print '--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*'
  Print '�Ѽƻ����G'
  Print '@Kind: �R�� Log ����'
  Print '       1: Trans_Log[�w�]].'
  Print '       2: DB_MAIL.'
  Print '       3: DB_False_Mail.'
  Print '@MM: �H�t�Φ~��A�O�d�̪�X�Ӥ��ơA�w�] 3 �Ӥ�C'
  Print '--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*'
  Print ''
 
  Insert into @TempTable
  Select '1' as Kind,
         'Trans_Log' as Kind_Name,
         'dw.dbo.trans_log' as Table_Name,
         'Trans_date' as Column_Name, 
         Convert(Varchar(7), Trans_date, 111) YM, 
         Count(Convert(Varchar(7), Trans_date, 111)) as YM_Cnt
    from dw.dbo.trans_log With (NoLock, NoWait)
   where 1 = @Kind
   group by Convert(Varchar(7), Trans_date, 111) 
  union
  Select '2' as Kind,
         'DB_MAIL' as Kind_Name,
         'msdb.dbo.sysmail_mailitems' as Table_Name,
         'send_request_date' as Column_Name, 
         Convert(Varchar(7), send_request_date, 111) YM, 
         Count(1) as YM_Cnt
    from msdb.dbo.sysmail_mailitems With (NoLock, NoWait)
   where 2 = @Kind
   group by Convert(Varchar(7), send_request_date, 111) 
  union
  Select '3' as Kind,
         'DB_False_Mail' as Kind_Name,
         'msdb.dbo.sysmail_faileditems' as Table_Name,
         'send_request_date' as Column_Name, 
         Convert(Varchar(7), send_request_date, 111) YM, 
         Count(1) as YM_Cnt
    from msdb.dbo.sysmail_faileditems With (NoLock, NoWait)
   where 3 = @Kind
   group by Convert(Varchar(7), send_request_date, 111) 
 
  Declare Cur_Clean_Log cursor for
    Select *
      from @TempTable
     where 1=1
       and Kind = @Kind
       and YM <= @delYM
       
  open Cur_Clean_Log
  fetch next from Cur_Clean_Log into @Kind, @Kind_Name, @Table_Name, @Column_Name, @YM, @YM_Cnt

  If @@fetch_status = 0
  begin
     Set @Msg = @Proc+'�i��M�� '+@Kind+'.'+@Kind_Name+' �t�Τ���e [ '+Convert(Varchar(10), @MM)+ ' ] �Ӥ���!!!'+@CR
     Select Kind_Name, YM, YM_Cnt from @TempTable where 1=1 and Kind =@Kind and YM <= @delYM Order by YM
     Exec uSP_Sys_Send_Mail @Proc, @Msg, '', 1
  end
  else 
  begin
     Set @Msg = @Proc+'...�õL @Kind='+@Kind+', [ '+Convert(Varchar(10), @MM)+ ' ] �Ӥ�e��ƵL���M��!!!'+@CR
     Select Kind_Name, YM, YM_Cnt from @TempTable where 1=1 and Kind =@Kind Order by YM
     Exec uSP_Sys_Send_Mail @Proc, @Msg, '', 1
  end
  
  print @Msg
  while @@fetch_status =0
  begin
    Set @Msg = '�M�� '+@Kind+'.'+@Kind_Name+' ['+@YM+'] ����['+Convert(Varchar(100), @YM_Cnt)+']'
    Set @strSQL = 'Declare @Deleted_Rows int'+@CR+
                  'Set @Deleted_Rows = 1'+@CR+
                  ''+@CR+
                  'While (@Deleted_Rows >0) '+@CR+
                  'Begin '+@CR+
                  '  begin Transaction '+@CR+
                  '    Delete Top(100000) '+@Table_Name+' '+@CR+ -- �C 100000 Commit �@��
                  '     where Convert(Varchar(7), '+@Column_Name+', 111) <= '''+Convert(Varchar(100), @YM)+''' '+@CR+
                  '    Set @Deleted_Rows = @@RowCount '+@CR+
                  '    Print ''Delete '+@Table_Name+' On YM:['+@YM+'] Rows:[''+Convert(Varchar(100), @Deleted_Rows)+''].'' '+@CR+
                  '  Commit Transaction '+@CR+
                  -- '  CheckPoint '+@CR+ -- for simple recovery model
                  'end '

    Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL, 1

    fetch next from Cur_Clean_Log into @Kind, @Kind_Name, @Table_Name, @Column_Name, @YM, @YM_Cnt
  end

  close Cur_Clean_Log
  deallocate Cur_Clean_Log
  Return(0)
end
GO
