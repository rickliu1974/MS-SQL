USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Sys_Kill_LY_Process]    Script Date: 08/18/2017 17:43:41 ******/
DROP PROCEDURE [dbo].[uSP_Sys_Kill_LY_Process]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[uSP_Sys_Kill_LY_Process]
WITH EXEC AS CALLER
AS
begin
  Declare @Proc Varchar(50) 
  Declare @Sender Varchar(100) 
  Declare @CR Varchar(4) 
  Declare @Body_Msg NVarchar(4000) 
  Declare @Run_Proc_Msg NVarchar(4000) 
  Declare @Run_Msg NVarchar(100)
  Declare @Proc_Name NVarchar(4000)
  Declare @Cnt Int = 0
  Declare @strSQL Varchar(Max)
  Declare @RowCount Table (cnt int)
  
  Declare @soft Varchar(100), @worknum int, @workname Varchar(100), 
                              @e_no Varchar(10), @e_name Varchar(20), 
                              @dp_no Varchar(10), @dp_name Varchar(20), 
                              @workdate DateTime, @last_batch Varchar(20), @nonuse_time int, 
                              @remark varchar(50)
  Declare @SendTime Varchar(20)

  Set @Proc = 'uSP_Sys_Kill_LY_Process'
  Set @Sender = 'it@ta-yeh.com.tw'
  Set @Cr =Char(13)+Char(10)
  Set @Body_Msg = ''
  Set @Run_Proc_Msg = ''
  Set @Run_Msg = ''
  Set @Proc_Name = '��V�t�ιO�ɤH���M���{��'
  
  
  -- 2015/02/04 Rickliu �ѩ�L�k�o�� AutoReIndex �۰ʭ����ɰ���A�ҥH�@�ߧP�O���V����X�гQ�]�߮ɡA�h���i��尣�s�u�@�~
  -- 0: ������, 1:���㤤
  set @strSQL =(select 'Union Select Convert(Int, l_field2) as l_field2 from '+Server_Name+'.'+DataBase_Name+'.dbo.ssyslb # '
                  from Sys_Ori_Databases
                 where Company_Name <> 'ALL'
                   for xml path('')
               )
  set @strSQL = Replace(@strSQL, '#', @CR)
  Print 'SQL:'+@strSQL

  set @Cnt = 0
  set @strSQL = 'select Sum(l_field2) as cnt '+@CR+
                '  from ('+@CR+
                Substring(@strSQL, 7, Len(@strSQL))+
                '       ) m'
  Print 'SQL:'+@strSQL

  delete @RowCount
  insert into @RowCount Exec (@strSQL)
  select @Cnt=cnt from @RowCount
  
  print 'Count:'+Convert(Varchar, @Cnt)
  
  if @Cnt <> 0
  begin
    set @Body_Msg = '�������A�]��V�t�Υثe���㤤!!'
    Exec uSP_Sys_Write_Log @Proc, @Body_Msg, @Body_Msg, 0
  end
  begin
    IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[#tmp_Process]') AND type in (N'P', N'PC'))
       DROP PROCEDURE [dbo].[#tmp_Process]
    
    select soft, worknum, workname, e_no, e_name, dp_no, dp_name, workdate, last_batch, nonuse_time, remark
      into #tmp_Process
      from (select distinct soft, worknum, workname, e_no, e_name, dp_no, dp_name, workdate, last_batch, nonuse_time, 
                   case 
                     when last_batch = '�q���_���u'
                     then '�q���_���u'
                     else '�O�ɧR��' 
                   end as remark
              from TYEIPDBS2.lytdb.dbo.V_LYT_Connent M
             where 1=1
               and first_kill_dp='V'
               and ((nonuse_time >= 40 and lock = '' and workname not like '%:::%')
                or  (isnull(last_batch, '') = '' And datediff(mi, workdate, getdate()) >= 40 And lock = ''))
               -- or last_batch = '�q���_���u'
               --and softkind=1
               --and Cnt = 20
            union
            -- 2014/10/15 Rickliu �� MRP �s�u�ƺ��ɡA�B�u���R�������s�u�ƶW�L 3 �H�ɡA�h���u���R���C
            select distinct m.soft, worknum, m.workname, m.e_no, m.e_name, m.dp_no, m.dp_name, m.workdate, m.last_batch, m.nonuse_time, '�H�ƹL��' as remark
              from TYEIPDBS2.lytdb.dbo.V_LYT_Connent M
                   inner join 
                   (select distinct substring(dp_no, 1, 2) as dp_no, max(nonuse_time) as nonuse_time, count(1) as cnt
                      from TYEIPDBS2.lytdb.dbo.V_LYT_Connent M
                     where 1=1
                       and first_kill_dp='V'
                       and ((cnt = 20 and softkind = 1 and lock = '' and workname not like '%:::%')
                        or  (isnull(last_batch, '') = '' And datediff(mi, workdate, getdate()) >= 40 And lock = ''))
                     group by substring(dp_no, 1, 2)
                    having count(*) > 2
                   ) d
                    on substring(m.dp_no, 1, 2) = substring(d.dp_no, 1, 2)
                   and m.nonuse_time = d.nonuse_time
            ) m
          
    Declare Cur_Process Cursor for
      select soft, worknum, workname, e_no, e_name, dp_no, dp_name, workdate, last_batch, nonuse_time, remark
        from #tmp_Process
        
    open Cur_Process
    fetch next from Cur_Process into @soft, @worknum, @workname, @e_no, @e_name, @dp_no, @dp_name, @workdate, @last_batch, @nonuse_time, @remark
    
    Set @SendTime = Convert(Varchar(30), Getdate(), 121)
  
    Set @Body_Msg = N'<H1>'+@Proc_Name+'</H1>'+
                    N'<H2>�o�e�ɶ��G'+@SendTime+'</H2>'+
                    N'<table border="1">' +
                    N'<tr>'+
                    N'  <th>�t�ε{��</th>'+
                    N'  <th>�����</th>'+
                    N'  <th>�u�@��</th>'+
                    N'  <th>���u�s��</th>'+
                    N'  <th>���u�m�W</th>'+
                    N'  <th>�����s��</th>'+
                    N'  <th>�����W��</th>'+
                    N'  <th>�W�u�ɶ�</th>'+
                    N'  <th>�̪�ާ@�ɶ�</th>'+
                    N'  <th>�O�ɮɶ�</th>'+
                    N'  <th>�R����]</th>'+
                    N'</tr>'
  
    while @@fetch_status =0
    begin
      Set @Body_Msg = @Body_Msg +
                      N'<tr>'+
                      N'<td>'+isnull(@Soft, '')+N'</td>'+
                      N'<td>'+isnull(Convert(Varchar(20), @worknum), '')+N'</td>'+
                      N'<td>'+Isnull(@workname, '')+N'</td>'+
                      N'<td>'+Isnull(@e_no, '')+N'</td>'+
                      N'<td>'+Isnull(@e_name, '')+N'</td>'+
                      N'<td>'+Isnull(@dp_no, '')+N'</td>'+
                      N'<td>'+Isnull(@dp_name, '')+N'</td>'+
                      N'<td>'+Isnull(Convert(Varchar(20), @workdate, 121), '')+N'</td>'+
                      N'<td>'+Isnull(Convert(Varchar(20), @last_batch, 121), '')+N'</td>'+
                      N'<td>'+Isnull(Convert(Varchar(20), @nonuse_time), '')+N'</td>'+
                      N'<td>'+Isnull(Convert(Varchar(20), @remark), '')+N'</td>'+
                      N'</tr>'
  
      delete TYEIPDBS2.lytdb.dbo.sqlnloc where worknum = @worknum
  
      fetch next from Cur_Process into @soft, @worknum, @workname, @e_no, @e_name, @dp_no, @dp_name, @workdate, @last_batch, @nonuse_time, @remark
    end
    Set @Body_Msg = @Body_Msg + N'</table>'
  
    close Cur_Process
    deallocate Cur_Process
    
    set @Proc_Name = @Proc+' '+@Proc_Name+'...�o�e�ɶ��G'+@SendTime
    if exists (select * from #tmp_Process)  
    begin
       EXECUTE msdb.dbo.sp_send_dbmail
               @recipients=@Sender,
               @subject= @Proc_Name,
               @body=@Body_Msg,
               @body_format = 'HTML'
       Exec SP_Write_Log @Proc, @Body_Msg, '', 0
    end
  end
end
GO
