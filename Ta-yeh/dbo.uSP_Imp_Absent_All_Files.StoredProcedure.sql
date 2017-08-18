USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_Absent_All_Files]    Script Date: 08/18/2017 17:43:40 ******/
DROP PROCEDURE [dbo].[uSP_Imp_Absent_All_Files]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*********************************************************************************
--�ҶԤU�Z�ɶ�
select * from  apsnt_bk where ps_date='2014/1/2'

update apsnt_bk
set ps_tm1e=b.ps_time
from  (
select ps_time,e_no from [dbo].[Ori_Txt#Apsnt_Tmp_1]  where  e_no='T09092' ) b
where ps_no collate Chinese_Taiwan_Stroke_CI_AS =e_no collate Chinese_Taiwan_Stroke_CI_AS
and ps_date='2013/12/13'
**********************************************************************************/

--Exec uSP_Imp_Apsnt_All_Files
CREATE Procedure [dbo].[uSP_Imp_Absent_All_Files]
  @b_date DateTime = Null, @e_Date DateTime = Null
as
begin
  -- �ۨ��{�Ǭ����]�w
  Declare @Proc Varchar(50) = 'uSp_Imp_Absent_All_Files'
  Declare @Cnt Int = 0
  Declare @RowCnt Int = 0
  Declare @Msg Varchar(4000) =''
  Declare @strSQL Varchar(Max)
 
  Declare @CR Varchar(5) = ' '+char(13)+char(10)
  Declare @RowCount Table (cnt int)
  Declare @sDB Varchar(50) = ''
  Declare @Errcode int = -1

  -- �ҶԬ����]�w
  Declare @Proce_SubName NVarchar(Max) = '�Ҧ��~���ҶԸ�ƶפJ�{��'
  
  -- �ɮ׬����]�w
  Declare @file_path varchar(255) = 'D:\Transform_Data\Import_Absent\'
  
  Declare @txt_head varchar(255) = ''
  Declare @Job_Name Varchar(255) = 'uSp_IMP_Absent'
  Declare @sDate Varchar(10) = ''
  
  if @b_Date is null 
     set @b_Date = GetDate()
  if @e_Date is null 
     set @e_Date = GetDate()

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  -- ����Ƶ{�A��
  select @Job_Name = name
    from msdb.dbo.sysjobs
   where name like '%Absent%'
  
  
  exec msdb.dbo.sp_update_job @job_name=@Job_Name, @enabled=0
  
  -- �إ߼Ȧs�榡��
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[#DirTree]') AND type in (N'U'))
     Drop Table #DirTree
     
  CREATE TABLE #DirTree 
   (Id int identity(1,1), 
    SubDirectory nvarchar(255),
    Depth smallint,
    FileFlag bit,
    ParentDirectoryID int 
   )

  UPDATE #DirTree 
     SET ParentDirectoryID =
         (SELECT MAX(Id) FROM #DirTree d2 
           WHERE Depth = d.Depth - 1 AND d2.Id < d.Id)
     FROM #DirTree d 

  -- �פJ�X�ԩҦb���Ҧ��ɮצW��
  Print 'File Path='+@file_path
  INSERT INTO #DirTree (SubDirectory, Depth, FileFlag)
  EXEC master.dbo.xp_dirtree @file_path, 10, 1
  
  -- �R�������n���ɮצW��
  Delete #DirTree 
   where depth <>'1' -- PS: �u�פJ�Ĥ@�h�ؿ����
      or (Isnumeric(Replace(subdirectory, '.Txt', '')) <>1)
      or (substring(subdirectory, 1, 1) in ('+', '-', '.'))
      or (Len(Replace(subdirectory, '.Txt', '')) <> 8)
      

  -- �j��妸�B�z�X�Ը��
  Declare Cur_Apsnt Cursor for
    Select Convert(Varchar(10), Convert(DateTime, Replace(Replace(Replace(Replace(subdirectory, '.Txt', ''), '+', ''), '-', ''), '.', '')), 111) as sDate
      from #DirTree
     order by subdirectory
      
  open Cur_Apsnt
  fetch next from Cur_Apsnt into @sDate

  while @@fetch_status =0
  begin
    if Not ((Convert(DateTIme, @sDate) >= @b_date) And (Convert(DateTIme, @sDate) <= @e_date))
    begin
       Print 'Skip '+@sDate
       fetch next from Cur_Apsnt into @sDate
       Continue
    end
    
    print 'Exec uSp_Imp_Absent ' + @sDate
    Exec uSp_Imp_Absent @sDate

    print 'Exec uSp_ETL_Absent ' + @sDate
    Exec uSp_ETL_Absent @sDate

    fetch next from Cur_Apsnt into @sDate
  end
  
  close Cur_Apsnt
  deallocate Cur_Apsnt
  -- �ҥαƵ{�A��
  exec msdb.dbo.sp_update_job @job_name=@Job_Name, @enabled=1

End_Exit:
  Return(@Cnt)
end
GO
