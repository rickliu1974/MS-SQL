USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_Imp_Absent]    Script Date: 08/18/2017 17:43:40 ******/
DROP PROCEDURE [dbo].[uSP_Imp_Absent]
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

--Exec uSP_Imp_Apsnt
CREATE Procedure [dbo].[uSP_Imp_Absent]
  @ps_date datetime = null
as
begin
  /*************************************************************************************************************
   2013/12/21 �P����z�T�{�A�P�@�ܶg���ΩP�鳣�]�w�����`�Z�ɬq�A�P���h�U�Z�����@�p�ɡA�X�ԶȬݤW�Z�ΤU�Z�A
              �[�Z�B�~�X�����z�L�d���P�O�C
              �W�Z�Ȭ�P�̦��X�Ԯɶ��A�U�Z�h�P�O�̱ߤU�Z�ɶ��C
              ���`�Z�G07:00~17:45, �P���Z�G07:00~16:45
  *************************************************************************************************************/
  -- �ۨ��{�Ǭ����]�w
  Declare @Proc Varchar(50) = 'uSP_Imp_Absent'
  Declare @Cnt Int = 0
  Declare @RowCnt Int = 0
  Declare @Msg Varchar(4000) =''
  Declare @strSQL Varchar(Max)
  --20131228 add by brian �W�[�U�Z�ɶ�
  Declare @strSQL1 Varchar(Max)
 
  Declare @CR Varchar(5) = ' '+char(13)+char(10)
  Declare @RowCount Table (cnt int)
  --Declare @rDB Varchar(50) = 'TYEIPDBS2.lytdbTA13.dbo.'
  Declare @rDB Varchar(50) = 'SYNC_TA13.dbo.' -- 2017/03/17 Rickliu �אּ�H�q�\��Ʈw���ѦҸ��
  Declare @wDB Varchar(50) = 'TYEIPDBS2.lytdbTA13.dbo.' -- 2017/03/17 Rickliu �s�W�^�g�ت���Ʈw
  
  --set @rDB = @wDB -- Rickliu 2017/03/24 �b�q�\�A���٨S�n�ɡA�Ȯɨϥ� @wDB ��Ʈw�����
  
  Declare @Errcode int = -1

  -- �ҶԬ����]�w
  Declare @Proce_SubName NVarchar(Max) = '�ҶԸ�ƶפJ�{��'
  Declare @Body_Msg NVarchar(Max) = ''
  Declare @Run_Msg NVarchar(Max) = ''
  
  -- �ɮ׬����]�w
  Declare @file_Name varchar(50) = ''
  Declare @file_ext varchar(10) = '.Txt'
  Declare @file_path varchar(255) = 'D:\Transform_Data\Import_Absent\'
  Declare @isExists Int -- �PŪ�ɮ׬O�_�s�b 1: �s�b, 2: ���s�b
  
  -- Declare @txt_head varchar(255) = 'Ori_Txt#' 
  Declare @txt_head varchar(255) = '' -- 2017/03/17 Rickliu �אּ�H�q�\��Ʈw���ѦҸ��
  Declare @tb_Name varchar(255) ='Apsnt'
  Declare @tb_Ori_Apsnt varchar(255) = @txt_head + @tb_Name -- Ex: Ori_Txt#Apsnt -- Text �X����J������
  Declare @tb_Apsnt_tmp varchar(255) = @tb_Ori_Apsnt + '_Tmp' -- Ex: Ori_Txt#Apsnt_Tmp -- �X�Ը�ƳB�z�Ȧs��
  
  -- �ҶԮɶ������]�w
  Declare @Time_Over Varchar(5) = '00:00' -- �ҶԸ�]�ɶ�
  
  Declare @Time1 Varchar(5) = '07:00' -- �W�Z�_��ɶ�
  Declare @Car1 Varchar(1) = '1' -- �W�Z�s��

  Declare @Time2 Varchar(5) = '18:00' -- �U�Z�����ɶ�
  Declare @Car2 Varchar(1) = '2' -- �U�Z�s��
  
  Declare @Time3 Varchar(5) = '19:00' -- �[�Z�_��ɶ�
  Declare @Car3 Varchar(1) = '3' -- �[�Z�W�Z�s��

  Declare @Time4 Varchar(5) = '06:29' -- �[�Z�����ɶ�
  Declare @Car4 Varchar(1) = '4' -- �[�Z�U�Z�s��
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @ps_date = isnull(@ps_date, getdate())
  set @File_Name = @File_path + Convert(varchar(12), @ps_date, 112) + @File_ext -- Ex: D:\�Ҷ�����\20131210.Txt
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*

  --Exec uSP_Advanced_Options 1
  
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@tb_Ori_Apsnt+']') AND type in (N'U'))
  begin
     set @Msg = '1.�M�� �X����J������['+@tb_Ori_Apsnt+'].'
     set @Cnt = 0
     set @strSQL = 'Drop Table '+@tb_Ori_Apsnt
     Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
     if @Cnt = @Errcode Goto End_Proc
  end

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Msg = '2.�إ� �X����J������['+@tb_Ori_Apsnt+'].'
  set @Cnt = 0
  set @strSQL = 'create table '+@tb_Ori_Apsnt +@Cr+
                '  (e_no Varchar(10),'+@Cr+
                '   ps_week varchar(10),'+@Cr+
                '   ps_date datetime,'+@Cr+
                '   ps_time varchar(10),'+@Cr+
                '   ps_class int,'+@Cr+
                '   e_name varchar(20), '+@Cr+
                '   Door_No varchar(5),  '+@Cr+ -- Rickliu 2016/12/21 �s�W�d���s���P�W��
                '   Door_Name varchar(50)) '
  Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  if @Cnt = @Errcode Goto End_Proc
  
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Msg = '3.�PŪ�Ҷ��ɮ�['+@File_Name+']�O�_�s�b.'
  set @Cnt = 0
  exec master.dbo.xp_fileexist @File_Name, @isExists OUTPUT
  if @isExists <> 1
  begin
     set @Cnt = @Errcode
     set @Msg = @Msg + '...�ɮפ��s�b.'
     Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
     if @Cnt = @Errcode goto End_Proc
  end
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Msg = '4.����פJ �X�Ը�Ʀ���J������ ['+@File_Path+'] ==> ['+@tb_Ori_Apsnt+'].'
  set @Cnt = 0
  set @strSQL = 'BULK INSERT '+@tb_Ori_Apsnt +@Cr+
                '       FROM '''+@File_Name+''' WITH(FIELDTERMINATOR='' '' , ROWTERMINATOR=''\n'', TABLOCK )'
 
  Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  if @Cnt = @Errcode Goto End_Proc
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Msg = '5.�P�_ ['+@tb_Ori_Apsnt+'] �O�_�����.'
  set @Cnt = 0
  set @strSQL = 'select Count(1) as Cnt From '+@tb_Ori_Apsnt
  
  delete @RowCount
  insert into @RowCount exec (@strSQL)
  set @Cnt =(select Cnt from @RowCount)
  set @Msg = @Msg + '...�פJ���� ['+cast(@Cnt as varchar)+']'
  
  If @Cnt = 0 
  begin
     Set @Msg = @Msg + '..�L��ơA��������!!'
     Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
     Goto End_Exit
  end
  Exec uSP_Sys_Write_Log @Proc, @Msg, '', 0
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].['+@tb_Apsnt_tmp+']') AND type in (N'U'))
  begin
     set @Msg = '6.���� �X�Ը�ƳB�z�Ȧs��['+@tb_Apsnt_tmp+'].'
     set @Cnt = 0
     set @strSQL = 'Drop Table '+@tb_Apsnt_tmp
     Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end
  if @Cnt = @Errcode Goto End_Proc
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Msg = '7.���� �X�Ը�ƳB�z�Ȧs��['+@tb_Apsnt_tmp+'].'
  set @Cnt = 0
  -- 2014/07/09 Rickliu �]�X�Ըɥ���ҥH�|����d�����A�Ӭ��F�O�d��l��d�����A��H�x�s�� ps_tm3s, ps_tm3e, 
  --                    �{���Y�P�_�� ps_tm4s, ps_tm4e ���ɶ��h�N���i��X�Ըɥ��ʧ@�A�ҥH�|���N�� ps_tm1s, ps_tm2s
  /***********************************************************************************************************************
  -- 2017/02/09 Rickliu �]���q���F�]���F���@�Ҥ@��F���A�B�D�i�����y���u�[�Z�A�U�Z�ɶ���F�N���ǮɤU�Z�A���S�L�k�z�L�H���U
                        �ɭ��u�Y�ɤU�Z�B�S�`�ȳҰʳ���d���p�C�]�������ƪ��Ψ��`�g�z���ܭn�D�A�z�L�t�ε{������׭q���U�Z�H
                        ���ӥ��ǮɤU�Z�̤@�v�׭q�X�ԤU�Z�ɶ��� 18:00�C
                        ��T�Ҫ� RICKLIU ���z�L�h���|ĳ�����A���q�ק�X�Ԭ����N�|Ĳ�ǰ��y��Ѥεn�������Ѹo�A���|�U���i
                        ��A�����h����w�n�D���@�k�A�]���A�����]�u��L�`�t�X���q�F���C
  ***********************************************************************************************************************/
  
  set @strSQL = ';With CTE_Q1 as ( '+@Cr+
                '   select e_no, e_name, '+@Cr+
                '          ps_date, '+@Cr+
                '          min(ps_time) as ps_time, '+@Cr+
                '          case '+@Cr+
                '            when datediff(mi, min(ps_time), max(ps_time)) < 30 then null '+@Cr+
                '            else max(ps_time) '+@Cr+
                '          end as ps_time2, '+@Cr+
                '          case '+@Cr+
                '            when max(ps_time) between ''00:00'' and ''07:00'' then max(ps_time) '+@Cr+
                '            else null '+@Cr+
                '          end as ps_time3 '+@Cr+
                '    from '+@tb_Ori_Apsnt+@Cr+
                '   group by e_no, e_name, ps_date '+@Cr+
                '), CTE_Q2 as ( '+@Cr+
                '  select distinct e_no, e_name, '+@Cr+
                '         substring(datename(weekday, ps_date), 3, 1) as ps_week, '+@Cr+
                '         ps_date, '+@Cr+
                '         max(ps_time) as ps_time, '+@Cr+
                '         max(ps_time2) as ps_time2, '+@Cr+
                '         max(ps_time3) as ps_time3 '+@Cr+
                '    from CTE_Q1 '+@Cr+
                '   group by e_no, e_name, ps_date '+@Cr+
                '), CTE_Q3 as ('+@Cr+
                '  select distinct e_no, e_dept '+@Cr+
                '    from '+@rDB+'pemploy '+@Cr+
                '), CTE_Q4 as ('+@Cr+
                '  select ps_date, min(ps_time) as ps_min_time, max(ps_time) as ps_max_time '+@Cr+
                '    from '+@tb_Ori_Apsnt+@Cr+
                '   group by ps_date '+@Cr+
                ') '+@Cr+
                ' '+@Cr+
                'select distinct m.ps_date as A_ps_date, m.ps_week, m.e_no as e_no, m.e_name, '+@Cr+
                -- 2017/02/09 Rickliu ���u�W�Z���d����
                '       case  '+@Cr+
                '         when (Replace(isnull(d1.ps_tm4s, ''''), '':'', '''') <> '''') then Convert(Varchar(5), Convert(Time, Replace(Rtrim(d1.ps_tm4s), '':'', ''''))) '+@Cr+
                '         else isnull(m.ps_time, '''') '+@Cr+
                '       end as ps_time_s, '+@Cr+
                -- 2017/02/09 Rickliu ���u�U�Z���d����
                '       case '+@Cr+
                -- 2017/02/09 Rickliu �Y�H�Ƴ�즳��g��V�X�԰O�����@���j�]�Z�U�Z���h�N���U�Z�X�Ըɥ��ʧ@�A�t�Ϋh�^��X�Ըɥ����U�Z�ɶ��C
                '         when isnull(m.ps_time2, '''')= '''' And (Replace(isnull(d1.ps_tm4e, ''''), '':'', '''') <> '''') then Convert(Varchar(5), Convert(Time, Replace(Rtrim(d1.ps_tm4e), '':'', ''''))) '+@Cr+
                -- 2017/02/09 Rickliu �Y�W�L 18:30 �����d�̡A�t�Τ@�v�۰ʳ]�w 18:30 ���d
                '         when (Replace(isnull(isnull(isnull(m.ps_time2, d1.ps_tm3e), d1.ps_tm4e), ''''), '':'', '''') = '''') And (Convert(Varchar(5), Convert(Time, getdate())) >= ''18:30'') then ''18:30'' '+@Cr+
                -- 2017/02/09 Rickliu ���u�Y�w�g�����d�B�ɶ��W�L 18:30 �̡A�t�Τ@�v�۰ʳ]�w 18:30 ���d�C
                '         when (Convert(Varchar(5), Convert(Time, Replace(Rtrim(d1.ps_tm3e), '':'', ''''))) > ''18:30'') then ''18:30'' '+@Cr+
                -- 2017/02/09 Rickliu �L�U�Z�d�B�L��l��d�B�L�X�Ըɥ��A���Ӥ餽�q�T���X�Ԭ����A�h�N��ӭ��ѰO���d�A�]�������^�� 18:30
                '         when (Replace(isnull(isnull(isnull(m.ps_time2, d1.ps_tm3e), d1.ps_tm4e), ''''), '':'', '''') = '''') and (isnull(ps_max_time, '''') >= ''18:30'') then ''18:30'' '+@Cr+
                -- 2017/02/09 Rickliu �H�W���󳣤��ŦX�h�O�d���
                '         else isnull(m.ps_time2, '''') '+@Cr+
                '       end as ps_time_e, '+@Cr+
                -- 2017/02/09 Rickliu �W�Z�ɶ�2
                '       isnull(m.ps_time3, '''') as ps_time2_s, '+@Cr+
                -- 2017/02/09 Rickliu �U�Z�ɶ�2
                '       '''' as ps_time2_e, '+@Cr+
                -- 2017/02/09 Rickliu ���O�d��l�W�Z���d����
                '       isnull(m.ps_time, '''') as ps_time3_s, '+@Cr+
                -- 2017/02/09 Rickliu ���O�d��l�U�Z���d����
                '       isnull(m.ps_time2, '''') as ps_time3_e, '+@Cr+
                -- 2017/02/09 Rickliu �W�Z�ɶ�4(�X�Ըɥ��椧�W�Z���d����)
                '       case  '+@Cr+
                '         when (Replace(isnull(d1.ps_tm4s, ''''), '':'', '''') <> '''') then Convert(Varchar(5), Convert(Time, Replace(Rtrim(d1.ps_tm4s), '':'', ''''))) '+@Cr+
                '         else '''' '+@Cr+
                '       end as ps_time4_s, '+@Cr+
                -- 2017/02/09 Rickliu �W�Z�ɶ�4(�X�Ըɥ��椧�U�Z���d����)
                '       case '+@Cr+
                '         when (Replace(isnull(d1.ps_tm4e, ''''), '':'', '''') <> '''') then Convert(Varchar(5), Convert(Time, Replace(Rtrim(d1.ps_tm4e), '':'', ''''))) '+@Cr+
                '         else '''' '+@Cr+
                '       end as ps_time4_e, '+@Cr+
                '       d.e_dept '+@Cr+
                '       into '+@tb_Apsnt_tmp+@Cr+
                '  from CTE_Q2 m '+@Cr+
                '       left join CTE_Q3 d on m.e_no collate Chinese_Taiwan_Stroke_CI_AS =d.e_no collate Chinese_Taiwan_Stroke_CI_AS '+@Cr+
                '       left join '+@rDB+'Apsnt d1 on m.e_no collate Chinese_Taiwan_Stroke_CI_AS =d1.ps_no collate Chinese_Taiwan_Stroke_CI_AS and m.ps_date = d1.ps_date '+@Cr+
                '       left join CTE_Q4 d2 on m.ps_date = d2.ps_date '

  Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  if @Cnt = @Errcode Goto End_Proc


  Set @Msg = '7.1.�X�Ը�ƳB�z�Ȧs��['+@tb_Apsnt_tmp+'] ���X���Ƭ�:'
  set @strSQL = 'select Count(1) as Cnt From '+@tb_Apsnt_tmp
  Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

  set @Msg = '8.�M�� ['+Convert(varchar(12), @ps_date, 111)+'] ���V�t�ΥX�Ը��.'
  set @Cnt = 0
  set @strSQL = -- 'Delete '+@rDB+@tb_Name+@Cr+
                'Delete '+@wDB+@tb_Name+@Cr+
                '  from '+@tb_Apsnt_tmp+' m ' +@Cr+
                ' where 1=1 '+@Cr+
                '   and ps_date = m.A_ps_date '+@Cr+
                '   and ps_no collate Chinese_Taiwan_Stroke_CI_AS = m.e_no collate Chinese_Taiwan_Stroke_CI_AS '+@Cr+
                '   and (m.ps_time_s collate Chinese_Taiwan_Stroke_CI_AS <> ps_tm1s collate Chinese_Taiwan_Stroke_CI_AS '+@Cr+
                '    or m.ps_time_e collate Chinese_Taiwan_Stroke_CI_AS <> ps_tm1e collate Chinese_Taiwan_Stroke_CI_AS '+@Cr+
                '    or m.ps_time2_s collate Chinese_Taiwan_Stroke_CI_AS <> ps_tm2s collate Chinese_Taiwan_Stroke_CI_AS) '
  
  Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  if @Cnt = @Errcode Goto End_Proc
  

  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
  set @Msg = '9.�פJ �X�Ը�Ʀܭ�V�t��['+@tb_Name+'].'
  set @Cnt = 0
  set @strSQL = --'insert into '+@rDB+@tb_Name+@Cr+
                'insert into '+@wDB+@tb_Name+@Cr+
                '(ps_date, ps_wk, ps_no, ps_tm1s, ps_tm1e, ps_tm2s, ps_tm2e, ps_tm3s, ps_tm3e, ps_tm4s, ps_tm4e, ps_dept)'+@Cr+
-- 2014/06/25 Rickliu �޲z�����ܥX�Ԭ����Ȩ��Ĥ@���P�̫�@����@�W�U�Z�����C
                'select A_ps_date, ps_week, e_no, ps_time_s, ps_time_e, ps_time2_s, ps_time2_e, ps_time3_s, ps_time3_e, ps_time4_s, ps_time4_e, e_dept '+@Cr+
                '  from '+@tb_Apsnt_tmp+' m ' +@Cr+
                ' where not exists'+@Cr+
                '       (select 1'+@Cr+
                '          from '+@wDB+@tb_Name+' d'+@Cr+
                '         where m.e_no collate Chinese_Taiwan_Stroke_CI_AS = d.ps_no collate Chinese_Taiwan_Stroke_CI_AS'+@Cr+
                '           and m.A_ps_date = d.ps_date '+@Cr+
                '       )'

  Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  --2014/01/02 brian�W�[�U�Z�ɶ�

  set @strSQL='UPDATE '+@wDB+@tb_Name++@Cr+
              '   SET ps_tm1s = ps_tm4s	'+@Cr+	
              ' where Rtrim(isnull(ps_tm4s, '''')) <> '''' '+@Cr+
              '   and Rtrim(isnull(ps_tm1s, '''')) <> '''' '+@Cr+
              ' '+@Cr+
              'UPDATE '+@wDB+@tb_Name++@Cr+
              '   SET ps_tm1e = ps_tm4e	'+@Cr+	
              ' where Rtrim(isnull(ps_tm4e, '''')) <> '''' '+@Cr+
              '   and Rtrim(isnull(ps_tm1e, '''')) <> '''' '

  Exec @Cnt = uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL

  if @Cnt = @Errcode Goto End_Proc
  --*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*--*
End_Proc:
/*
  -- �o�e MAIL �������H��
--20140106 brian�W�[�P�_�P����L��d��Ƥ��H�e���~
--   �P���@  1  �P���� 7

if @weekday <> 7    
	begin
	*/
	  if @Cnt = @Errcode 
	  begin
		 set @Run_Msg = '['+Convert(varchar(12), @ps_date, 111)+'] '+ @Proce_SubName + '...���楢��!!!'
		 set @Body_Msg = '���~�T���G'+@Cr+@Msg + '...���楢��!!!'
		 
		 Declare @sCmd Varchar(1000)
		 Declare @sMsg Varchar(1000)
         Declare @weekday int
         Exec uSP_Sys_Write_Log @Proc, @Run_Msg, @Body_Msg, -1, 2
         
         select @Cnt = count(1) 
           from trans_log 
          where process=@Proc
            and recordcount = -1
            --and (sqlcmd like '%sp_send_dbmail%' )
            and Trans_date >= Convert(Varchar(20), GETDATE(), 111)
            and Trans_date <= Convert(Varchar(20), GETDATE()+1, 111)
         
         set @weekday=(SELECT DATEPART(WEEKDAY, GETDATE()-1))
         /**************************************************************************************************
         �`�N�G�|���o�@�q�O�D�n�]���@���P������S���W�Z�ɤ��|�C�����oMAIL�X�h�A
               �b���ˮ� Trans_Log �O�_���e���~�T�� 10 ���A�Y���u�n�o�@��MAIL�X�h�Y�i�C-- Rickliu 2014/04/07
         **************************************************************************************************/
         if (@Cnt = 10) and (@weekday in (6, 7))
            Exec uSP_Sys_Write_Log @Proc, @Run_Msg, @Body_Msg, -1

		 
	  end
End_Exit:
  --Exec uSP_Advanced_Options 1
  Return(@Cnt)
end
GO
