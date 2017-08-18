USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_RPT_12003]    Script Date: 08/18/2017 17:43:41 ******/
DROP PROCEDURE [dbo].[uSP_RPT_12003]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE   PROCEDURE [dbo].[uSP_RPT_12003]
AS
BEGIN
/*
20140109 brian 
1.���]�t�w�N�H�� 2��
2.���]�tCZZ002�^�бj�b�� 1��


==============================================================================================
-- �ۭq�@�� DATETIME ���O���ܼ�  
DECLARE @myDate AS DATETIME  
  
-- �H 2008 �~ 3 �� 19 �鬰����I�A��X 2008 �~�Ĥ@�Ѫ����  
-- �`�N�G��H�u�~�v�����  
SET @myDate = DATEADD(yy, DATEDIFF(yy, '', '2008/3/19'), '')  
  
-- �H�u��v�����A���� 1 ��  
SELECT DATEADD(day, -1, @myDate) [2008�~3��19�骺�e�@�~�̫�@��(��T��@��)]  
================================================================================================
*/

  --�ثe�U�����H��
  select M.deptname,count(M.deptname) as 'dept_people'  
         into #tmp_people
    from (select Org_Order,e_rdate,
                 case 
                     When Org_Order like '110%'  then '���ƪ���' 
                     When Org_Order like '120%'  then '�`�g�z��'
                     When Org_Order like '130%'  then '������`'
                     When Org_Order like 'A00%'  then '�`�޲z�B'
                     When Org_Order like 'A11%'  then '�޲z��'
                     When Org_Order like 'A21%'  then '�]�|��'
                     When Org_Order like 'A31%'  then '��T��'
                     When Org_Order like 'A41%'  then '�����'
                     When Org_Order like 'A51%'  then '���ʳ�'
                     When Org_Order like 'A61%'  then '�����'
                     When Org_Order like 'B11%' and e_no like 'T%' then '�T����~��'
                     When Org_Order like 'B12%'  then '�a�x������'
                     When Org_Order like 'B13%'  then '�q�l�ӰȽ�'
                     When Org_Order like 'B16%'  then '��ڶT����'
                     When Org_Order like 'B19%'  then '�ȪA��'
                     When Org_Order like 'D11%'  then '�Ͳ���'
                     When Org_Order like 'D21%'  then '�~�㳡'
                     When Org_Order like 'E10%'  then '�s�{-�w�N'
                     else  '��L'
                 end as deptname
            from uV_TA13#EMP_LEVEL
           Where leave='N' and e_no like 'T%'  
         )  M 
   group by M.deptname
    
     
    
   --�W�멳�H��
   select deptname,count(1) as pre_m_people 
          into #tmp_lastm_people
     from (select Org_Order,e_rdate,
                  case 
                    When Org_Order like '110%'  then '���ƪ���' 
                    When Org_Order like '120%'  then '�`�g�z��'
                    When Org_Order like '130%'  then '������`'
                    When Org_Order like 'A00%'  then '�`�޲z�B'
                    When Org_Order like 'A11%'  then '�޲z��'
                    When Org_Order like 'A21%'  then '�]�|��'
                    When Org_Order like 'A31%'  then '��T��'
                    When Org_Order like 'A41%'  then '�����'
                    When Org_Order like 'A51%'  then '���ʳ�'
                    When Org_Order like 'A61%'  then '�����'
                    When Org_Order like 'B11%' and e_no like 'T%' then '�T����~��'
                    When Org_Order like 'B12%'  then '�a�x������'
                    When Org_Order like 'B13%'  then '�q�l�ӰȽ�'
                    When Org_Order like 'B16%'  then '��ڶT����'
                    When Org_Order like 'B19%'  then '�ȪA��'
                    When Org_Order like 'D11%'  then '�Ͳ���'
                    When Org_Order like 'D21%'  then '�~�㳡'
                    When Org_Order like 'E10%'  then '�s�{-�w�N'
                    else  '��L'
                  end as deptname
             from uV_TA13#EMP_LEVEL
             --�W�멳�H��  ��¾=Y �� ��¾��>=������  +  ����¾ 
            Where (leave='Y' and e_ldate >= dateadd(mm,datediff(mm,'',getdate()),'') ) 
               or leave='N' and e_no like 'T%'
          ) Y
    group by deptname 
    
    --�s�i�H��
    select deptname,count(1) as new_people 
           into #tmp_new_people
      from (select *,
                   case 
                     When Org_Order like '110%'  then '���ƪ���' 
                     When Org_Order like '120%'  then '�`�g�z��'
                     When Org_Order like '130%'  then '������`'
                     When Org_Order like 'A00%'  then '�`�޲z�B'
                     When Org_Order like 'A11%'  then '�޲z��'
                     When Org_Order like 'A21%'  then '�]�|��'
                     When Org_Order like 'A31%'  then '��T��'
                     When Org_Order like 'A41%'  then '�����'
                     When Org_Order like 'A51%'  then '���ʳ�'
                     When Org_Order like 'A61%'  then '�����'
                     When Org_Order like 'B11%' and e_no like 'T%' then '�T����~��'
                     When Org_Order like 'B12%'  then '�a�x������'
                     When Org_Order like 'B13%'  then '�q�l�ӰȽ�'
                     When Org_Order like 'B16%'  then '��ڶT����'
                     When Org_Order like 'B19%'  then '�ȪA��'
                     When Org_Order like 'D11%'  then '�Ͳ���'
                     When Org_Order like 'D21%'  then '�~�㳡'
                     When Org_Order like 'E10%'  then '�s�{-�w�N'
                     else  '��L'
                   end as deptname
              from uV_TA13#EMP_LEVEL
              --��¾���>=������
             Where e_rdate >= dateadd(mm,datediff(mm,'',getdate()),'') 
               and e_no like 'T%'
           ) Y
     group by deptname 
    
    --������¾
    select deptname,count(1) as leave_people
           into #tmp_leave_people
      from (select *,
                   case 
                     When Org_Order like '110%'  then '���ƪ���' 
                     When Org_Order like '120%'  then '�`�g�z��'
                     When Org_Order like '130%'  then '������`'
                     When Org_Order like 'A00%'  then '�`�޲z�B'
                     When Org_Order like 'A11%'  then '�޲z��'
                     When Org_Order like 'A21%'  then '�]�|��'
                     When Org_Order like 'A31%'  then '��T��'
                     When Org_Order like 'A41%'  then '�����'
                     When Org_Order like 'A51%'  then '���ʳ�'
                     When Org_Order like 'A61%'  then '�����'
                     When Org_Order like 'B11%' and e_no like 'T%' then '�T����~��'
                     When Org_Order like 'B12%'  then '�a�x������'
                     When Org_Order like 'B13%'  then '�q�l�ӰȽ�'
                     When Org_Order like 'B16%'  then '��ڶT����'
                     When Org_Order like 'B19%'  then '�ȪA��'
                     When Org_Order like 'D11%'  then '�Ͳ���'
                     When Org_Order like 'D21%'  then '�~�㳡'
                     When Org_Order like 'E10%'  then '�s�{-�w�N'
                     else  '��L'
                   end as deptname
              from uV_TA13#EMP_LEVEL
              --��¾=Y and ��¾��� >= ������
             Where (leave='Y') and (e_ldate >= dateadd(mm,datediff(mm,'',getdate()),'') ) 
               and e_no like 'T%'
           ) Y
     group by deptname 
    
    
    --������J
    --WHERE ����A�̾ڳ]�w�Ƚվ�
    select deptname,count(1) as deptin
           into #tmp_deptin
      from (select Org_Order,m.e_rdate,
                   case 
                     When Org_Order like '110%'  then '���ƪ���' 
                     When Org_Order like '120%'  then '�`�g�z��'
                     When Org_Order like '130%'  then '������`'
                     When Org_Order like 'A00%'  then '�`�޲z�B'
                     When Org_Order like 'A11%'  then '�޲z��'
                     When Org_Order like 'A21%'  then '�]�|��'
                     When Org_Order like 'A31%'  then '��T��'
                     When Org_Order like 'A41%'  then '�����'
                     When Org_Order like 'A51%'  then '���ʳ�'
                     When Org_Order like 'A61%'  then '�����'
                     When Org_Order like 'B11%' and m.e_no like 'T%' then '�T����~��'
                     When Org_Order like 'B12%'  then '�a�x������'
                     When Org_Order like 'B13%'  then '�q�l�ӰȽ�'
                     When Org_Order like 'B16%'  then '��ڶT����'
                     When Org_Order like 'B19%'  then '�ȪA��'
                     When Org_Order like 'D11%'  then '�Ͳ���'
                     When Org_Order like 'D21%'  then '�~�㳡'
                     When Org_Order like 'E10%'  then '�s�{-�w�N'
                     else  '��L'
                   end as deptname
              from uV_TA13#EMP_LEVEL m
                   join SYNC_TA13.dbo.pemploy d on m.e_no=d.e_no 
                   --��¾=Y and ��¾��� >= ������
             Where m.e_rdate = dateadd(mm,datediff(mm,'',getdate()),'') 
               and m.e_no like 'T%'
           ) Y
     group by deptname
    
    
    
    --������X
    --WHERE ����A�̾ڳ]�w�Ƚվ�
    select deptname,count(1) as deptout
           into #tmp_deptout
      from (select Org_Order,m.e_rdate,
                   case 
                     When Org_Order like '110%'  then '���ƪ���' 
                     When Org_Order like '120%'  then '�`�g�z��'
                     When Org_Order like '130%'  then '������`'
                     When Org_Order like 'A00%'  then '�`�޲z�B'
                     When Org_Order like 'A11%'  then '�޲z��'
                     When Org_Order like 'A21%'  then '�]�|��'
                     When Org_Order like 'A31%'  then '��T��'
                     When Org_Order like 'A41%'  then '�����'
                     When Org_Order like 'A51%'  then '���ʳ�'
                     When Org_Order like 'A61%'  then '�����'
                     When Org_Order like 'B11%' and m.e_no like 'T%' then '�T����~��'
                     When Org_Order like 'B12%'  then '�a�x������'
                     When Org_Order like 'B13%'  then '�q�l�ӰȽ�'
                     When Org_Order like 'B16%'  then '��ڶT����'
                     When Org_Order like 'B19%'  then '�ȪA��'
                     When Org_Order like 'D11%'  then '�Ͳ���'
                     When Org_Order like 'D21%'  then '�~�㳡'
                     When Org_Order like 'E10%'  then '�s�{-�w�N'
                     else  '��L'
                   end as deptname
              from uV_TA13#EMP_LEVEL m
                   join SYNC_TA13.dbo.pemploy d on m.e_no=d.e_no 
                   --��¾��� >= ������
             Where m.e_rdate = dateadd(mm,datediff(mm,'',getdate()),'') 
               and m.e_no like 'T%'
           ) Y
     group by deptname
    
    --SQ Query
    
    select a.deptname,a.dept_people,b.pre_m_people,c.new_people,d.leave_people,e.deptin,f.deptout
        from #tmp_people a
             left join #tmp_lastm_people b 
               on a.deptname=b.deptname
             left join #tmp_new_people c 
               on a.deptname=c.deptname
             left join #tmp_leave_people d 
               on a.deptname=d.deptname
             left join #tmp_deptin e 
               on a.deptname=e.deptname
             left join #tmp_deptout f 
               on a.deptname=e.deptname
  
End
GO
