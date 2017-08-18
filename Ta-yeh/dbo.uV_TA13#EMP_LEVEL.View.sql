USE [DW]
GO
/****** Object:  View [dbo].[uV_TA13#EMP_LEVEL]    Script Date: 08/18/2017 17:43:39 ******/
DROP VIEW [dbo].[uV_TA13#EMP_LEVEL]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create View [dbo].[uV_TA13#EMP_LEVEL]
as
  -- 2017/03/17 Rickliu �쥻�O�H TYEIPDBS2.lytdbta13.SP_EMP_LEVEL Store_Proce ���X TB_EMP_LEVEL ��ƪ�A�]��έq�\��Ʈw�覡�A�ҥH�אּ VIEW ���o��ơC
  -- �����Ҭݯ�_�� CTE �覡�A�[�t��ƳB�z�C
  select *,
         case 
           when dp_no <> Chg_Dept_Level then  ''
           else rtrim(e_no)
         end as Result_e_no, -- ����촣�ѵ� SmartQuery ���e�ѼƩҥ�
         case 
           when Chg_Duty_Level >=3 then  ''
           else rtrim(Convert(Varchar(5), Chg_dept_level))
         end as Result_dept_level, -- ����촣�ѵ� SmartQuery ���e�ѼƩҥ�
         case 
           when Chg_Duty_Level >=3 then  ''
           else rtrim(Convert(Varchar(5), Chg_duty_level))
         end as Result_duty_level, -- ����촣�ѵ� SmartQuery ���e�ѼƩҥ�
         case 
           when Chg_Duty_Level >=3 then  ''
           else rtrim(e_mstno)
         end as Result_e_mstno, -- ����촣�ѵ� SmartQuery ���e�ѼƩҥ�
         getdate() as cr_date
    from (Select distinct
                 Rtrim(M.e_no) as e_no, 
                 Rtrim(M.e_name) as e_name, 
                 replace(replace(dp_no, 'Z', '1'), 'B2', 'B1') as Org_Order, --
                 Case
                   When dp_no in ('Z1000', 'Z3000', '11000', '13000') then 1
                   When substring(dp_no, 1, 2) in ('Z2', '12') then 2
                   When substring(dp_no, 2, 1)  = '0'  then 3
                   When substring(dp_no, 1, 3) in ('Z11', 'Z12', 'Z13', '111', '112', '113') or substring(dp_no, 3, 1) ='0' then 4
                   
                   When substring(dp_no, 4, 1) ='0' then 5
                   When substring(dp_no, 5, 1) ='0' then 6
                   else 7
                 end as dp_lv, -- ��´���h�s��
                 Rtrim(D1.dp_no) as dp_no, 
                 Substring(Rtrim(D1.dp_no), 1, Len(Rtrim(D1.dp_no))-1) as upper_dp, -- �W�h��´
                 Rtrim(D1.dp_name) as dp_name, 
                 Case 
                   When e_no like '%ZZ%' then '�����b��'
                   else Rtrim(M.e_duty)
                 end as e_duty,
                 M.e_mstno, -- �էO�s��
                 rtrim(Replace(Convert(Varchar(100),dp_rem),'�³����s��:','')) as 'old_dpno',
                 M.e_rdate, -- ��¾���
                 M.e_ldate, -- ��¾���
                 case 
                   -- 2015/11/13 Rickliu ������Фŧ�ʡA�]�X�Է|�������Ƨ@����¾�P�_�̾ڡC
                   when (convert(varchar(10), e_ldate, 111)='1900/01/01') and (D2.status ='1')
                   then 'Y'
                   else 'N'
                 end as status, -- EIP �b���ҥ�
                 rtrim(substring(dp_no, 1, 2)) as Chg_Master_Dept,
/**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**
 2015/11/11 Rickliu ¾�ٵ��šAL3(�t)�H�W���ťi�H�ݨ�Ҧ����
    7: ���ƪ���
    6: �`�g�z�B���`��
    5: �޲z����z
    4: �g�z�šB�ȪA�ժ��B�~�ȧU�z�B���ʲժ��B�~�@�Ҫ�
    3: �`�޲z�B�B�޲z���B�]�|��
       �Ҫ��šB�էO�]�w LS ��
       �u�����s�����դ���ݡA��L��P�����Ҧ����ҥi�H��
    -----------------------------------------------------------^ �i�H�ݱo��Ҧ����
    2: �ժ��B�D���A�D�~�@�ҷ~�ȽҪ�
    1: �M���ΧU�z
    0: ��¾���u
 --**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**/
                 Case --���u¾��(�@���i�d�߸�Ʃһ�¾���A�ФűN�H�U���ǽհʡA�P�_�����ǬO�ѤW���U�P�_
                   -- L0
                   When e_ldate <> '1900/01/01' then 0 --��¾���u

                   -- L7
                   When e_Duty like '%���ƪ�%' or substring(dp_no, 1, 2) in ('A3') then 7

                   -- L6
                   When e_Duty like '%�`�g�z%' or e_Duty like '%���`%' then 6

                   -- L5
                   When e_Duty like '%�޲z%��z%' then 5

                   -- L4
                   When e_Duty like '%�g�z%' Or e_Duty like '%�ȪA�ժ�%' or e_Duty like '%�~%�U%' or e_Duty like '%���ʲժ�%' then 4 
                   When substring(dp_no, 1, 2) = 'B2' and e_Duty = '%�~�ȽҪ�%' then 4

                   -- L2
                   When e_Duty like '%�ժ�%' or e_Duty like '%�D��%' then 2
                   When substring(dp_no, 1, 2) <> 'B2' and e_Duty = '%�~�ȽҪ�%' then 2

                   -- L3
                   When (e_Duty like '%�Ҫ�%' or e_mstno = 'LS') then 3
                   When (e_Duty like '%ĵ��%' or e_mstno = 'LS') then 1 
                   When (substring(dp_no, 1, 2) = 'B2' and e_Duty = '%�~�ȽҪ�%') then 3
                   When (substring(dp_no, 1, 2) in ('A0', 'A1', 'A2')) then 3 --�`�޲z�B�B�޲z���B�]�|��
                        --2015/11/11 Rickliu ���`���� �u�����s�����դ���ݡA��L��P�����Ҧ����ҥi�H��
                   When (substring(dp_no, 1, 2) = 'A4' And dp_no <> 'A4110') then 3
				        --�Ӷ}
                   When (substring(dp_no, 1, 2) = 'A5') then 3

                   -- L1
                   else 1 --�M���ΧU�z
                 end as Chg_Duty_Level, 

/**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**
 2015/11/11 Rickliu ��������
   ����Ƶ���(�M�ų����s��   )�G���ƪ��B�`�g�z�B���`�šB�`�޲z�B�B�޲z���B�]�|��
   �Ʒ~�s����(�����s���e 1 �X)�G�ȭq�޲z¾����z¾�٪�
   ���ŵ���  (�����s���e 2 �X)�G�g�z�B�ȪA�ժ��B�~�U�B�~�@�Ҫ�
   �үŵ���  (�����s���e 3 �X)�G�Ҫ��šB�էO�]�w LS ��
   �էO����  (�����s���e 4 �X)�G�Z��¾�H���H�ΤW�z���C�����
 --**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**/
                 Case --��������(�@����ƹ����������)�A�ФűN�H�U���ǽհʡA�P�_�����ǬO�ѤW���U�P�_
                   -- �̤p������
                   When e_ldate <> '1900/01/01' then rtrim(e_dept) --��¾���u

                   -- ����Ƶ���
                   When e_Duty like '%���ƪ�%' then '' 
                   When e_Duty like '%�`�g�z%' or e_Duty like '%���`%' then ''
                   When substring(rtrim(e_dept), 1, 2) in ('A0', 'A1', 'A2', 'A3', 'A5') then ''

                   -- �Ʒ~�s����
                   When e_Duty like '%�޲z%��z%' then substring(rtrim(e_dept), 1, 1)  --��z(�i�d�ߦU�����H�U���)

                   -- ���ŵ���
                   When e_Duty like '%�g�z%' then substring(rtrim(e_dept), 1, 2) 
                   When e_Duty like '%�ȪA�ժ�%' or e_Duty like '%�~%�U%' then substring(rtrim(e_dept), 1, 2)
                   When (substring(dp_no, 1, 2) = 'B2' and e_Duty = '%�~�ȽҪ�%') then substring(rtrim(e_dept), 1, 2) 

                   -- �үŵ���
                   When (e_Duty like '%�Ҫ�%' or e_mstno = 'LS') then substring(e_dept, 1, 3) --�Ҫ�(�i�d�ߦۤv�����H�U�����)

                   -- �էO����
                   When (substring(dp_no, 1, 2) <> 'B2' and e_Duty = '%�~�ȽҪ�%') then substring(rtrim(e_dept), 1, 4)
                   When e_Duty like '%�ժ�%' or e_Duty like '%�D��%' then substring(rtrim(e_dept), 1, 4) --�ժ� or �D��(�ȥi�d�߳��ݸ��)

                   -- �̤p������
                   else rtrim(e_dept)
                 end as Chg_Dept_Level, 

/**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**
 2015/11/11 Rickliu ��������
                  |       |
    ���� < �j�L < | �� �L | < �p�L < �@�� < ����
    (     2      )|(  1  )|(         0         ) ==> ��������
    L2�G���ƪ��B�`�g�z�B���`�šB�]�|���B��T���B�����(���t���s����)�B���ʳ�
    L1�G�`�޲z�B�B�޲z���B�g�z�šB�Ҫ��šB�ժ��šB�~�U�šB�D���šB�էO�]�w LS ��
    L0�G�Z��¾�H���H�ΤW�z���C�����
 --**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**--**/
                 
                 Case --�ܮw��������v���A�ФűN�H�U���ǽհʡA�P�_�����ǬO�ѤW���U�P�_
                   -- L0
                   When e_ldate <> '1900/01/01' then 0  --��¾���u

                   -- L2
                   When e_Duty like '%���ƪ�%' then 2
                   When e_Duty like '%�`�g�z%' or e_Duty like '%���`%' then 2
                   When dp_no like 'A[2345]000%' then 2
                   When dp_no like 'A[235]%' then 2
                   -- 2015/11/11 Rickliu ���`���� �P���F���s�����դ���ݨ�L��P�H�����\�d��
                   When (dp_no like 'A4%') And (dp_no not like 'A411%') then 2 
				   -- �Ӷ}
                   When (dp_no like 'A5%')  then 2 
                   
                   -- L1
                   When e_Duty like '%�޲z%��z%' then 1 
                   When e_Duty like '%�g�z%' or e_Duty like '%�Ҫ�%' or e_Duty like '%�ժ�%' then 1
                   When e_Duty like '%�~%�U%' or e_Duty like '%�D��%' or e_mstno = 'LS' then 1

                   -- L0
                   else 0 --�L��������v��
                 end as stock_amt_level, 
          
                 Rtrim(D3.PassWord) as EIP_Password, -- EIP �K�X
                 Rtrim(D2.EMail) as EMail, -- �j�~�ӤH MAIL
                 Rtrim(D2.MSNAccount) as MSNAccount, 
                 Rtrim(D2.SKYPEAccount) as SKYPEAccount, 
                 case
                   when (convert(varchar(10), e_ldate, 111) >= '2013/10/01') Or
                        (convert(varchar(10), e_ldate, 111) = '1900/01/01')
                   then Substring(D2.AccountID, 1, 2) 
                   else ''
                 end as sDept, --
                 Rtrim(D2.AccountID) as AccountID,
                 Case 
                   When rtrim(D2.OfficeExt) = '' then '000'
                   else rtrim(D2.OfficeExt)
                 end as OfficeExt, -- ������
                 Case 
                   When rtrim(D2.MsnAccount) = '' then '000'
                   else rtrim(D2.MsnAccount)
                 end as PC_Name, -- �w�]�q���W��
                 case 
                   when convert(varchar(10), e_ldate, 111)='1900/01/01' 
                   then 'N' -- �b¾
                   else 'Y' -- ��¾
                 end as 'leave', -- ��¾�_
                 Rtrim(D2.Remark) as Remark
            From SYNC_TA13.dbo.pemploy as M 
                 Inner join SYNC_TA13.dbo.pdept as D1 
                    on M.e_dept = D1.dp_no 
                  -- 2014/03/20 �ФŨϥ� Inner join  �]�� EIP �i��L���b���A�� SQ �����������u��Ƥ~��վ\�X���ݭ�����C
                  left join WebEIP2.dbo.Account_Data_View as D2 
                    on m.e_no = D2.pager Collate Chinese_Taiwan_Stroke_CI_AS
                   and M.e_name Collate Chinese_Taiwan_Stroke_CI_AS = D2.FullName
                  Left outer join WebEIP2.dbo.AFS_AccountView as D3
                    on D2.AccountID = D3.AccountID
           Where (M.e_ldate >= '1900/1/1') 
             And (M.e_no <> 'LY') 
             And (substring(M.e_no, 1, 3) <> 'TAZ')
         ) m
GO
