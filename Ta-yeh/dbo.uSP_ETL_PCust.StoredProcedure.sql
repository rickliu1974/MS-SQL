USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_ETL_PCust]    Script Date: 08/18/2017 17:18:56 ******/
DROP PROCEDURE [dbo].[uSP_ETL_PCust]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_ETL_PCust]
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_TA13#ETL_PCust
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [�ק� Trans_Log �T�����e]
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  

  Declare @Proc Varchar(50) = 'uSP_ETL_PCust'
  Declare @Cnt Int =0
  Declare @Cnt_Ori Int =0
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @subject Varchar(500)= 'Exec '+@Proc+' Error!!'
  Declare @Result int = 0
  
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Fact_PCust]') AND type in (N'U'))
  begin
     Set @Msg = '�R����ƪ� [Fact_PCust]'
     set @strSQL= 'DROP TABLE [dbo].[Fact_PCust]'

     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end

  begin try
    set @Msg = '�إ߸�ƪ� [Fact_PCust]'
    Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, 0

    -- 2017/04/21 Rickliu ���O�̫�@�����
    ;With CTE_Q1 as (
       select r_name, max(r_date) as r_date
         from SYNC_TA13.dbo.prate
        group by r_name
    -- 2017/04/21 Rickliu �̷s���O
    ), CTE_Q2 as (
       select a.r_name, a.r_rate, a.r_date as r_date
         from SYNC_TA13.dbo.prate as a
         left join CTE_Q1 as b
           on a.r_name=b.r_name and a.r_date=b.r_date
        where a.r_name=b.r_name and a.r_date=b.r_date
    )

    select -- Distinct
           M.[ct_class], --���O
           Rtrim(M.[ct_no]) as ct_no, --�s��
           Rtrim(M.[ct_name]) as ct_name, --�W��
           Rtrim(M.[ct_sname]) as ct_sname, --²��
           -- 2017/08/07 Rickliu �u�n�O IT�~�ȭӤH�BIZ�{���Ȥ�BZZ��L�Ȥ� �@�߳��N�s�� ���X�� 000001
           Case
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 8 and Substring(M.ct_no, 1, 2) in ('IT', 'IZ', 'ZZ')
             then Substring(M.ct_no, 1, 2)+'000001'
             else Substring(Rtrim(M.[ct_no]), 1, 8) 
           end as ct_no8, --�Ȥ�K�X�s��(��@�����s���A���X�h���`���q�s��)
           -- 20170428 Rickliu ���� ct_sname �覡�A���|�����~���p�A�N��ħ�� ct_fld3+ct_fld4 �覡�e�{
           --LTrim(RTrim(Replace(Replace(Replace(Replace(Replace(Replace(m.ct_sname, '#', ''), '@', ''), '-', ''), 'T', ''), 'P', ''), '/', ''))) as ct_sname8, -- �Ȥ�K�X�W��
           -- 2017/08/07 Rickliu �u�n�O IT�~�ȭӤH�BIZ�{���Ȥ�BZZ��L�Ȥ� �@�߳��k��@�W��
           Case
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 8 and Substring(M.ct_no, 1, 2) = 'IT' then '�~�ȭӤH'
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 8 and Substring(M.ct_no, 1, 2) = 'IZ' then '�{���Ȥ�'
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 8 and Substring(M.ct_no, 1, 2) = 'ZZ' then '��L�Ȥ�'
             else rtrim(m.ct_fld3)+case when rtrim(m.ct_fld4) <> '' then '-' else '' end+rtrim(m.ct_fld4) 
           end as ct_sname8,

           [ct_addr1]=Rtrim(Convert(Varchar(255), M.ct_addr1)), --���q�a�}
           [ct_addr2]=Rtrim(Convert(Varchar(255), M.ct_addr2)), --�o���a�}
           [ct_addr3]=Rtrim(Convert(Varchar(255), M.ct_addr3)), --�e�f�a�}
           Rtrim(M.[ct_tel]) as ct_tel, --�q��
           Rtrim(M.[ct_fax]) as ct_fax, --�ǯu
           Rtrim(M.[ct_unino]) as ct_unino, --�Τ@�s��
           Rtrim(M.[ct_presid]) as ct_presid, --�t�d�H
           Rtrim(M.[ct_contact]) as ct_contact, --�p���H
           M.[ct_payfg], --�������
           Rtrim(M.[ct_sales]) as ct_sales, --�~�ȭ�
           M.[ct_p_limit], --�b���B��
           M.[ct_b_limit], --�H���B��
           Rtrim(M.[ct_bkno]) as ct_bkno, --�Ȧ�b��
           Rtrim(M.[ct_bknm]) as ct_bknm, --�Ȧ�W��
           [ct_rem]=Rtrim(Convert(Varchar(255), M.ct_rem)), --�Ƶ�
           Rtrim(M.[ct_dept]) as ct_dept, --�����s��
           M.[ct_payrate], --������
           Rtrim(M.[ct_ivtitle]) as ct_ivtitle, --�o�����Y
           Rtrim(M.[ct_porter]) as ct_porter, --�f�B���q
           Rtrim(M.[ct_credit]) as ct_credit, --�д�|�I�ڹ�H
           M.[ct_pmode], --�I�ڤ覡
           M.[ct_pdate], --�I�ڤѼ�
           M.[ct_prenpay], --���M��
           M.[ct_prepay], --�w����
           M.[ct_last_dt], --�̪�����
           M.[ct_flg], --�Ȥ�ݼt��
           M.[ct_t_fg], --�ǿ�X��
           M.[ct_grade], --����
           M.[ct_area], --�a��,��O�N�X
           -- 2015/02/03 ������O
           case
             when RTrim(isnull(M.ct_curt_id, '')) = '' then 'NT'
             else RTrim(isnull(M.ct_curt_id, ''))
           end as ct_curt_id,

           Rtrim(M.[ct_cont_sp]) as ct_cont_sp, --�p���H¾��
           M.[ct_pay], --�T���I�ڤ覡
           Rtrim(M.[ct_regist]) as ct_regist, --�u�t�n�O��
           M.[ct_worker], --���u��
           M.[ct_capital], --�ꥻ�B
           M.[ct_skpay], --�T���w����
           M.[ct_sknpay], --�T�����M��
           M.[ct_accno2], --�T�����I�b���
           M.[ct_chkno2], --�T�����I�����
           M.[ct_cdate], -- ���ɤ��
           M.[ct_payer], --�i�P�s�дڹ�H
           M.[ct_advance], --�i�P�s�w����
           M.[ct_debt], --�i�P�s���M��
           M.[ct_abroad], --�ꤺ.��~
           Rtrim(M.[ct_fld1]) as ct_fld1, --�Ȥ�|�t�Ӫ����@(�ȡG�}�o����B�t�G�t�X�_�l��)
           Rtrim(M.[ct_fld2]) as ct_fld2, --�Ȥ�|�t�Ӫ����G(�ȡG��������B�t�G�����t�X��)
           -- 2017/08/07 Rickliu �u�n�O IT�~�ȭӤH�BIZ�{���Ȥ�BZZ��L�Ȥ� �@�߳��k��@�W��
           --Rtrim(M.[ct_fld3]) as ct_fld3, --�Ȥ�|�t�Ӫ����T(�ȡG�`���q�O�B�t�G���ϥ�)
           Case
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 8 and Substring(M.ct_no, 1, 2) = 'IT' then '�~�ȭӤH'
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 8 and Substring(M.ct_no, 1, 2) = 'IZ' then '�{���Ȥ�'
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 8 and Substring(M.ct_no, 1, 2) = 'ZZ' then '��L�Ȥ�'
             else rtrim(m.ct_fld3)
           end as ct_fld3,
           Rtrim(M.[ct_fld4]) as ct_fld4, --�Ȥ�|�t�Ӫ����|(�ȡG�����O�B�t�G�t�ӵ���)
           Rtrim(M.[ct_fld5]) as ct_fld5, --�Ȥ�|�t�Ӫ�����(�ȡG�ȪA�M��1�B�t�Gñ�������)
           Rtrim(M.[ct_fld6]) as ct_fld6, --�Ȥ�|�t�Ӫ�����(�ȡG�ȪA�M��2�B�t�G���ϥ�)
           Rtrim(M.[ct_fld7]) as ct_fld7, --�Ȥ�|�t�Ӫ����C(�ȡG�ȪA�M��3�B�t�G���ϥ�)
           Rtrim(M.[ct_fld8]) as ct_fld8, --�Ȥ�|�t�Ӫ����K(�ȡG�ȪA�M��4�B�t�G���ϥ�)
           Rtrim(M.[ct_fld9]) as ct_fld9, --�Ȥ�|�t�Ӫ����E(�ȡG���ڤ覡�B�t�G���ϥ�)
           Rtrim(M.[ct_fld10]) as ct_fld10, --�Ȥ�|�t�Ӫ����Q(�ȡG�o���W�١B�t�G���ϥ�)
           Rtrim(M.[ct_fld11]) as ct_fld11, --�Ȥ�|�t�Ӫ����Q�@(�ȡG�I�ڱ���B�t�G���ϥ�)
           Rtrim(M.[ct_fld12]) as ct_fld12, --�Ȥ�|�t�Ӫ����Q�G(�ȡG�½s�� 2012�}�b�ϥ�)
           M.[ct_sea], --�T���B�e�覡
           M.[ct_invofg], --�}�ߵo������
           M.[ct_udec], --���������p�Ʀ��
           M.[ct_tdec], --������B���p�Ʀ��
           M.[ct_busine], --��~�O
           Rtrim(M.[ct_banpay]) as ct_banpay, --��׻Ȧ�
           M.[ct_loc], --�ϰ�s��
           M.[ct_sour], --�Ȥ�ӷ�
           M.[ct_kind], --�Ȥ����O
           M.[ct_tkday], --���X�g��
           Chg_ctclass=
             Case 
               when M.ct_class = '1' and Len(m.ct_no) = 9 then '�Ȥ�'
               when M.ct_class = '2' and Len(m.ct_no) = 5 then '�t��'
             end,
           -- �Ȥ�
           --2016/09/09 Rickliu �]��´�X�s�A�W�[�����q���A�]���N�쥻 2~5 �אּ 2~6
           Case
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 6 and Substring(M.ct_no, 1, 2) = 'IT' then 'IT0000'
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 6 and Substring(M.ct_no, 1, 2) = 'IZ' then 'IZ0000'
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 6 and Substring(M.ct_no, 1, 2) = 'ZZ' then 'ZZ0000'
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 6 then Substring(Rtrim(M.[ct_no]), 1, 6) 
             else ''
           end as Chg_BU_NO, -- �Ȥ��`���q�s��
           --2017/06/21 Rickliu �]�ʤj�Ȥ�ݹL�o�`���q�Ƚs�A�]���s�W�����
           Case
             when m.ct_class = '1' and Len(Rtrim(M.[ct_no])) > 6 and 
                  Substring(M.ct_no, 1, 2) Not in ('IT', 'IZ', 'ZZ') and
                  (m.ct_fld3 like '%�`���q%' or m.ct_fld4 like '%�`���q%')
             then 'Y'
             else 'N'
           end as Chg_Is_BU, -- �O�_���`���q
           [Chg_ctno_Port_Office] = Rtrim(V1.Code_Name),
           [Chg_ctno_CustKind_CustCity] = Rtrim(Isnull(V2.Code_Name, V6.Code_Name)),
           [Chg_ctno_CustChain] = Rtrim(V3.Code_Name),
           [Chg_ct_dept_Name] = Rtrim(D1.DP_NAME), 
           [Chg_ct_sales_Name] = Rtrim(D2.E_NAME),
           [Chg_credit_Name] =Rtrim(D3.ct_sname),
           [Chg_busine_Name] = Rtrim(P4.Tn_Contact),
           [Chg_loc_Name] = Rtrim(P5.Tn_Contact),
           [Chg_sour_Name] = Rtrim(P6.Tn_Contact),
           [Chg_Customer_kind_Name] = Rtrim(P7.Tn_Contact),
           [Chg_payfg_Name] = Rtrim(V4.code_name),
           [Chg_porter_Name] = Rtrim(P8.tr_name),
           [Chg_pmode_Name] = Rtrim(V5.code_name),
           Chg_fld1_Year = substring(M.[ct_fld1], 1, 4),
           Chg_fld1_Month = substring(M.[ct_fld1], 6, 2), 
           Chg_fld2_Year = substring(M.[ct_fld2], 1, 4),
           Chg_fld2_Month = substring(M.[ct_fld2], 6, 2),
           Chg_Hunderd_Customer = -- �~�ȫȤ�ʤj�W��
             Case
               when isnull(P9.Customer, '') <> '' then 'Y'
               else 'N'
             end,
           Chg_Hunderd_Customer_Name =
             Case
               when isnull(P9.Customer, '') <> '' then Rtrim(Customer)
               else ''
             end,
           -- 2014/1/27 Rick �W�[�P�_�O�_�����q�Ȥ�
           Case
             when ((Upper(substring(m.ct_no, 1, 2)) = 'I9' Or
                    Upper(substring(m.ct_no, 1, 5)) = 'I1826' Or
                    Upper(substring(m.ct_no, 1, 5)) = 'IZ000') and 
                    (m.ct_class = '1')
                  ) 
             then 'Y'
             else 'N'
           end as Chg_IS_Lan_Custom,
           -- 2014/10/09 Rickliu �W�[�P�O���īȤ�μt��
           Case
             when (m.ct_name like '����%' or m.ct_name like '����%' or m.ct_name like '���Ф���%' or 
                   m.ct_name like '����%' or m.ct_name like '����%' or (LTrim(RTrim(m.ct_name)) = '') or
                   m.ct_name like '���b%' or m.ct_name like '�˩�%' or
                   Rtrim(replace(m.ct_fld2, '�L', '')) <> '' or Rtrim(m.ct_name) = '' or Rtrim(replace(m.ct_fld3, '�L', '')) = '' or
                   m.ct_no like 'IT%' or m.ct_no like 'IZ%' or m.ct_no like 'ZZ%'
                  ) -- 2017/04/28 �L���P�_�O�_���Ȥ�μt��
             then 'Y'
             else ''
           end as Chg_ct_close,
           -- 2015/02/03 Rickliu �s�W�ײv
           Chg_rate_date = P10.r_date,
           Chg_rate = P10.r_rate,
           -- 2015/02/26 Rickliu �W�[ �Ȥ� ��������, ���s���ĤE�X 1: �ʳf, 2: 3C, 3: �H��, 4:�N�u(OEM)
           Case
             when m.ct_class = '1' and Len(m.ct_no) = 9 then Substring(RTrim(m.ct_no), 9, 1)
             when m.ct_class = '1' and Len(m.ct_no) = 5 then '0'
             else ''
           end as Chg_Cust_Sale_Class,
           Case
             when m.ct_class = '1' and Len(m.ct_no) = 5 then '�t��'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '1' then '�ʳf'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '2' then '3C'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '3' then '�H��'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '4' then '�N�u(OEM)'
             else ''
           end as Chg_Cust_Sale_Class_Name,
           Case
             when m.ct_class = '1' and Len(m.ct_no) = 5 then RTrim(m.ct_sname)+'-�t��'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '1' then Replace(Replace(RTrim(m.ct_fld3), '@', ''), '#', '')+'-�ʳf'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '2' then Replace(Replace(RTrim(m.ct_fld3), '@', ''), '#', '')+'-3C'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '3' then Replace(Replace(RTrim(m.ct_fld3), '@', ''), '#', '')+'-�H��'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '4' then Replace(Replace(RTrim(m.ct_fld3), '@', ''), '#', '')+'-�N�u(OEM)'
             else ''
           end as Chg_Cust_Sale_Class_sName,
           -- 2015/03/05 Rickliu �W�[ �Ȥ᦬�����������f�~�������O
           Case
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '1' then 'A'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '2' then 'B'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '3' then 'C'
             when m.ct_class = '1' and Substring(RTrim(m.ct_no), 9, 1) = '4' then 'F'
             else ''
           end as Chg_Cust_Dis_Mapping,
           -- 2017/08/04 Rickliu �s�W �t�ӱ��ʭ��ΫȤ�~�� �O�_��¾
           case when (Convert(Varchar(4), D2.e_ldate, 112) = '1900') then 'N' else 'Y' end as ct_sale_leave,
           pcust_update_datetime = getdate(),
           pcust_timestamp = m.timestamp_column
           into Fact_pcust
      from SYNC_TA13.dbo.PCust M
           left join SYNC_TA13.dbo.pdept D1 On M.ct_dept = D1.DP_NO -- �������
           left join SYNC_TA13.dbo.Pemploy D2 on M.ct_sales = D2.E_NO -- ���u���
           left join SYNC_TA13.dbo.PCust D3 On M.ct_class = D3.ct_class and M.ct_credit = D3.ct_no
           left join SYNC_TA13.dbo.pattn P4 On P4.tn_class='4' and M.ct_busine = P4.TN_NO  -- �Ȥ����O
           left join SYNC_TA13.dbo.pattn P5 On P5.tn_class='5' and M.ct_loc = P5.TN_NO -- �ϰ�O
           left join SYNC_TA13.dbo.pattn P6 On P6.tn_class='6' and M.ct_sour = P6.TN_NO -- �Ȥ�ӷ��O
           left join SYNC_TA13.dbo.pattn P7 On P7.tn_class='7' and M.ct_kind = P7.TN_NO-- �Ȥ����O
           left join SYNC_TA13.dbo.struc P8 On M.ct_porter = P8.tr_no-- �f�B���q

           --2015/03/06 Rickliu �ܧ�ĥ� Ori_Xls#Sys_Code �t�ӲĤ@�X���q�O �� �Ȥ�Ĥ@�X���~�P
           --left join V_Ctno_Port_Office V1 On M.ct_class = V1.ct_class and Substring(M.[ct_no], 1, 1)= V1.ctno_Port_Office_Kind-- 
           Left join Ori_Xls#Sys_Code V1 
                  On (V1.code_class ='6' and M.ct_class = '1' And Substring(M.[ct_no], 1, 1)= V1.Code_End Collate Chinese_Taiwan_Stroke_CI_AS)
                  Or (V1.code_class ='6' and M.ct_class = '2' And Substring(M.[ct_no], 1, 1)= V1.Code_End Collate Chinese_Taiwan_Stroke_CI_AS)


           Left join Ori_Xls#Sys_Code V2 
                  On (V2.Code_class = '1' And Len(m.ct_no) = 9 And 
                      Case
                        when Substring(M.[ct_no], 2, 4) like '[0-9][0-9][0-9][0-9]' then Convert(Int, Substring(M.[ct_no], 2, 4)) 
                        else null
                      end
                      between Convert(Int, V2.Code_Begin) and Convert(Int, V2.Code_End)
                     )
                  
                  
           --2015/03/06 Rickliu �ܧ�ĥ� Ori_Xls#Sys_Code �Ȥ�ĤE�X�������O�A�ӫ~���������j�����O�X
           --left join V_Ctno_CustChain V3 On V3.ct_class = '1' and Substring(M.[ct_no], 9, 1) COLLATE Chinese_Taiwan_Stroke_CI_AS = V3.ctno_CustChain_Kind COLLATE Chinese_Taiwan_Stroke_CI_AS-- 
           Left join Ori_Xls#Sys_Code V3 On V3.code_class ='3' and Substring(M.[ct_no], 9, 1) = V3.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS
           
           --2015/03/06 Rickliu �ܧ�ĥ� Ori_Xls#Sys_Code ����ζi������
           --left join V_Ctno_Payfg_Name V4 On M.ct_payfg = V4.ct_payfg
           left join Ori_Xls#Sys_Code V4 On V3.code_class ='4' and M.ct_payfg = V4.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS
            
           --2015/03/06 Rickliu �ܧ�ĥ� Ori_Xls#Sys_Code �ӫ~ ��/�I�ڤѼƤ覡
           --left join V_Ctno_Pmode_Name V5 On M.ct_pmode = V5.ct_pmode
           left join Ori_Xls#Sys_Code V5 On V3.code_class ='5' and M.ct_pmode = V5.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS
           left join ori_xls#Hunderd_Customer P9 on m.ct_class='1' -- 2013/2/2 �P���ɳX�ͫ�T�{�O�ϥ�²�٥h����
                 and m.ct_name collate Chinese_PRC_BIN like '%'+P9.customer+'%' collate Chinese_PRC_BIN

           Left join Ori_Xls#Sys_Code V6 
                  On (V6.Code_class = '24' And Len(m.ct_no) = 9 And 
                      Case
                        when Substring(M.[ct_no], 2, 1) like '[A-Z]' then Substring(M.[ct_no], 2, 1) Collate Chinese_Taiwan_Stroke_CI_AS
                        else null
                      end
                      between V6.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS and V6.Code_End Collate Chinese_Taiwan_Stroke_CI_AS
                     )

		   --2015/07/01 NanLiao��g�޿� ���o�ߤ@��
           --left join (select r_name, r_rate, max(r_date) as r_date
           --             from SYNC_TA13.dbo.prate
           --            group by r_name, r_rate
           left join CTE_Q2 P10 
             on Case
                  when RTrim(Isnull(M.ct_curt_id, '')) = '' then 'NT'
                  else M.ct_curt_id 
                end = P10.r_name


    /*==============================================================*/
    /* Index: pcust_Timestamp                                       */
    /*==============================================================*/
    --set @Msg = '�إ߯��� [Fact_pcust.pcust_Timestamp]'
    --Print @Msg
    --if exists (select * from sys.indexes where object_id = object_id('[dbo].[fact_pcust]') and name = 'pk_pcust')
    --   alter table [dbo].[pk_pcust] drop constraint [pk_pcust]

    --alter table [dbo].[fact_pcust] add  constraint [pk_pcust] primary key nonclustered ([pcust_timestamp] asc) with 
    --(pad_index  = off, statistics_norecompute  = off, sort_in_tempdb = off, ignore_dup_key = off, online = off, allow_row_locks  = on, allow_page_locks  = on) on [primary]

    /*==============================================================*/
    /* Index: no                                                    */
    /*==============================================================*/
    set @Msg = '�إ߯��� [Fact_pcust.no]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pcust') and name  = 'no' and indid > 0 and indid < 255) drop index dbo.fact_pcust.no
    create clustered index no on dbo.fact_pcust (ct_no, ct_class) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: class                                                 */
    /*==============================================================*/
    set @Msg = '�إ߯��� [Fact_pcust.Class]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pcust') and name  = 'class' and indid > 0 and indid < 255) drop index dbo.fact_pcust.class
    create index class on dbo.fact_pcust (ct_class) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: ct_payrate                                            */
    /*==============================================================*/
    set @Msg = '�إ߯��� [Fact_pcust.ct_payrate]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pcust') and name  = 'ct_payrate' and indid > 0 and indid < 255) drop index dbo.fact_pcust.ct_payrate
    create index ct_payrate on dbo.fact_pcust (ct_payer, ct_class) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: ctbus                                                 */
    /*==============================================================*/
    set @Msg = '�إ߯��� [Fact_pcust.ctbus]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pcust') and name  = 'ctbus' and indid > 0  and indid < 255) drop index dbo.fact_pcust.ctbus
    create index ctbus on dbo.fact_pcust (ct_busine, ct_class) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: ctkind                                                */
    /*==============================================================*/
    set @Msg = '�إ߯��� [Fact_pcust.ctkind]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pcust') and name  = 'ctkind' and indid > 0 and indid < 255) drop index dbo.fact_pcust.ctkind
    create index ctkind on dbo.fact_pcust (ct_kind, ct_class) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: ctloc                                                 */
    /*==============================================================*/
    set @Msg = '�إ߯��� [Fact_pcust.ctloc]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pcust') and name  = 'ctloc' and indid > 0 and indid < 255) drop index dbo.fact_pcust.ctloc
    create index ctloc on dbo.fact_pcust (ct_loc, ct_class) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: ctname                                                */
    /*==============================================================*/
    set @Msg = '�إ߯��� [Fact_pcust.ctname]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pcust') and name = 'ctname' and indid > 0 and indid < 255) drop index dbo.fact_pcust.ctname
    create index ctname on dbo.fact_pcust (ct_name) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /* Index: sname                                                 */
    /*==============================================================*/
    set @Msg = '�إ߯��� [Fact_pcust.sname]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_pcust') and name  = 'sname' and indid > 0 and indid < 255) drop index dbo.fact_pcust.sname
    create index sname on dbo.fact_pcust (ct_sname, ct_class) with fillfactor= 30 on "PRIMARY"

    /*==============================================================*/
    /*  ����ˮ�                                                    */
    /*==============================================================*/
    set @Msg =(select ct_no+'('+ct_fld3+'),'
                 from (select chg_bu_no, count(distinct ct_fld3) as cnt
                         from fact_pcust  
                        where chg_bu_no <> ''
                        group by chg_bu_no
                       having count(distinct ct_fld3) > 1
                      ) m
                      inner join fact_pcust d
                         on m.chg_bu_no = d.chg_bu_no
                  for xml path('')
              )
    set @Msg = @Proc+'�Ȥ��`���q���@�P�G'+Reverse(Substring(Reverse(@Msg), 2, Len(@Msg)))
    if @Msg is not null
       Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, 0, 1      
      
  end try
  begin catch
    set @Result = -1
    set @Msg = @Msg+'...(���~�T��:'+ERROR_MESSAGE()+')...(���~�C:'+Convert(Varchar(10), ERROR_LINE())+')'

    Exec uSP_Sys_Write_Log @Proc, @Msg, @strSQL, @Result
  end catch

  Return(@Result)
end
GO
