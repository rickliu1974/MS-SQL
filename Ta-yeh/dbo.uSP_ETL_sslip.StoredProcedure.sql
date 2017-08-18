USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_ETL_sslip]    Script Date: 07/24/2017 14:43:59 ******/
DROP PROCEDURE [dbo].[uSP_ETL_sslip]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create Proc [dbo].[uSP_ETL_sslip]
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_ETL_sslip
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [�ק� Trans_Log �T�����e]
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  
  /***************************************************************************************************
   �ѩ����`�A�ϥΡi��ڥD�ɡj�P�i�����ɡj�� BI ��h�i��ϥ��W���B�z�A�H�Q���d�߳t�ױo�H���ɡC
   ���Ϥ���K�i��ϥ��W�����L�{�A���P��椧���ΥH Table Name �[�H�Ϲj�C
   �ϥ��W���t�@�Ӧn�B�N�O�A���Y��¦����ɮ׶i��ϥ��W���ɡA���{���h�L���ק�Y�i�N��¦���[�H�ǤJ�C
  ***************************************************************************************************/
   
  Declare @Proc Varchar(50) = 'uSP_ETL_sslip'
  Declare @Cnt Int =0
  Declare @Cnt_Ori Int =0
  Declare @Err_Code Int = -1
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @slip_fg_Add Varchar(50)
  Declare @slip_fg_Less Varchar(50)
  Declare @Result Int = 0
  
  Set @slip_fg_Add  = (select code_end from Ori_Xls#Sys_Code With(NoLock) where code_class = '90' and code_begin ='+')
  Set @slip_fg_Less = (select code_end from Ori_Xls#Sys_Code With(NoLock) where code_class = '90' and code_begin ='-')
    
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Fact_sslip]') AND type in (N'U'))
  begin
     Set @Msg = '�R����ƪ� [Fact_sslip]'
     set @strSQL= 'DROP TABLE [dbo].[Fact_sslip]'

     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end
  
  begin try
    set @Msg = 'ETL sslip to [Fact_sslip]...'
    Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, 0

    ;With CTE_Q1 as (
       select sp_class, sp_slip_fg, sp_no,
              --�X�p���B(��)
              [Chg_sp_stot]= 
                 Case 
                   WHEN SP_SLIP_FG like @slip_fg_Add THEN Isnull(SP_TOT, 0)
                   WHEN SP_SLIP_FG like @slip_fg_Less THEN Isnull(-SP_TOT, 0)
                   else 0
                 END * isnull(sp_rate, 0) ,
              --��~�|(��)
              [Chg_SP_TAX] =
                 CASE 
                   WHEN SP_SLIP_FG like @slip_fg_Add THEN Isnull(SP_TAX, 0)
                   WHEN SP_SLIP_FG like @slip_fg_Less THEN Isnull(-SP_TAX, 0)
                   else 0
                 END * isnull(sp_rate, 0),
              --�w���I���B(��)
              [Chg_SP_PAY] =
                 CASE 
                   WHEN SP_SLIP_FG like @slip_fg_Add THEN Isnull(SP_PAY, 0)
                   WHEN SP_SLIP_FG like @slip_fg_Less THEN Isnull(-SP_PAY, 0)
                   else 0
                 END * isnull(sp_rate, 0), 
              --�������B(��)
              [Chg_SP_DIS] =
                 CASE 
                   WHEN SP_SLIP_FG like @slip_fg_Add THEN Isnull(-SP_DIS, 0)
                   WHEN SP_SLIP_FG like @slip_fg_Less THEN Isnull(SP_DIS, 0)
                   else 0
                 END * isnull(sp_rate, 0),
              -- �[�����B
              [Chg_SP_PAMT] = 
                 CASE 
                   WHEN SP_SLIP_FG like @slip_fg_Add THEN Isnull(SP_PAMT, 0)
                   WHEN SP_SLIP_FG like @slip_fg_Less THEN Isnull(-SP_PAMT, 0)
                 END * isnull(sp_rate, 0), 
              -- ����B
              [Chg_SP_MAMT] =
                 CASE 
                   WHEN SP_SLIP_FG like @slip_fg_Add THEN Isnull(-SP_MAMT, 0)
                   WHEN SP_SLIP_FG like @slip_fg_Less THEN Isnull(SP_MAMT, 0)
                 END * isnull(sp_rate, 0),
              -- 2015/03/02 Rickliu ���i����(sp_ave_p ���w�g�p��L�ײv�F)
              [Chg_sp_ave_p] =
                 CASE 
                   WHEN SP_SLIP_FG like @slip_fg_Add THEN Isnull(sp_ave_p, 0) 
                   WHEN SP_SLIP_FG like @slip_fg_Less THEN Isnull(-sp_ave_p, 0) 
                 END,
              -- 2015/03/02 Rickliu ��i�륭������(sp_mave_p ���w�g�p��L�ײv�F)
              [Chg_sp_mave_p] =
                 CASE 
                   WHEN SP_SLIP_FG like @slip_fg_Add THEN Isnull(sp_mave_p, 0) 
                   WHEN SP_SLIP_FG like @slip_fg_Less THEN Isnull(-sp_mave_p, 0) 
                 END
         from SYNC_TA13.dbo.sslip With(NoLock) 
        where sp_date >= '2012/12/01'
    ), CTE_Q2 as (
       select DD_CLASS,DD_SLIP_FG,DD_SPNO,DD_CTNO,
              [Chg_sp_dis2]= 
                 Case 
                   WHEN DD_SLIP_FG like @slip_fg_Add THEN sum(Isnull(-DD_DIS, 0))
                   WHEN DD_SLIP_FG like @slip_fg_Less THEN sum(Isnull(DD_DIS, 0))
                   else 0
                 END * isnull(DD_RATE, 1) ,
              [Chg_sp_dis_tax2]= 
                 Case 
                   WHEN DD_SLIP_FG like @slip_fg_Add THEN sum(Isnull(-DD_DETAX, 0))
                   WHEN DD_SLIP_FG like @slip_fg_Less THEN sum(Isnull(DD_DETAX, 0))
                   else 0
                 END * isnull(DD_RATE, 1) ,
              [Chg_sp_dis_tot2]= 
                 Case 
                   WHEN DD_SLIP_FG like @slip_fg_Add THEN sum(Isnull(-DD_DIS, 0)) + sum(Isnull(-DD_DETAX, 0))
                   WHEN DD_SLIP_FG like @slip_fg_Less THEN sum(Isnull(DD_DIS, 0)) + sum(Isnull(DD_DETAX, 0))
                   else 0
                 END * isnull(DD_RATE, 1) 
         from SYNC_TA13.dbo.SDISDT With(NoLock) 
        where DD_CLASS='2'
          and dd_date >= '2012/12/01'
        group by DD_CLASS,DD_SLIP_FG,DD_SPNO,DD_CTNO,DD_RATE
    )

    select M.[sp_class], M.[sp_slip_fg], M.[sp_date], [sp_pdate], 
           Rtrim(M.[sp_no]) as sp_no, 
           Rtrim([sp_ordno]) as sp_ordno,
           Rtrim([sp_ctno]) as sp_ctno, 
           Rtrim(D1.[ct_name]) as sp_ctname, 
           Rtrim(m.[sp_ctname]) as Ori_sp_ctname, 
           RTrim(Convert(Varchar(100), [sp_ctadd2])) as sp_ctadd2, 
           RTrim([sp_sales]) as sp_sales, 
           RTrim(D2.[E_Name]) as Chg_sales_Name, 
           
           [sp_pyno], 
           RTrim([sp_dpno]) as sp_dpno,
           RTrim([sp_maker]) as sp_maker, [sp_conv], [sp_tot], [sp_tax], [sp_dis], [sp_pay],
           [sp_pamt], [sp_cash], [sp_mamt], [sp_pay_kd], 
           RTrim([sp_rate_nm]) as sp_rate_nm, [sp_rate],
           [sp_ono], [sp_tno], [sp_wkno], [sp_tailno], [sp_tal_rec], [sp_ave_p],
           [sp_acapno], [sp_acspno], [sp_invoice], [sp_itot], [sp_inv_kd], [sp_tax_kd],
           [sp_i_date], [sp_invtype], [sp_dkind], [sp_wcd], [sp_rem]=RTrim(Convert(Varchar(100), [sp_rem])), [sp_npack],
           [sp_nokind], [sp_tfg], [sp_postfg], [sp_ntnpay], [sp_acspseq], [sp_addtax],
           [sp_trdspno], [sp_mave_p], [sp_caseno], [sp_mafkd], [sp_mafno], [sp_adjkd],
           [sp_surefg], [sp_suredt], [sp_sureman], [sp_mstno], [sp_bimtype],
           --������O
           [Chg_sp_class]=RTrim(D9.Code_Name),
           --��ں���
           [Chg_sp_slip_fg]=RTrim(D10.Code_Name), 
           -- 2014/06/24 Rick �s�W��ڧ���W��
           [Chg_sp_slip_name]=Substring(RTrim(D10.Code_Name Collate Chinese_Taiwan_Stroke_CI_AS), 1, 2)+Rtrim(M.[sp_no]),
           [Chg_sp_dp_name] = RTrim(D18.dp_name),
           -- �f����
           [Chg_sp_date_Year]=Year(sp_date),
           [Chg_sp_date_Quarter]=D16.cal_Quarter,
           [Chg_sp_date_Month]=Month(sp_date),
           [Chg_sp_date_YearWeek]=D16.cal_yearweek,
           -- 2014/5/15 Rick �s�W ��ڦ~��
           [Chg_sp_date_YM]=substring(convert(varchar(10), sp_date, 111), 1, 7),
           [Chg_sp_date_MD]=substring(convert(varchar(10), sp_date, 111), 6, 5),
           -- 2014/1/27 Rick �s�W ��ڤ�����P���_�W
           [Chg_sp_date_Weekno]=D16.cal_Week,
           [Chg_sp_date_Weekname]=D16.cal_week_name,
           [Chg_sp_date_Week_Range]=D16.cal_Week_Range,
           -- �b�ڤ��
           [Chg_sp_pdate_Year]=D17.cal_year,
           [Chg_sp_pdate_Quarter]=D17.cal_Quarter,
           [Chg_sp_pdate_Month]=D17.cal_month,
           [Chg_sp_pdate_YearWeek]=D17.cal_YearWeek,
           -- 2014/5/15 Rick �s�W ��ڦ~��
           [Chg_sp_pdate_YM]=substring(convert(varchar(10), sp_pdate, 111), 1, 7),
           [Chg_sp_pdate_MD]=substring(convert(varchar(10), sp_pdate, 111), 6, 5),
           -- 2014/1/27 Rick �s�W ��ڽдڤ�����P���_�W
           [Chg_sp_pdate_Weekno]=D17.cal_Week,
           [Chg_sp_pdate_Weekname]=D17.cal_week_name,
           [Chg_sp_pdate_Week_Range]=D17.cal_Week_Range,
          
           -- �|�O
           [Chg_sp_tax_kd]= RTrim(D7.Code_Name Collate Chinese_Taiwan_Stroke_CI_AS),
           /*
             Case sp_tax_kd
               when '1' then '���|'
               when '2' then '�s�|'
               when '3' then '�K�|'
               else '�L��'
             end,
           */
           -- �o�����O
           [Chg_sp_inv_name]= RTrim(D6.Code_Name),
           /*
             Case sp_inv_kd
               when '1' then '�T�p��'
               when '2' then '�G�p��'
               when '3' then '���Ⱦ�'
               else ''
             end,
           */
           -- �}�ߤ覡
           [Chg_sp_invtype_name]= RTrim(D19.Code_Name),
           /*
             Case sp_invtype
               when '1' then '���}'
               when '2' then '�H��}��'
               when '3' then '�妸�}��'
               else ''
             end,
           */
           --�~�ȲէO�s��
           [Chg_sp_mst_name] = Isnull(D8.mst_name, ''), 

           -- �X�p���B
           [Chg_sp_stot],
           -- ��~�|
           [Chg_SP_TAX],
    
           -- �X�p���B(�t�|)          
           [Chg_sp_stot_tax]=isnull([Chg_sp_stot], 0)+isnull([Chg_SP_TAX], 0),

           -- �w�I���B(��)
           [Chg_SP_PAY],
           -- �������B(��)
           [Chg_SP_DIS],
           -- �[�����B
           [Chg_SP_PAMT],
           -- ����B
           [Chg_SP_MAMT],
           -- 2015/03/02 Rickliu ���i����
           [Chg_sp_ave_p],
           -- 2015/03/02 Rickliu ��i�륭������
           [Chg_sp_mave_p],
           -- 2013/1/8 �]�|�L�ҫ��ܡA��o�����X���ťիh�N���|���B
           -- 2014/1/16 �ثe�w�g�אּ���o�Ȥ�o���������P�O�O�_�n���t�|�A���]�|���T�w���ǳ����ϥΨ즹���ҼȮɤ���A�Y�n��h�ШϥΥH�U Remark ����
           /*
           -- �����b��(�t�|)
           [Chg_sp_ntnpay_plus] = 
             Case 
               when RTrim(ct_invofg) = '1' then [sp_ntnpay] -- �T�p���|
               else [sp_ntnpay] * 1.05 -- �G�p�t�|
             End, 
           */
           -- �����b��(�t�|)
           [Chg_sp_ntnpay_plus] = 
             Case 
               when RTrim(sp_invoice) = '' then [sp_ntnpay] * 1.05
               else sp_ntnpay
             End * Isnull(sp_rate, 1), 

           --�`�p���B(�P�f���B-�h�f���B+��~�|-����)(�t�|)
           [Chg_sp_sum_tot]= [Chg_SP_STOT]+[Chg_SP_TAX]+[Chg_SP_DIS]+[Chg_SP_PAMT]-[Chg_SP_MAMT],

           --������J�`�������B
           [Chg_sp_dis2],
           --������J�`�����|��
           [Chg_sp_dis_tax2],
           --������J�`�������B + �|��
           [Chg_sp_dis_tot2],

           --�`�p���B2(�P�f���B-�h�f���B+��~�|-����)(�t�|)
           [Chg_sp_sum_tot2]= [Chg_SP_STOT]+[Chg_SP_TAX]+[Chg_sp_dis_tot2]+[Chg_SP_PAMT]-[Chg_SP_MAMT],

           --�������B���~�ȥ����ڪ��B(�t�|)
           [Chg_SP_Not_Recv_Tot]= [Chg_SP_STOT]+[Chg_SP_TAX]-[Chg_SP_DIS]+[SP_PAMT]-[Chg_SP_MAMT]-[Chg_SP_PAY],
            
           -- �ϰ�ؼлP�F���v
           [Chg_Area_Target]=Isnull(D14.Amt, 0),
 
           -- 2015/02/27 Rickliu �s�W�����P��q������, ���s���ĤE�X 1: �ʳf, 2: 3C, 3: �H��, 4:�N�u(OEM)
           [Chg_Dept_Cust_Chain_No] = RTrim(sp_dpno)+'-'+D1.Chg_Cust_Sale_Class,
           [Chg_Dept_Cust_Chain_Name] = RTrim(D18.dp_name)+'-'+D1.Chg_Cust_Sale_Class_Name collate Chinese_Taiwan_Stroke_CI_AS,

            -- ��ڦ��h�ڷ~�Z�F���v
            Chg_SP_Pay_Rate = case 
                                when sp_tot <> 0 then Chg_SP_PAY / Round(sp_tot + sp_tax, 0)
                                else 0
                              end,
/**********************************************************************************************************************************************************
  �Ȥ�/�t�Ӱ򥻸�ư�
 **********************************************************************************************************************************************************/
           [###pcust###]='### pcust ###',
           D1.*,

/**********************************************************************************************************************************************************
  ���u�򥻸�ư�
 **********************************************************************************************************************************************************/
           [###pmploy###]='### pemploy ###',
           D2.*,
           sslip_update_datetime = getdate(),
           sslip_timestamp = m.timestamp_column
           into Fact_sslip
      from SYNC_TA13.dbo.sslip M With(NoLock) 
/* 20140415 Rickliu--�L�Ҫ��(6)�U���]�H�ܦb�L�B�A���ݦۤv���w�s�ҥH�������ơA(7)�^�f��h�����P�f�A�ҥH�����t�ơC
                               WHEN SP_SLIP_FG in ('0', '2', '5', '7', '8', 'C', 'R') THEN Isnull(SP_TOT, 0)
                               WHEN SP_SLIP_FG in ('1', '3', '4', '6', '9') THEN Isnull(-SP_TOT, 0)
*/
           inner join CTE_Q1 D 
             on M.sp_class=D.sp_class
            and M.sp_slip_fg=D.sp_slip_fg
            and M.sp_no=D.sp_no
           -- �Ȥ�/�t�Ӹ��
           left join Fact_pcust D1 With(NoLock) 
             on M.sp_ctno=D1.ct_no 
            and D1.ct_class=
                  Case 
                    when M.sp_class in ('1', '2', '3', '8') Or (M.sp_class = '6' and Len(M.sp_ctno) = 9) then '1'
                    when M.sp_class in ('0', '4') Or (M.sp_class ='6' and Len(M.sp_ctno) = 5) then '2'
                    else ''
                  end 
           -- ���u���
           left join Fact_pemploy D2 With(NoLock) 
             on M.sp_sales=D2.e_no -- �~�ȭ�
           -- ���u���
           left join Fact_pemploy D4 With(NoLock) 
             on M.sp_maker=D4.e_no -- �s��H��
           -- ���e��H
           left join Fact_pcust D5 With(NoLock) 
             on M.sp_pyno=D5.ct_no 
            and D5.ct_class = '1'

           left join Ori_Xls#Sys_Code D6 With(NoLock) 
             on D6.code_class ='13' 
            and Convert(Varchar(1), Convert(Int, M.sp_inv_kd)) = D6.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS 

           left join Ori_Xls#Sys_Code D7 With(NoLock) 
             on D6.code_class ='12' 
            and Convert(Varchar(1), Convert(Int, M.sp_tax_kd)) = D6.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS 

           -- �էO���
           left join SYNC_TA13.dbo.masterm D8 
             on M.sp_mstno=D8.Mst_no

           left join Ori_Xls#Sys_Code D9 With(NoLock) 
             on D9.code_class ='7' 
            and M.sp_class = D9.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS 

           left join Ori_Xls#Sys_Code D10 With(NoLock) 
             on D10.code_class ='8' 
            and M.sp_class = D10.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS 
            and M.sp_slip_fg = D10.Code_End Collate Chinese_Taiwan_Stroke_CI_AS

           -- �ӤH�ؼ�
           left join uV_Sales_Base D13 With(NoLock) 
             on D13.bo_class='1' 
            and M.sp_sales=D13.Bo_no
            and Year(M.sp_date)=D13.bo_yr
            and Month(M.sp_date)=bo_Month
            
           -- �ϰ�ؼ�
           left join Ori_xls#Target_Area_Sales D14 With(NoLock) 
             on Year(M.sp_date)=D14.[Year]
            and Month(M.sp_date)=D14.[Month]
            and D1.ct_loc Collate Chinese_PRC_BIN=D14.Area Collate Chinese_PRC_BIN

           -- 2014/1/27 Rick �s�W ��ڤ�����P���_�W
           left join calendar D16 With(NoLock) on M.SP_DATE = D16.cal_date
           
           -- 2014/1/27 Rick �s�W ��ڽдڤ�����P���_�W
           left join calendar D17 With(NoLock) on M.SP_PDATE = D17.cal_date
           
           -- 2015/02/27 Rickliu �s�W�����N�X
           left join SYNC_TA13.dbo.pdept D18 With(NoLock) on M.sp_dpno = D18.dp_no
           -- 2015/03/06 Rickliu �s�W �o���}�ߤ覡
           left join Ori_Xls#Sys_Code D19 With(NoLock) 
             on D19.code_class ='14' 
            and Convert(Varchar(1), Convert(Int, M.sp_invtype, 0)) = D19.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS 
           
           --2016/09/13 NanLiao�W�[
           left join CTE_Q2 D20
             on DD_CLASS='2'
            and DD_SLIP_FG = m.SP_SLIP_FG
            and DD_SPNO = m.SP_NO
            and DD_CTNO = m.SP_CTNO
    where m.sp_date >= Convert(Varchar(5), year(DateAdd(year, -5, getdate())))+ '/01/01'
    
    /*==============================================================*/
    /* Index: sslip_Timestamp                                      */
    /*==============================================================*/
    --Set @Msg = '�إ߯��� [Fact_sslip.sslip_Timestamp]'
    --Print @Msg
    --if exists (select * from sys.indexes where object_id = object_id('[dbo].[fact_sslip]') and name = 'pk_sslip')
    --   alter table [dbo].[pk_sslip] drop constraint [pk_sslip]

    --alter table [dbo].[fact_sslip] add  constraint [pk_sslip] primary key nonclustered ([sslip_timestamp] asc) with 
    --(pad_index  = off, statistics_norecompute  = off, sort_in_tempdb = off, ignore_dup_key = off, online = off, allow_row_locks  = on, allow_page_locks  = on) on [primary]


    /*==============================================================*/
    /* Index: caseno                                                */
    /*==============================================================*/
    Set @Msg = '�إ߯��� [Fact_sslip.caseno]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslip') and name  = 'caseno' and indid > 0 and indid < 255) drop index dbo.fact_sslip.caseno
    create index caseno on dbo.fact_sslip (sp_caseno) with fillfactor= 30 on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: class                                                 */
    /*==============================================================*/
    Set @Msg = '�إ߯��� [Fact_sslip.class]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslip') and name  = 'class' and indid > 0 and indid < 255) drop index dbo.fact_sslip.class
    create index class on dbo.fact_sslip (sp_class) with fillfactor= 30 on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: ctno                                                  */
    /*==============================================================*/
    Set @Msg = '�إ߯��� [Fact_sslip.ctno]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslip') and name  = 'ctno' and indid > 0 and indid < 255) drop index dbo.fact_sslip.ctno
    create index ctno on dbo.fact_sslip (sp_ctno, sp_class) with fillfactor= 30 on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: mafno                                                 */
    /*==============================================================*/
    Set @Msg = '�إ߯��� [Fact_sslip.mafno]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslip') and name  = 'mafno' and indid > 0 and indid < 255) drop index dbo.fact_sslip.mafno
    create index mafno on dbo.fact_sslip (sp_mafno, sp_mafkd) with fillfactor= 30 on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: ordno                                                 */
    /*==============================================================*/
    Set @Msg = '�إ߯��� [Fact_sslip.ordno]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslip') and name  = 'ordno' and indid > 0 and indid < 255) drop index dbo.fact_sslip.ordno
    create index ordno on dbo.fact_sslip (sp_ordno, sp_class) with fillfactor= 30 on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: slip_fg                                               */
    /*==============================================================*/
    Set @Msg = '�إ߯��� [Fact_sslip.slip_fg]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslip') and name  = 'slip_fg' and indid > 0 and indid < 255) drop index dbo.fact_sslip.slip_fg
    create index slip_fg on dbo.fact_sslip (sp_slip_fg) with fillfactor= 30 on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: spctp                                                 */
    /*==============================================================*/
    Set @Msg = '�إ߯��� [Fact_sslip.spctp]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslip') and name  = 'spctp' and indid > 0 and indid < 255) drop index dbo.fact_sslip.spctp
    create index spctp on dbo.fact_sslip (sp_pdate, sp_ctno, sp_class) with fillfactor= 30 on "PRIMARY"
    
    
    /*==============================================================*/
    /* Index: spda                                                  */
    /*==============================================================*/
    Set @Msg = '�إ߯��� [Fact_sslip.spda]'
    Print @Msg
    if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslip') and name  = 'spda' and indid > 0 and indid < 255) drop index dbo.fact_sslip.spda
    create index spda on dbo.fact_sslip (sp_date, sp_slip_fg) with fillfactor= 30 on "PRIMARY"

  end try
  begin catch
    set @Result = @Err_Code
    set @Msg = @Proc+'...(���~�T��:'+ERROR_MESSAGE()+', '+@Msg+')...(���~�C:'+Convert(Varchar(10), ERROR_LINE())+')'

    Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, @Cnt
  end catch
  Return(@Cnt)
end
GO
