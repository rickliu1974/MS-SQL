USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[uSP_ETL_sslpdt]    Script Date: 08/18/2017 17:43:39 ******/
DROP PROCEDURE [dbo].[uSP_ETL_sslpdt]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Proc [dbo].[uSP_ETL_sslpdt]
as
begin
  /*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
   Program Name: uSP_ETL_sslpdt
   Create Date: 2013/03/01
   Creator: Rickliu
   Updated Date: 2013/09/06 [�ק� Trans_Log �T�����e]
  *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/  
  Declare @Proc Varchar(50) = 'uSP_ETL_sslpdt'
  Declare @Cnt Int =0
  Declare @Cnt_Ori Int =0
  Declare @Master_Stock_Cnt Int = 0
  Declare @Msg Varchar(Max) =''
  Declare @strSQL Varchar(Max)=''
  Declare @CR Varchar(4) = char(13)+char(10)
  Declare @slip_fg_Add Varchar(50)
  Declare @slip_fg_Less Varchar(50)
  
  Set @slip_fg_Add  = (select code_end from Ori_Xls#Sys_Code where code_class = '90' and code_begin ='+')
  Set @slip_fg_Less = (select code_end from Ori_Xls#Sys_Code where code_class = '90' and code_begin ='-')
  
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Fact_sslpdt]') AND type in (N'U'))
  begin
     Set @Msg = '�R����ƪ� [Fact_sslpdt]'
     set @strSQL= 'DROP TABLE [dbo].[Fact_sslpdt]'

     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end

  
  set @Msg = 'ETL sslpdt to [Fact_sslpdt]...'
  begin try
    -- 2017/04/05 Rickliu �����D����A���Ӹ�ƳB�z�t���ܱo�ܺC�A��P���ӬO�S�����ҿת��D�O�� Index�A�t�~��� CTE ���g�覡�A�ݬO�_��[�ֳt�סC
    ;with CTE_Q1 as (
      Select sd_class,
             sd_slip_fg,
             sd_ctno,
             Max(sd_no) AS Max_SP_NO, 
             Max(CONVERT(varchar(10), sd_date, 111)) AS sp_date, 
             min(sd_price) as sd_price,
             sd_skno,
             'Y' as Flag -- Rickliu 2017/06/06  �o�̬O���̫�@������A�����]�t�ث~�B���B�p�� 1����
        From SYNC_TA13.dbo.sslpdt With(NoLock)
       where 1=1
         and sd_name is not null
         and sd_date >= '2012/12/01'
         and sd_slip_fg = '2'
         and isnull(sd_csno, '') <> ''
         and sd_sendfg = 0
         and sd_price > 1
         and sd_sendfg = 0
       Group by sd_class, sd_slip_fg, sd_ctno, sd_skno
    ), CTE_Q2 as (
      Select sd_class,
             sd_slip_fg,
             --2016/09/09 Rickliu �]��´�X�s�A�W�[�����q���A�]���N�쥻 2~5 �אּ 2~6
             substring(sd_ctno, 1, 6) as BU_NO,
             Max(sd_no) AS Max_SP_NO, 
             Max(CONVERT(varchar(10), sd_date, 111)) AS sp_date, 
             min(sd_price) as sd_price,
             sd_skno,
             'Y' as Flag -- Rickliu 2017/06/06  �o�̬O���̫�@������A�����]�t�ث~�B���B�p�� 1����
        From SYNC_TA13.dbo.sslpdt With(NoLock) 
       where 1=1
         and sd_name is not null
         and sd_date >= '2012/12/01'
         and sd_slip_fg = '2'
             --2016/09/09 Rickliu �]��´�X�s�A�W�[�����q���A�]���N�쥻 2~5 �אּ 2~6
         and isnull(sd_csno, '') <> ''
         and sd_sendfg = 0
         and sd_price > 1
         and sd_sendfg = 0
       Group by sd_class, sd_slip_fg, substring(sd_ctno, 1, 6), sd_skno
    )

    select m.[sd_class], m.[sd_slip_fg], m.[sd_date], d9.sp_pdate, 
           Rtrim(m.[sd_no]) as sd_no, 
           Rtrim(m.[sd_ctno]) as sd_ctno, 
           Rtrim(m.[sd_skno]) as sd_skno,
           Rtrim(D4.[sk_name]) as sd_name, m.[sd_whno], m.[sd_whno2], m.[sd_qty], m.[sd_price], m.[sd_dis],
           m.[sd_stot], m.[sd_spec], [sd_rem]=Convert(Varchar(255), m.[sd_rem]), m.[sd_unit], m.[sd_unit_fg], m.[sd_ave_p],
           m.[sd_pave_p], m.[sd_rate], m.[sd_val1], m.[sd_val2], m.[sd_rqty], m.[sd_nqty],
           m.[sd_bmno], m.[sd_nokind], m.[sd_sekind], m.[sd_ordno], m.[sd_postfg], m.[sd_acspseq],
           m.[sd_csno], m.[sd_csrec], m.[sd_sendfg], m.[sd_mafkd], m.[sd_mave_p], m.[sd_mstot],
           m.[sd_adjkd], m.[sd_surefg], m.[sd_seqfld], m.[sd_lotno],

           rtrim(sp_sales) as Chg_sp_sales, -- Rickliu 2014/3/17 �s�W�~�ȭ�
           rtrim(Chg_sales_Name) as Chg_sales_Name, -- Rickliu 2014/5/14 �s�W�~�ȭ��W��

           rtrim(ct_sales) as ct_sales, -- Rickliu 2014/07/03 �s�W�Ȥ���ݷ~�ȭ�
           rtrim(Chg_ct_sales_Name) as Chg_ct_sales_Name, -- Rickliu 2014/07/03 �s�W�Ȥ���ݷ~�ȭ��W��
           rtrim(ct_name) as Chg_ct_name, --Rickliu 2014/1/25 �s�W ���q�W��
           rtrim(ct_sname) as Chg_ct_sname,--Rickliu 2014/1/25 �s�W ���q²��

           [Chg_sd_whno_name]=Rtrim(Isnull(D2.wh_name,  '')),
           [Chg_sd_class]=Rtrim(D5.Code_Name),
           [Chg_sd_slip_fg]=Rtrim(D6.Code_Name), 
           /***[��ں���]******************************************************************************
            =0�i�f��(+)     =1�i�h��(-)     =2�X�f��(+)     =3�X�h��(-)     =4�ɳf��(-)     =5�ٳf��(+)
            =6�U���(+)     =7�U�^��(-)     =8�J�w��(+)     =9�X�w��(-)     =A�վ��        =B�ռ���
            =C�A�ȳ�(+)     =R�ɤJ��(+)     =S���ٳ�        =T�L�I��
           *******************************************************************************************/
           [Chg_sd_qty]= 
             Case 
               WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN IsNull(m.SD_QTY, 0)
               WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.SD_QTY, 0)
               else 0
             END, 
           [Chg_sd_sale_qty]= 
             Case 
               WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN IsNull(m.SD_QTY, 0)
               else 0
             END, 
           [Chg_sd_ret_qty]= 
             Case 
               WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.SD_QTY, 0)
               else 0
             END, 
           [Chg_sd_stot]= 
             Case 
               WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN Isnull(m.SD_STOT, 0)
               WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.SD_STOT, 0)
               else 0
             END * isnull(d9.sp_rate, 1),
           [Chg_sd_ave_p]=
             Case 
               WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN Isnull(m.sd_ave_p, 0) 
               WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.sd_ave_p, 0)
               else 0
             END * isnull(d9.sp_rate, 1),
           
           
           -- �O�_���P���D�P�ӫ~
           -- 2017/08/07 Rickliu ��H�ӫ~�򥻸�ƨ��o�O�_���D�P
           --[Chg_sale_Month_Master] = 
           --   Case 
           --     when D10.SK_NO is not Null and D10.Sale_Year= D9.Chg_sp_date_year and D10.Sale_Month=D9.Chg_sp_date_month
           --      Then 'Y'
           --     else 'N'
           --   end,
           -- 2017/08/07 Rickliu ��H�ӫ~�򥻸�ƨ��o�O�_���D�P
           --[Chg_sale_Month_MasterName] =
           --   Case 
           --     when D10.SK_NO is not Null and D10.Sale_Year= D9.Chg_sp_date_year and D10.Sale_Month=D9.Chg_sp_date_month
           --       Then D10.sk_Mname
           --     else 'NA'
           --   end,
           -- 2017/08/07 Rickliu ��H�ӫ~�򥻸�ƨ��o�O�_���D�P
           -- �O�_���дڤ�D�P�ӫ~
           --[Chg_inv_Month_Master] = -- 2013/6/1 �s�W�P��D�P�W��
           --   Case 
           --     when D8.SK_NO is not Null and D8.Sale_Year= D9.Chg_sp_pdate_year and D8.Sale_Month=D9.Chg_sp_pdate_month
           --       Then 'Y'
           --     else 'N'
           --   end,
           -- 2017/08/07 Rickliu ��H�ӫ~�򥻸�ƨ��o�O�_���D�P
           --[Chg_inv_Month_MasterName] = -- 2013/6/1 �s�W�дڥD�P�W��
           --   Case 
           --     when D8.SK_NO is not Null and D8.Sale_Year= D9.Chg_sp_pdate_year and D8.Sale_Month=D9.Chg_sp_pdate_month
           --       Then D8.sk_Mname
           --     else 'NA'
           --   end,
           -- 2013/1/17 ��z�P���`�T�{�A��f�~������󵥩󤤽L���ɤ~�{�C�~�Z(�������T�{���t�p�L�����~�Z)�A��᪺�~�Z�������̦��e�{�C
           -- 2013/1/18 ��z�P�]�ȪL�ҽT�{�A�Ҧ��P�泣�O�H��@�~���馩�A�ҥH�u�n���D�ɧ馩�i��p��Y�i�C
           [Chg_sd_sale_overmid_tot]= --(�����B�����|���B)
             Case 
               WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN 
                 Case 
                   when m.[sd_price] >= [s_apice4] then Isnull(m.SD_STOT, 0)
                   else 0
                 end 
               WHEN m.SD_SLIP_FG like '[13479]' THEN 
                 Case 
                   when m.[sd_price] >= [s_apice4] then Isnull(-m.SD_STOT, 0)
                   else 0
                 end
               else 0
             END * isnull(d9.sp_rate, 1),
           [Chg_sd_sale_overmid_dis]= --(�����B�����|���B)
             Case 
               WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN 
                 Case 
                   when m.[sd_price] >= [s_apice4] then Isnull(m.sd_dis, 0)
                   else 0
                 end
               WHEN m.SD_SLIP_FG like '[13479]' THEN 
                 Case 
                   when m.[sd_price] >= [s_apice4] then Isnull(-m.sd_dis, 0)
                   else 0
                 end
               else 0
             END * isnull(d9.sp_rate, 1),
           -- 2013/1/17 ��z�P���`�T�{�A��f�~������󵥩�p�L���ɫh�t�~�{�C�E�y�~�Z�A��᪺�~�Z�������̦��e�{�C
           [Chg_sd_sale_oversmall_tot]= --(�����B�����|���B)
             Case 
               WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN 
                 Case 
                   when m.[sd_price] >= [s_apice5] then 0
                   else Isnull(m.SD_STOT, 0)
                 end
               WHEN m.SD_SLIP_FG like '[13479]' THEN 
                 Case 
                   when m.[sd_price] > [s_apice5] then 0
                   else Isnull(-m.SD_STOT, 0)
                 end
               else 0
             END * isnull(d9.sp_rate, 1),
           [Chg_sd_sale_oversmall_dis]= --(�����B�����|���B)
             Case 
               WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN 
                 Case 
                   when m.[sd_price] >= [s_apice5] then 0
                   else Isnull(sd_dis, 0)
                 end
               WHEN m.SD_SLIP_FG like '[13479]' THEN 
                 Case 
                   when m.[sd_price] > [s_apice5] then 0
                   else Isnull(-sd_dis, 0)
                 end
               else 0
             END * isnull(d9.sp_rate, 1),
           -- ���
           [Chg_sd_price] = 
             Case 
               WHEN m.SD_SLIP_FG collate Chinese_Taiwan_Stroke_CI_AS like '[02568ABCRST]' THEN Isnull(m.sd_price, 0)
               WHEN m.SD_SLIP_FG collate Chinese_Taiwan_Stroke_CI_AS like '[13479]' THEN Isnull(-m.sd_price, 0)
               else 0
             END * isnull(d9.sp_rate, 1), 
           -- �t�|���
           [Chg_sd_price_tax] = 
             Convert(decimal (18,4), 
               Case 
                 WHEN m.SD_SLIP_FG collate Chinese_Taiwan_Stroke_CI_AS like '[02568ABCRST]' THEN Isnull(m.sd_price, 0)
                 WHEN m.SD_SLIP_FG collate Chinese_Taiwan_Stroke_CI_AS like '[13479]' THEN Isnull(-m.sd_price, 0)
                 else 0
               END *
               -- 2015/04/23 Rickliu �|ĳ�Mĳ�N�~�Z���H�t�|�A�b���h�N�C�����Ӫ��|���[�H�p��, �è��p���I�ĥ|��, �H�קK���ӵ|���P���Y�|�����Ҹ��t
               Case 
                  when SP_Tax <> 0 then 1.05
                  else 0
               end * 
               isnull(d9.sp_rate, 1)), 
           -- �p�p �t�|��
           [Chg_sd_stot_tax] = 
             Convert(decimal (18,4), 
               Case 
                 WHEN m.SD_SLIP_FG collate Chinese_Taiwan_Stroke_CI_AS like '[02568ABCRST]' THEN Isnull(m.SD_STOT, 0)
                 WHEN m.SD_SLIP_FG collate Chinese_Taiwan_Stroke_CI_AS like '[13479]' THEN Isnull(-m.SD_STOT, 0)
                 else 0
               END * 1.05 * isnull(d9.sp_rate, 1)), 
           -- �p�p�|��
           [Chg_sd_tax] = 
             Convert(decimal (18,4), 
               Case 
                 WHEN m.SD_SLIP_FG collate Chinese_Taiwan_Stroke_CI_AS like '[02568ABCRST]' THEN Isnull(m.SD_STOT, 0)
                 WHEN m.SD_SLIP_FG collate Chinese_Taiwan_Stroke_CI_AS like '[13479]' THEN Isnull(-m.SD_STOT, 0)
                 else 0
               END  * 0.05 * isnull(d9.sp_rate, 1)), 
           -- 2014/10/02 Rickliu �s�W�C�󤤽L��(��� - ���L��)
           Chg_Low_sd_price = IsNull(m.sd_price, 0) - IsNull(s_price4, 0),
           
           -- 2015/03/05 Rickliu �s�W�����p�p (��즨�� * �ƶq)
           Chg_Cost_stot = 
              Case 
                WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN Isnull(m.sd_ave_p, 0) 
                WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.sd_ave_p, 0)
                else 0
              END * isnull(d9.sp_rate, 1) * sd_Qty,

           -- 2014/10/02 Rickliu �s�W��Q (��� - ��즨��) * �ƶq
           Chg_Profit = 
             (Case 
                WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN Isnull(m.sd_price, 0) 
                WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.sd_price, 0)
                else 0
              END * isnull(d9.sp_rate, 1)  - 
              Case 
                WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN Isnull(m.sd_ave_p, 0)
                WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.sd_ave_p, 0)
                else 0
              END * isnull(d9.sp_rate, 1)) * Isnull(sd_Qty, 0),

           -- 2014/10/02 Rickliu �s�W��Q�v (��Q / �P�f�b�B)
           Chg_Profit_Rate =
              Case 
                WHEN Isnull(m.sd_price, 0) * Isnull(sd_Qty, 0) = 0 THEN 0
                WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN 
                     ((Isnull(m.sd_price, 0)   * isnull(d9.sp_rate, 1) * Isnull(sd_Qty, 0))   - 
                      (Isnull(m.sd_ave_p, 0) * isnull(d9.sp_rate, 1) * Isnull(sd_Qty, 0)))  / 
                      (Isnull(m.sd_price, 0)   * isnull(d9.sp_rate, 1) * Isnull(sd_Qty, 0))
                WHEN m.SD_SLIP_FG like '[13479]' THEN 
                     ((Isnull(m.sd_price, 0)   * isnull(d9.sp_rate, 1) * Isnull(-sd_Qty, 0))  - 
                      (Isnull(m.sd_ave_p, 0) * isnull(d9.sp_rate, 1) * Isnull(-sd_Qty, 0))) / 
                      (Isnull(m.sd_price, 0)   * isnull(d9.sp_rate, 1) * Isnull(-sd_Qty, 0))
                else 0
              end,
              -- 2015/01/19 Rickliu �s�W�P�⦨�� (�зǦ��� * �ƶq)
           Chg_save_stot =
             Case 
               WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN IsNull(sk_save, 0) * Isnull(sd_Qty, 0)
               WHEN m.SD_SLIP_FG like '[13479]' THEN IsNull(sk_save, 0) * Isnull(-sd_Qty, 0)
               else 0
             end,
           -- �����u������(���̫Ȥ᦬�ڤ覡�����O�i���u��) Chg_SP_Dis �w�g���t�ײv
           Chg_sd_dis_avg =
             Round(
               Case
                 when Chg_SP_Dis = 0 then 0
                 else Chg_SP_Dis / (select count(1) from SYNC_TA13.dbo.sslpdt where sd_slip_fg = m.sd_slip_fg and sd_no = m.sd_no) 
               end, 4),

           Chg_sd_dis_rate =
            Convert(decimal (18,4), 
              Case
                 when Chg_SP_Dis = 0 then 0
                 else (Case 
                         WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN Isnull(m.SD_STOT, 0)
                         WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.SD_STOT, 0)
                         else 0
                       END * 
                       isnull(d9.sp_rate, 1) *
                       Case
                         when sp_tax <> 0 then 1.05
                         else 1
                       end
                      ) / Chg_SP_Dis  
                 --select count(*) from SYNC_TA13.dbo.sslpdt where sd_slip_fg = m.sd_slip_fg and sd_no = m.sd_no) 
               end),
           -- ��ڦ��h�ڷ~�Z�F���v
           Chg_SP_Pay_Rate,
           -- �ӫ~���h�ڷ~�Z���B(�t�|�B�t����)
           Chg_sd_Pay_Sale =
             Convert(decimal (18,4), 
               Case
                 when (SP_TOT = 0) Then 0 -- �w�����B
                 else (m.sd_stot * (case when sp_tax <> 0  then 1.05 else 1 end) * isnull(sp_rate, 1)) * Chg_SP_Pay_Rate
               end
               ),
           -- 2015/10/08 Rickliu  �ӫ~���h�ڷ~�Z��Q(�t�|�B�t����)
           Chg_sd_Pay_Profit =
             Convert(decimal (18,4), 
               Case
                 when (SP_TOT = 0) Then 0 -- �w�����B
                 else ((Case 
                          WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN Isnull(m.sd_price, 0) 
                          WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.sd_price, 0)
                          else 0
                        END * isnull(d9.sp_rate, 1)  - 
                        Case 
                          WHEN m.SD_SLIP_FG like '[02568ABCRST]' THEN Isnull(m.sd_ave_p, 0)
                          WHEN m.SD_SLIP_FG like '[13479]' THEN Isnull(-m.sd_ave_p, 0)
                          else 0
                        END * isnull(d9.sp_rate, 1)) * Isnull(sd_Qty, 0)
                        * (case when sp_tax <> 0  then 1.05 else 1 end) * isnull(sp_rate, 1)) 
                        * Chg_SP_Pay_Rate
               end
               ),
/**********************************************************************************************************************************************************
  ���Y��ư�
 **********************************************************************************************************************************************************/
           [###sslip###]='### sslip ###',
           -- 2015/03/02 Rickliu �s�W���i����
           [Chg_sp_ave_p],
           -- 2015/03/02 Rickliu ��i�륭������
           [Chg_sp_mave_p],
           --2014/05/15 Rickliu -- �Ȧb���Ӫ��Ĥ@���e�{�P�f���B
           [Chg_SP_Sale_Tot]=
             case
               when m.sd_slip_fg = '2' and sd_seqfld=1 then sp_tot * isnull(sp_rate, 1)
             else 0
             end,
           --2014/05/15 Rickliu -- �Ȧb���Ӫ��Ĥ@���e�{�h�f���B
           [Chg_SP_Ret_Tot]=
             case
               when m.sd_slip_fg = '3' and sd_seqfld=1 then sp_tot * isnull(sp_rate, 1)
               else 0
             end,
           --2014/05/15 Rickliu-- �Ȧb���Ӫ��Ĥ@���e�{�|�����B(�������Y��ơA�w�g���p��L���O)
           [Chg_sp_tax]=
             Case
               when sd_seqfld=1 then Chg_sp_tax * isnull(sp_rate, 1)
               else 0
             end,
           --2014/05/15 Rickliu-- �Ȧb���Ӫ��Ĥ@���e�{�������B(�������Y��ơA�w�g���p��L���O)
           [Chg_sp_dis]=
             Case
               when sd_seqfld=1 then Chg_sp_dis * isnull(sp_rate, 1)
               else 0
             end,
------------------------------------------------------------------------------------------------------------------------           
------------------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------------------------           
           --2016/03/31 NanLiao-- �Ȧb���Ӫ��Ĥ@���e�{�������B�t�|(�������Y��ơA�w�g���p��L���O)
           [Chg_sp_dis_tax]=
             Case
                when m.sd_seqfld=1 
                then Chg_sp_dis * isnull(sp_rate, 1) * 1.05
               else 0
             end,

           --2016/03/31 NanLiao-- �b���Ӫ��e�{����Tag(�������Y��ơA�w�g���p��L���O)
           [Chg_sp_dis_flg] =
             Case
               when SUBSTRING(m.sd_ctno, 9, 1) = '2' then 'AB'
               else 'AA'
             end,
------------------------------------------------------------------------------------------------------------------------           
------------------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------------------------     
           --2016/09/13 NanLiao-- �Ȧb���Ӫ��Ĥ@���e�{�������B(�����������ơA�w�g���p��L���O)
           [Chg_sp_dis2]=
             Case
               when sd_seqfld=1 then Chg_sp_dis2
               else 0
             end,

           --2016/09/13 NanLiao-- �Ȧb���Ӫ��Ĥ@���e�{�����|��(�����������ơA�w�g���p��L���O)
           [Chg_sp_dis_tax2]=
             Case
               when sd_seqfld=1 then Chg_sp_dis_tax2
               else 0
             end,
             
           --2016/09/13 NanLiao-- �Ȧb���Ӫ��Ĥ@���e�{�������B�t�|(�����������ơA�w�g���p��L���O)
           [Chg_sp_dis_tot2]=
             Case
               when sd_seqfld=1 then Chg_sp_dis_tot2
               else 0
             end,

------------------------------------------------------------------------------------------------------------------------           
------------------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------------------------           
           --2014/05/15 Rickliu-- �Ȧb���Ӫ��Ĥ@���e�{�X�p���B(�������Y��ơA�w�g���p��L���O)
           [Chg_sp_sum_tot]=
             Case
               when sd_seqfld=1 then Chg_sp_sum_tot
               else 0
             end,
           --2014/05/15 Rickliu-- �Ȧb���Ӫ��Ĥ@���e�{�����I���B(�������Y��ơA�w�g���p��L���O)
           [Chg_Non_SP_Pay]=
             Case
               when sd_seqfld=1 then Chg_sp_sum_tot - Chg_SP_PAY
               else 0
             end,
           -- 2015/02/26 Rickliu �W�[ �[����B���
           -- �[�����B
           [Chg_SP_PAMT]=
             Case
               when sd_seqfld=1 then [Chg_SP_PAMT]
               else 0
             end,
           -- ����B
           [Chg_SP_MAMT]=
             Case
               when sd_seqfld=1 then [Chg_SP_MAMT]
               else 0
             end,
           -- 2014/06/24 Rickliu �s�W��ڧ���W��
           [Chg_sp_slip_name],
          
           --�D�ɪ��P����
           [Chg_sp_date_Year],
           [Chg_sp_date_Quarter],
           [Chg_sp_date_Month],
           [Chg_sp_date_YearWeek],
           [Chg_sp_date_YM],
           [Chg_sp_date_MD],
           -- 2014/1/27 Rickliu �s�W ��ڤ�����P���_�W
           [Chg_sp_date_Weekno],
           [Chg_sp_date_Weekname],
           [Chg_sp_date_Week_Range],

           --�D�ɪ��дڤ��
           Chg_sp_pdate_Year, 
           Chg_sp_pdate_Quarter, 
           Chg_sp_pdate_Month, 
           Chg_sp_pdate_YearWeek, 
           [Chg_sp_pdate_YM],
           [Chg_sp_pdate_MD],
           -- 2014/1/27 Rickliu �s�W ��ڽдڤ�����P���_�W
           [Chg_sp_pdate_Weekno],
           [Chg_sp_pdate_Weekname],
           [Chg_sp_pdate_Week_Range],
           -- 2015/03/07 Rickliu �s�W ���Y����ڪ���
           [sp_maker],
           [sp_rem],
           -- 2015/03/07 Rickliu ���y�ثe�O�H�ռ���ڧ@�������h�f��]�Ҧb�A�æb���Y��������g��ڽs���A�Ӧb����������g�h�f��]�A�]�����F�έp�h�f��]�A
           -- �b���s�W�S�����
           [Chg_Logistics_Rej_flg] = 
             case
               when Isnumeric(Substring(sp_rem, 5, 10)) = 1 and Len(RTrim(Substring(sp_rem, 5, 10))) = 10 and Isnull(D13.Code_Begin, '') like '[0-7RSC]'
               then Isnull(D13.Code_Begin, '')
               else ''
             end,               
           [Chg_Logistics_Rej_sp_name] = 
             case
               when Isnumeric(Substring(sp_rem, 5, 10)) = 1 and Len(RTrim(Substring(sp_rem, 5, 10))) = 10 and Isnull(D13.Code_Begin, '') like '[0-7RSC]'
               then RTrim(Substring(Sp_rem, 1, 3))
               else ''
             end,               
           [Chg_Logistics_Rej_sp_no] = 
             case
               when Isnumeric(Substring(sp_rem, 5, 10)) = 1 and Len(RTrim(Substring(sp_rem, 5, 10))) = 10 and Isnull(D13.Code_Begin, '') like '[0-7RSC]'
               then Isnull(Substring(sp_rem, 5, 10), '')
               else ''
             end ,
           [Chg_Logistics_Rej_sp_date] = 
             case
               when Isnumeric(Substring(sp_rem, 5, 10)) = 1 and Len(RTrim(Substring(sp_rem, 5, 10))) = 10 and Isnull(D13.Code_Begin, '') like '[0-7RSC]'
               then 
                   (select sp_date
                      from fact_sslip 
                     where sp_slip_fg = Isnull(D13.Code_Begin, '') Collate Chinese_Taiwan_Stroke_CI_AS 
                       and sp_no = Substring(sp_rem, 5, 10) Collate Chinese_Taiwan_Stroke_CI_AS)
               else ''
             end,
           [Chg_Logistics_Rej_sp_ctno] = 
             case
               when Isnumeric(Substring(sp_rem, 5, 10)) = 1 and Len(RTrim(Substring(sp_rem, 5, 10))) = 10 and Isnull(D13.Code_Begin, '') like '[0-7RSC]' 
               then 
                   (select sp_ctno 
                      from fact_sslip 
                     where sp_slip_fg = Isnull(D13.Code_Begin, '') Collate Chinese_Taiwan_Stroke_CI_AS 
                       and sp_no = Substring(sp_rem, 5, 10) Collate Chinese_Taiwan_Stroke_CI_AS)
               else ''
             end,
           [Chg_Logistics_Rej_sp_ctname] = 
             case
               when Isnumeric(Substring(sp_rem, 5, 10)) = 1 and Len(RTrim(Substring(sp_rem, 5, 10))) = 10 and Isnull(D13.Code_Begin, '') like '[0-7RSC]'
               then
                    (select sp_ctname 
                       from fact_sslip 
                      where sp_slip_fg = Isnull(D13.Code_Begin, '') Collate Chinese_Taiwan_Stroke_CI_AS 
                        and sp_no = Substring(sp_rem, 5, 10)) Collate Chinese_Taiwan_Stroke_CI_AS
               else ''
             end,
           -- 2017/08/04 rickliu �s�W��ګȤᤧ�~�ȬO�_��¾
           ct_sale_leave,
/**********************************************************************************************************************************************************
  ���u�򥻸�ư�
 **********************************************************************************************************************************************************/
           [###pemploy###]='### pemploy ###',
           --2014/05/15 Rickliu -- �N��ڤW�����n���]�@�ֱa�J�A�p���K�i���ݭn�B�~��D�ɧ@ JOIN�A�i�H�[�t�d�߳t�סC
           sp_sales, -- �~�ȭ��s
           E_NAME as sp_sales_name, -- �~�ȦW��
           ct_class, -- �Ȥ�μt�����O
           sp_ctno, -- �Ȥ�μt�ӽs��
           ct_no8, -- �Ȥ�K�X�s��
           ct_sname8, -- �Ȥ�K�X�W��
           sp_ctname, -- �Ȥ�μt�ӦW��
           ct_sname, -- �Ȥ�μt��²��
           ct_fld3, 
           e_dept, -- �����s��
           Chg_sp_dp_Name,
           Chg_dp_name, -- �����W��
           ct_loc, -- �ϰ�s��
           Chg_loc_Name, -- �ϰ�W��
           Chg_ctno_CustKind_CustCity, -- �Ȥ�����|�����O�W��
           Chg_ctno_CustChain, -- �q���O�W��
           Chg_busine_Name, -- ��~�O�W��
           ct_sour, -- �Ȥ�ӷ��s��
           Chg_sour_Name, -- �Ȥ�ӷ��W��
           Chg_ctno_Port_Office, -- ���~�P|���q�O�W��
           ct_kind, -- �Ȥ����O
           Chg_ctclass, --�Ȥ����O�W��
           sp_mstno, -- �~�ȲէO�s��
           Chg_sp_mst_name, --�~�ȲէO�W��
           chg_leave, -- ��¾�_
/**********************************************************************************************************************************************************
  �Ȥ�򥻸�ư�
 **********************************************************************************************************************************************************/
           [###pcust###]='### pcust ###',
           -- 2015/02/26 Rickliu �W�[ �Ȥ� �P�f����, ���s���ĤE�X 1: �ʳf, 2: 3C, 3, �H��
           Chg_Cust_Sale_Class,
           Chg_Cust_Sale_Class_Name,
           Chg_Cust_Sale_Class_sName,
           D9.chg_ct_close,
           -- 2014/01/27 Rickliu�s�W ���q�Ȥ�
           D9.Chg_IS_Lan_Custom,
           Chg_BU_NO, --�`���q�s��
           rtrim(ct_fld3) as Chg_ct_fld3, --Rickliu 2014/1/24 �s�W �`���q
           
           Chg_Hunderd_Customer, --Rickliu 2014/3/17 �s�W �O�_���ʤj
           Chg_Hunderd_Customer_Name, --Rickliu 2014/1/24 �ʤj�Ȥ�W��
           -- 2014/10/02 Rickliu �s�W�����̫�@������X��
           Chg_CTNO_Last_Order_Flag = IsNull(D11.Flag, ''),

           -- 2014/10/02 Rickliu �s�W�`���q�̫�@������X��
           Chg_BUNO_Last_Order_Flag = IsNull(D12.Flag, ''),
/**********************************************************************************************************************************************************
  �f�~�򥻸�ư�
 **********************************************************************************************************************************************************/
           -- �ѩ�@�Ӹ�ƪ����x�s����W�L 8,060 �r���A�ҥH�����D���n���׶i��
           [###sstock###]='### sstock ###',
           [sk_bcode], --���X�N��
           [sk_kind], --���O
           [sk_ivkd], --�o�����O
           [sk_spec], --�W��
           [sk_unit], --�򥻳��
           [sk_color], --�C�� �� ���~�U�[���(�褸�~��)
           [sk_size], --�ؤo �� �s�~��f���(�褸�~��)
           [sk_use], --�γ~
           [sk_aux_q], --���U���ƶq
           [sk_aux_u], --���U���
           [sk_save], --�зǦ���
           [sk_bqty], --�g�٧�q
           [sk_fign], --�ϧ��ɦW
           [sk_pic], -- Web �覡�e�{�ϧ��ɦW, 2017/07/12 Rickliu
           [sk_tfg], --�ǿ�X��
           [st_cspes], --�f�~�W�� �� ���n
           [st_espes], --�f�~�W��(�^)
           [st_unit], --���
           [st_inerqty], --���]�˳��]�˼�
           [st_inerunt], --���]�˳��
           [st_outqty], --�~�]�˳���
           [st_outunt], --�~�]�˳��
           [st_appr], --�C�c���n
           [st_apprunt], --���n���(CBM��CFT)
           [st_nw], --�b��
           [st_gw], --��
           [st_gwunt], --���q���
           [st_itclass], --�f�~����
           [st_cccode], --CCC�X
           [st_sespes], --�~�W(�^)
           [st_20cyqty], --20�`�f�d�i�e�Ǽƶq
           [st_40cyqty], --40�`�f�d�i�e�Ǽƶq
           [st_45cyqty], --45�`�f�d�i�e�Ǽƶq
           [sk_fld1], --�����@ �� �f�~���e��
           [sk_fld2], --�����G �� �f�~�W��
           [sk_fld3], --�����T �� �e�q
           [sk_fld4], --�����| �� �~�P-ABT/M&M
           [sk_fld5], --������ �� �u�`/�`��
           [sk_fld6], --������ �� ���D
           [sk_fld11], --�����Q�@-���P�T�{��
           [sk_fld12], --�����Q�G-(�½s�� 2012�}�b�ϥ�)
           [sk_ikind], --�������O
           [s_supp], --������
           [s_locate], --�s��a�I
           [s_price1], --���즨��
           [s_price2], --����
           [s_price3], --�j�L��
           [s_price4], --���L��
           [s_price5], --�p�L��
           [s_price6], --�@���
           [s_aprice], --�ثe��������
           [s_m_ave], --�륭������
           [s_updat1], --�̪�i�f��
           [s_updat2], --�̪�P�f��
           [s_lprice1], --�̪�i��
           [s_lprice2], --�̪�P��
           [s_accno1], --�i�f ��إN��
           [sk_nowqty], --�ثe�s�q
           [st_proven], --�f�~�ӷ�
           [st_lenght], --��n�W�� ��
           [st_width], --��n�W�� �e
           [st_height], --��n�W�� ��
           [st_lwhunit], --��n�W����e�����
           [st_unw], --�f�~���b��
           [st_ugw], --�f�~����
           [st_uappr], --�f�~�����n
           [st_uarea], --�f�~��쭱�n
           [st_areaunt], --�f�~���n���
           [st_fign2], --���ɦW2
           [st_fign3], --���ɦW3
           [st_fign4], --���ɦW4
           [sk_poseq], --�D�����ӧǸ�
           [sk_abcode], --�]�˱��X�N��
           [s_apice2], --��  ��(��)
           [s_apice3], --�j�L��(��)
           [s_apice4], --���L��(��)
           [s_apice5], --�p�L��(��)
           [s_apice6], --�@���(��)
           [sk_whno], --�J�w�ܮw
           [Chg_skno_accno],
           [Chg_skno_accno_Name],
           [Chg_skno_BKind],
           [Chg_skno_Bkind_Name],
           [Chg_skno_BKind2],
           [Chg_skno_Bkind_Name2],
           [Chg_skno_SKind],
           [Chg_skno_SKind_Name],
           [Chg_kind_Name],
           [Chg_ivkd_Name],
           [Chg_StartYear],
           [Chg_StartMonth],
           [Chg_EndYear],
           [Chg_EndMonth],
           [Chg_supp_Name],
           [Chg_supp_SName],
           [Chg_locate_MArea],
           [Chg_locate_DArea],
           [Chg_locate_row],
           [Chg_locate_col],
           [Chg_updat1_Year],
           [Chg_updat1_Month],
           [Chg_updat2_Year],
           [Chg_updat2_Month],
           [Chg_New_Arrival_Date], -- 2015/03/06 Rickliu �s�W�s�~��f��
           -- 2014/07/02 Rickliu �s�W���ʷs�~
           --Chg_IS_New_Stock =
           --  Case
           --    when (D4.[sk_no] is not null) And ([Chg_sp_date_YM] >= D4.[Chg_New_Arrival_YM]) then 'Y'
           --    else 'N'
           --  end,
           -- 2017/08/01 Rickliu �׭q�s�~�w�q�A�אּ�H�@�~�����޶i�����~���s�~
           [Chg_IS_New_Stock],
           D4.[Chg_New_First_Qty],
           chg_new_arrival_ym,
           
           -- 2017/08/07 Rickliu �����ϥΰӫ~���P�C��
           -- Chg_Stock_NonSales, 
           -- 2017/08/07 Rickliu ��H�ӫ~�򥻸�ƨ��o�O�_���D�P
           [Chg_IS_Master_Stock],
           stock_kind_list,
/**********************************************************************************************************************************************************
  �w�s��ư�
 **********************************************************************************************************************************************************/
           [###sware###]='### sware ###',
           -- 2014/06/17 Rickliu�s�W AA ����w���s�q
           [Chg_WD_AA_sQty],
           -- 2014/06/17 Rickliu�s�W AA ����{���w�s��
           [Chg_WD_AA_first_Qty],           
           -- 2014/06/17 Rickliu�s�W AA ����w�s�t����
           [Chg_WD_AA_first_diff_Qty],
           -- 2014/06/07 Rickliu�s�W AA �����{���w�s��
           [Chg_WD_AA_last_Qty],
           -- 2014/06/17 Rickliu�s�W AA �����w�s�t����
           [Chg_WD_AA_last_diff_Qty],
             
           -- 2014/06/17 Rickliu�s�W AB ����w���s�q
           [Chg_WD_AB_sQty],
           -- 2014/06/17 Rickliu�s�W AB ����{���w�s��
           [Chg_WD_AB_first_Qty],           
           -- 2014/06/17 Rickliu�s�W AB ����w�s�t����
           [Chg_WD_AB_first_diff_Qty],
           -- 2014/06/07 Rickliu�s�W AB �����{���w�s��
           [Chg_WD_AB_last_Qty],
           -- 2014/06/17 Rickliu�s�W AB �����w�s�t����
           [Chg_WD_AB_last_diff_Qty],
           
           -- 2014/06/17 Rickliu�s�W AC ����w���s�q
           [Chg_WD_AC_sQty],
           -- 2014/06/17 Rickliu�s�W AC ����{���w�s��
           [Chg_WD_AC_first_Qty],           
           -- 2014/06/17 Rickliu�s�W AC ����w�s�t����
           [Chg_WD_AC_first_diff_Qty],
           -- 2014/06/07 Rickliu�s�W AC �����{���w�s��
           [Chg_WD_AC_last_Qty],
           -- 2014/06/17 Rickliu�s�W AC �����w�s�t����
           [Chg_WD_AC_last_diff_Qty],

           -- 2014/06/17 Rickliu�s�W AG ����w���s�q
           [Chg_WD_AG_sQty],
           -- 2014/06/17 Rickliu�s�W AG ����{���w�s��
           [Chg_WD_AG_first_Qty],           
           -- 2014/06/17 Rickliu�s�W AG ����w�s�t����
           [Chg_WD_AG_first_diff_Qty],
           -- 2014/06/07 Rickliu�s�W AG �����{���w�s��
           [Chg_WD_AG_last_Qty],
           -- 2014/06/17 Rickliu�s�W AG �����w�s�t����
           [Chg_WD_AG_last_diff_Qty],

           -- 2014/06/16 Rickliu �s�W���P�~
           [Chg_IS_Dead_Stock],
           [Chg_Dead_First_Qty],
           [Chg_Dead_First_Amt],
           [Chg_Dead_Stock_YM],
           
           -- 2015/02/27 Rickliu �s�W�����P��q������, ���s���ĤE�X 1: �ʳf, 2: 3C, 3: �H��, 4:�N�u(OEM)
           [Chg_Dept_Cust_Chain_No],
           [Chg_Dept_Cust_Chain_Name],

           sslpdt_update_datetime = getdate(),
           sslpdt_timestamp = m.timestamp_column
           into Fact_sslpdt 
      from SYNC_TA13.dbo.sslpdt M With(NoLock)
           Left join SYNC_TA13.dbo.sware D2 With(NoLock) On M.sd_whno=D2.wh_no 
           Left join Fact_sstock D4 With(NoLock) On M.sd_skno=D4.sk_no
           Left join Ori_Xls#Sys_Code D5 With(NoLock) On  D5.code_class ='7' and M.sd_class = D5.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS 
           Left join Ori_Xls#Sys_Code D6 With(NoLock) On  D6.code_class ='8' and M.sd_class = D6.Code_Begin Collate Chinese_Taiwan_Stroke_CI_AS and M.sd_slip_fg = D6.Code_End Collate Chinese_Taiwan_Stroke_CI_AS
           Left join Fact_sslip D9 With(NoLock) On sd_class = D9.sp_class and sd_slip_fg = D9.sp_slip_fg and sd_no = D9.sp_no 
           -- 2017/08/07 Rickliu ��H�ӫ~�򥻸�ƨ��o�O�_���D�P
           --left join Ori_XLS#Master_Stock D8 With(NoLock)
           --  On Year(D9.sp_pdate) = D8.Sale_Year and Month(D9.sp_pdate) = D8.Sale_month
           -- and M.sd_skno collate Chinese_Taiwan_Stroke_CI_AS = D8.sk_no collate Chinese_Taiwan_Stroke_CI_AS 
           -- 2017/08/07 Rickliu ��H�ӫ~�򥻸�ƨ��o�O�_���D�P
           --left join Ori_XLS#Master_Stock D10 With(NoLock) 
           --  On Year(m.sd_date) = D10.Sale_Year and Month(m.sd_date) = D10.Sale_month
           -- 2013/6/1 �ܧ�D�P���c�A���� MasterKey ���A�D�P�� SK_NAME �h�אּ�D�P�ӫ~�W��
           -- and M.sd_skno collate Chinese_Taiwan_Stroke_CI_AS = D10.Masterkey collate Chinese_Taiwan_Stroke_CI_AS
           -- and M.sd_skno collate Chinese_Taiwan_Stroke_CI_AS = D10.sk_no collate Chinese_Taiwan_Stroke_CI_AS
            -- 2014/10/02 Rickliu �W�[�̫�����@������X��
           left join CTE_Q1 D11
            On m.sd_class = d11.sd_class
           and m.sd_slip_fg = d11.sd_slip_fg
           and m.sd_ctno = d11.sd_ctno
           and m.sd_no = d11.Max_SP_NO
           and m.sd_skno = d11.sd_skno
           and m.sd_price = d11.sd_price -- Rickliu 2017/06/08 CTE_Q1 �W�[�P�_�L�o�ث~���O�B���B�p��1 �]���ث~

           -- 2014/10/02 Rickliu �W�[�̫��`���q�@������X��
           left join CTE_Q2 D12
            On m.sd_class = d12.sd_class
           and m.sd_slip_fg = d12.sd_slip_fg
           --2016/09/09 Rickliu �]��´�X�s�A�W�[�����q���A�]���N�쥻 2~5 �אּ 2~6
           --and Substring(m.sd_ctno, 1, 5) = d12.BU_NO
           and Substring(m.sd_ctno, 1, 6) = d12.BU_NO
           and m.sd_no = d12.Max_SP_NO
           and m.sd_skno = d12.sd_skno
           and m.sd_price = d12.sd_price -- Rickliu 2017/06/08 CTE_Q1 �W�[�P�_�L�o�ث~���O�B���B�p��1 �]���ث~
           ---
           left join Ori_Xls#Sys_Code D13 With(NoLock) 
             On M.sd_date >= '2015/03/01' 
            and M.sd_slip_fg = 'B'
            and D13.code_class ='8' 
            and D13.Code_Name = RTrim(Substring(Sp_rem, 1, 3)) Collate Chinese_Taiwan_Stroke_CI_AS
    where 1=1
      -- 2017/08/08 Rickliu �q 2013 �~�_��A�O�d�񤭦~���
      and m.sd_date > '2013/01/01'
      and m.sd_date >= Convert(Varchar(5), year(DateAdd(year, -5, getdate())))+ '/01/01' 
    
  /*==============================================================*/
  /* Index: sslpdt_Timestamp                                      */
  /*==============================================================*/
  --set @Msg = '�إ߯��� [Fact_sslpdt.sslpdt_timestamp]'
  --Print @Msg
  --if exists (select * from sys.indexes where object_id = object_id('[dbo].[fact_sslpdt]') and name = 'pk_sslpdt')
  --   alter table [dbo].[pk_sslpdt] drop constraint [pk_sslpdt]

  --alter table [dbo].[fact_sslpdt] add  constraint [pk_sslpdt] primary key nonclustered ([sslpdt_timestamp] asc) with 
  --(pad_index  = off, statistics_norecompute  = off, sort_in_tempdb = off, ignore_dup_key = off, online = off, allow_row_locks  = on, allow_page_locks  = on) on [primary]


  /*==============================================================*/
  /* Index: sdno                                                  */
  /*==============================================================*/
  Set @Msg = '�إ߯��� [fact_sslpdt.sdno]'
  Print @Msg
  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'sdno' and indid > 0 and indid < 255) 
     drop index dbo.fact_sslpdt.sdno
  create clustered index sdno on dbo.fact_sslpdt (sd_no, sd_slip_fg) with fillfactor= 30  on "PRIMARY"
  /*==============================================================*/
  /* Index: class                                                 */
  /*==============================================================*/
  Set @Msg = '�إ߯��� [fact_sslpdt.class]'
  Print @Msg
  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'class' and indid > 0 and indid < 255) 
     drop index dbo.fact_sslpdt.class
  create index class on dbo.fact_sslpdt (sd_class) with fillfactor= 30 on "PRIMARY"
  /*==============================================================*/
  /* Index: CSNO                                                  */
  /*==============================================================*/
  Set @Msg = '�إ߯��� [fact_sslpdt (sd_csno)...'
  Print @Msg
  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'CSNO' and indid > 0 and indid < 255) 
     drop index dbo.fact_sslpdt.CSNO
  create index CSNO on dbo.fact_sslpdt (sd_csno) with fillfactor= 30 on "PRIMARY"
  /*==============================================================*/
  /* Index: ctno                                                  */
  /*==============================================================*/
  Set @Msg = '�إ߯��� [fact_sslpdt (sd_ctno, sd_class)...'
  Print @Msg
  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'ctno' and indid > 0 and indid < 255) 
     drop index dbo.fact_sslpdt.ctno
  create index ctno on dbo.fact_sslpdt (sd_ctno, sd_class) with fillfactor= 30 on "PRIMARY"
  /*==============================================================*/
  /* Index: ordno                                                 */
  /*==============================================================*/
  Set @Msg = '�إ߯��� [fact_sslpdt.ordno]'
  Print @Msg
  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'ordno' and indid > 0 and indid < 255) 
     drop index dbo.fact_sslpdt.ordno
  create index ordno on dbo.fact_sslpdt (sd_ordno, sd_class) with fillfactor= 30 on "PRIMARY"
  /*==============================================================*/
  /* Index: sdda                                                  */
  /*==============================================================*/
  Set @Msg = '�إ߯��� [fact_sslpdt.sdda]'
  Print @Msg
  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'sdda' and indid > 0 and indid < 255) 
     drop index dbo.fact_sslpdt.sdda
  create index sdda on dbo.fact_sslpdt (sd_date, sd_class) with fillfactor= 30 on "PRIMARY"
  /*==============================================================*/
  /* Index: skda                                                  */
  /*==============================================================*/
  Set @Msg = '�إ߯��� [fact_sslpdt.skda]'
  Print @Msg
  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'skda' and indid > 0 and indid < 255) 
     drop index dbo.fact_sslpdt.skda
  create index skda on dbo.fact_sslpdt (sd_skno, sd_date) with fillfactor= 30 on "PRIMARY"
  /*==============================================================*/
  /* Index: skno                                                  */
  /*==============================================================*/
  Set @Msg = '�إ߯��� [fact_sslpdt.skno]'
  Print @Msg
  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'skno' and indid > 0 and indid < 255) 
     drop index dbo.fact_sslpdt.skno
  create index skno on dbo.fact_sslpdt (sd_skno, sd_ctno, sd_date, sd_class) with fillfactor= 30 on "PRIMARY"
  /*==============================================================*/
  /* Index: slip_fg                                               */
  /*==============================================================*/
  Set @Msg = '�إ߯��� [Fact_sslpdt.slip_fg]'
  Print @Msg
  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'slip_fg' and indid > 0 and indid < 255) 
     drop index dbo.fact_sslpdt.slip_fg
  create index slip_fg on dbo.fact_sslpdt (sd_slip_fg) with fillfactor= 30 on "PRIMARY"
  /*==============================================================*/
  /* Index: IX_sslpdt_N1                                          */
  /*==============================================================*/
  --  Print 'Create index [IX_sslpdt_N1] on [dbo].[Fact_sslpdt] ([Chg_skno_BKind], [sd_slip_fg], [Chg_sp_pdate_Year])...'
  --  if exists (select 1 from sysindexes where id = object_id('dbo.fact_sslpdt') and name  = 'IX_sslpdt_N1' and indid > 0 and indid < 255) 
  --     drop index dbo.fact_sslpdt.slip_fg
  --  create index [IX_sslpdt_N1] on [dbo].[Fact_sslpdt] ([Chg_skno_BKind], [sd_slip_fg], [Chg_sp_pdate_Year])
  --  INCLUDE ([Chg_sales_Name],[Chg_sd_Pay_Sale],[Chg_sp_pdate_Month],[sp_sales])


/*********************************************************************************************************************************************************
  2017/08/15 Ricliu �H�U����O�M�� �D�P�B�s�~�B���P���ʤj�ΫD�ʤj �Q�f�v ����ҥ�
********************************************************************************************************************************************************/
  IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fact_sslpdt_Near_Year]') AND type in (N'U'))
  begin
     Set @Msg = '�R����ƪ� [fact_sslpdt_Near_Year]'
     set @strSQL= 'DROP TABLE [dbo].[fact_sslpdt_Near_Year]'

     Exec uSP_Sys_Exec_SQL @Proc, @Msg, @strSQL
  end

  select Chg_sp_pdate_YM,
         Chg_bu_no, ct_fld3 as Chg_bu_Name,
         ct_sales, chg_ct_sales_name,
         ct_no8, ct_sname8+
         case
           when chg_ct_close  = 'Y' then '[��]'
           else ''
         end as ct_sname8,
         chg_ct_close,
         sd_skno, sd_name, stock_kind_list,
         chg_is_master_stock,
         chg_is_New_Stock,
         chg_is_Dead_Stock,
         Chg_Hunderd_Customer,
         Sum(Chg_sd_qty) as Chg_sd_qty,
         Sum(Chg_sd_stot) as Chg_sd_stot
         into fact_sslpdt_Near_Year
    from fact_sslpdt with(NoLock)
   where 1=1
     and sd_class in ('1', '8')
     and Chg_sp_pdate_YM >='2013/01'
     and substring(ct_no8, 1, 2) Not in ('IT', 'IZ', 'ZZ')
     and substring(sd_skno, 1, 1) = 'A'
     and Chg_sp_pdate_YM >= Convert(Varchar(7), DateAdd(mm, -12,getdate()), 111)
     and ct_fld3 <> ''
   group by Chg_sp_pdate_YM, Chg_bu_no, ct_fld3,
            ct_sales, chg_ct_sales_name,
            ct_no8, ct_sname8, chg_ct_close,
            sd_skno, sd_name, stock_kind_list,
            chg_is_master_stock,
            chg_is_New_Stock,
            chg_is_Dead_Stock,
             Chg_Hunderd_Customer
   order by Chg_sp_pdate_YM, Chg_bu_no, ct_fld3,
            ct_sales, chg_ct_sales_name,
            ct_no8, ct_sname8, chg_ct_close,
            sd_skno, sd_name, stock_kind_list,
            chg_is_master_stock,
            chg_is_New_Stock,
            chg_is_Dead_Stock,
            Chg_Hunderd_Customer
   
  end try
  begin catch
    set @Cnt = -1
    set @Msg = @Proc+'...(���~�T��:'+ERROR_MESSAGE()+', '+@Msg+')...(���~�C:'+Convert(Varchar(10), ERROR_LINE())+')'

    Exec uSP_Sys_Write_Log @Proc, @Msg, @Msg, @Cnt
  end catch
  Return(@Cnt)
end
GO
