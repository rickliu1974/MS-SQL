USE [DW]
GO
/****** Object:  Table [dbo].[Fact_sslip]    Script Date: 07/24/2017 14:43:43 ******/
DROP TABLE [dbo].[Fact_sslip]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Fact_sslip](
	[sp_class] [char](1) NOT NULL,
	[sp_slip_fg] [char](1) NOT NULL,
	[sp_date] [datetime] NULL,
	[sp_pdate] [datetime] NULL,
	[sp_no] [nvarchar](10) NULL,
	[sp_ordno] [nvarchar](16) NULL,
	[sp_ctno] [nvarchar](10) NULL,
	[sp_ctname] [nvarchar](60) NULL,
	[Ori_sp_ctname] [nvarchar](60) NULL,
	[sp_ctadd2] [varchar](100) NULL,
	[sp_sales] [nvarchar](10) NULL,
	[Chg_sales_Name] [nvarchar](10) NULL,
	[sp_pyno] [nchar](10) NULL,
	[sp_dpno] [nvarchar](8) NULL,
	[sp_maker] [nvarchar](10) NULL,
	[sp_conv] [nchar](10) NULL,
	[sp_tot] [decimal](20, 6) NULL,
	[sp_tax] [decimal](20, 6) NULL,
	[sp_dis] [decimal](20, 6) NULL,
	[sp_pay] [decimal](20, 6) NULL,
	[sp_pamt] [decimal](20, 6) NULL,
	[sp_cash] [decimal](20, 6) NULL,
	[sp_mamt] [decimal](20, 6) NULL,
	[sp_pay_kd] [char](1) NULL,
	[sp_rate_nm] [nvarchar](5) NULL,
	[sp_rate] [decimal](20, 6) NULL,
	[sp_ono] [nchar](10) NULL,
	[sp_tno] [nchar](11) NULL,
	[sp_wkno] [nchar](10) NULL,
	[sp_tailno] [nchar](10) NULL,
	[sp_tal_rec] [int] NULL,
	[sp_ave_p] [decimal](20, 6) NULL,
	[sp_acapno] [nchar](10) NULL,
	[sp_acspno] [nchar](10) NULL,
	[sp_invoice] [nchar](10) NULL,
	[sp_itot] [decimal](20, 6) NULL,
	[sp_inv_kd] [decimal](20, 6) NULL,
	[sp_tax_kd] [decimal](20, 6) NULL,
	[sp_i_date] [datetime] NULL,
	[sp_invtype] [decimal](20, 6) NULL,
	[sp_dkind] [decimal](20, 6) NULL,
	[sp_wcd] [char](1) NULL,
	[sp_rem] [varchar](100) NULL,
	[sp_npack] [int] NULL,
	[sp_nokind] [bit] NOT NULL,
	[sp_tfg] [bit] NOT NULL,
	[sp_postfg] [bit] NOT NULL,
	[sp_ntnpay] [decimal](20, 6) NULL,
	[sp_acspseq] [decimal](20, 6) NULL,
	[sp_addtax] [bit] NOT NULL,
	[sp_trdspno] [nchar](14) NULL,
	[sp_mave_p] [decimal](20, 6) NULL,
	[sp_caseno] [nchar](20) NULL,
	[sp_mafkd] [int] NULL,
	[sp_mafno] [nchar](16) NULL,
	[sp_adjkd] [int] NULL,
	[sp_surefg] [bit] NOT NULL,
	[sp_suredt] [datetime] NULL,
	[sp_sureman] [nchar](10) NULL,
	[sp_mstno] [nchar](10) NULL,
	[sp_bimtype] [int] NULL,
	[Chg_sp_class] [nvarchar](255) NULL,
	[Chg_sp_slip_fg] [nvarchar](255) NULL,
	[Chg_sp_slip_name] [nvarchar](12) NULL,
	[Chg_sp_dp_name] [nvarchar](20) NULL,
	[Chg_sp_date_Year] [int] NULL,
	[Chg_sp_date_Quarter] [int] NULL,
	[Chg_sp_date_Month] [int] NULL,
	[Chg_sp_date_YearWeek] [int] NULL,
	[Chg_sp_date_YM] [varchar](10) NULL,
	[Chg_sp_date_MD] [varchar](10) NULL,
	[Chg_sp_date_Weekno] [int] NULL,
	[Chg_sp_date_Weekname] [varchar](4) NULL,
	[Chg_sp_date_Week_Range] [varchar](30) NULL,
	[Chg_sp_pdate_Year] [int] NULL,
	[Chg_sp_pdate_Quarter] [int] NULL,
	[Chg_sp_pdate_Month] [int] NULL,
	[Chg_sp_pdate_YearWeek] [int] NULL,
	[Chg_sp_pdate_YM] [varchar](10) NULL,
	[Chg_sp_pdate_MD] [varchar](10) NULL,
	[Chg_sp_pdate_Weekno] [int] NULL,
	[Chg_sp_pdate_Weekname] [varchar](4) NULL,
	[Chg_sp_pdate_Week_Range] [varchar](30) NULL,
	[Chg_sp_tax_kd] [nvarchar](255) NULL,
	[Chg_sp_inv_name] [nvarchar](255) NULL,
	[Chg_sp_invtype_name] [nvarchar](255) NULL,
	[Chg_sp_mst_name] [nchar](80) NOT NULL,
	[Chg_sp_stot] [decimal](38, 9) NULL,
	[Chg_SP_TAX] [decimal](38, 9) NULL,
	[Chg_sp_stot_tax] [decimal](38, 9) NULL,
	[Chg_SP_PAY] [decimal](38, 9) NULL,
	[Chg_SP_DIS] [decimal](38, 9) NULL,
	[Chg_SP_PAMT] [decimal](38, 9) NULL,
	[Chg_SP_MAMT] [decimal](38, 9) NULL,
	[Chg_sp_ave_p] [decimal](20, 6) NULL,
	[Chg_sp_mave_p] [decimal](20, 6) NULL,
	[Chg_sp_ntnpay_plus] [numeric](38, 7) NULL,
	[Chg_sp_sum_tot] [decimal](38, 9) NULL,
	[Chg_sp_dis2] [decimal](38, 6) NULL,
	[Chg_sp_dis_tax2] [decimal](38, 6) NULL,
	[Chg_sp_dis_tot2] [decimal](38, 6) NULL,
	[Chg_sp_sum_tot2] [decimal](38, 6) NULL,
	[Chg_SP_Not_Recv_Tot] [decimal](38, 9) NULL,
	[Chg_Area_Target] [float] NOT NULL,
	[Chg_Dept_Cust_Chain_No] [nvarchar](10) NULL,
	[Chg_Dept_Cust_Chain_Name] [nvarchar](30) NULL,
	[Chg_SP_Pay_Rate] [decimal](38, 6) NULL,
	[###pcust###] [varchar](13) NOT NULL,
	[ct_class] [char](1) NULL,
	[ct_no] [nvarchar](10) NULL,
	[ct_name] [nvarchar](60) NULL,
	[ct_sname] [nvarchar](12) NULL,
	[ct_no8] [nvarchar](8) NULL,
	[ct_sname8] [nvarchar](121) NULL,
	[ct_addr1] [varchar](255) NULL,
	[ct_addr2] [varchar](255) NULL,
	[ct_addr3] [varchar](255) NULL,
	[ct_tel] [nvarchar](35) NULL,
	[ct_fax] [nvarchar](20) NULL,
	[ct_unino] [nvarchar](10) NULL,
	[ct_presid] [nvarchar](40) NULL,
	[ct_contact] [nvarchar](40) NULL,
	[ct_payfg] [char](1) NULL,
	[ct_sales] [nvarchar](10) NULL,
	[ct_p_limit] [decimal](20, 6) NULL,
	[ct_b_limit] [decimal](20, 6) NULL,
	[ct_bkno] [nvarchar](20) NULL,
	[ct_bknm] [nvarchar](50) NULL,
	[ct_rem] [varchar](255) NULL,
	[ct_dept] [nvarchar](8) NULL,
	[ct_payrate] [int] NULL,
	[ct_ivtitle] [nvarchar](60) NULL,
	[ct_porter] [nvarchar](10) NULL,
	[ct_credit] [nvarchar](10) NULL,
	[ct_pmode] [char](1) NULL,
	[ct_pdate] [int] NULL,
	[ct_prenpay] [decimal](20, 6) NULL,
	[ct_prepay] [decimal](20, 6) NULL,
	[ct_last_dt] [datetime] NULL,
	[ct_flg] [bit] NULL,
	[ct_t_fg] [bit] NULL,
	[ct_grade] [char](2) NULL,
	[ct_area] [char](3) NULL,
	[ct_curt_id] [nvarchar](8) NULL,
	[ct_cont_sp] [nvarchar](20) NULL,
	[ct_pay] [nchar](8) NULL,
	[ct_regist] [nvarchar](14) NULL,
	[ct_worker] [decimal](20, 6) NULL,
	[ct_capital] [nchar](30) NULL,
	[ct_skpay] [decimal](20, 6) NULL,
	[ct_sknpay] [decimal](20, 6) NULL,
	[ct_accno2] [nchar](8) NULL,
	[ct_chkno2] [nchar](8) NULL,
	[ct_cdate] [datetime] NULL,
	[ct_payer] [nchar](10) NULL,
	[ct_advance] [decimal](20, 6) NULL,
	[ct_debt] [decimal](20, 6) NULL,
	[ct_abroad] [decimal](20, 6) NULL,
	[ct_fld1] [nvarchar](60) NULL,
	[ct_fld2] [nvarchar](60) NULL,
	[ct_fld3] [nvarchar](60) NULL,
	[ct_fld4] [nvarchar](60) NULL,
	[ct_fld5] [nvarchar](60) NULL,
	[ct_fld6] [nvarchar](60) NULL,
	[ct_fld7] [nvarchar](60) NULL,
	[ct_fld8] [nvarchar](60) NULL,
	[ct_fld9] [nvarchar](60) NULL,
	[ct_fld10] [nvarchar](60) NULL,
	[ct_fld11] [nvarchar](60) NULL,
	[ct_fld12] [nvarchar](60) NULL,
	[ct_sea] [nchar](20) NULL,
	[ct_invofg] [char](1) NULL,
	[ct_udec] [decimal](20, 6) NULL,
	[ct_tdec] [decimal](20, 6) NULL,
	[ct_busine] [nchar](20) NULL,
	[ct_banpay] [nvarchar](20) NULL,
	[ct_loc] [nchar](10) NULL,
	[ct_sour] [nchar](10) NULL,
	[ct_kind] [nchar](10) NULL,
	[ct_tkday] [int] NULL,
	[Chg_ctclass] [varchar](4) NULL,
	[Chg_BU_NO] [nvarchar](6) NULL,
	[Chg_Is_BU] [varchar](1) NULL,
	[Chg_ctno_Port_Office] [nvarchar](255) NULL,
	[Chg_ctno_CustKind_CustCity] [nvarchar](255) NULL,
	[Chg_ctno_CustChain] [nvarchar](255) NULL,
	[Chg_ct_dept_Name] [nvarchar](20) NULL,
	[Chg_ct_sales_Name] [nvarchar](10) NULL,
	[Chg_credit_Name] [nvarchar](12) NULL,
	[Chg_busine_Name] [nvarchar](40) NULL,
	[Chg_loc_Name] [nvarchar](40) NULL,
	[Chg_sour_Name] [nvarchar](40) NULL,
	[Chg_Customer_kind_Name] [nvarchar](40) NULL,
	[Chg_payfg_Name] [nvarchar](255) NULL,
	[Chg_porter_Name] [nvarchar](30) NULL,
	[Chg_pmode_Name] [nvarchar](255) NULL,
	[Chg_fld1_Year] [nvarchar](4) NULL,
	[Chg_fld1_Month] [nvarchar](2) NULL,
	[Chg_fld2_Year] [nvarchar](4) NULL,
	[Chg_fld2_Month] [nvarchar](2) NULL,
	[Chg_Hunderd_Customer] [varchar](1) NULL,
	[Chg_Hunderd_Customer_Name] [nvarchar](255) NULL,
	[Chg_IS_Lan_Custom] [varchar](1) NULL,
	[Chg_ct_close] [varchar](1) NULL,
	[Chg_rate_date] [datetime] NULL,
	[Chg_rate] [decimal](20, 6) NULL,
	[Chg_Cust_Sale_Class] [nvarchar](1) NULL,
	[Chg_Cust_Sale_Class_Name] [varchar](9) NULL,
	[Chg_Cust_Sale_Class_sName] [nvarchar](4000) NULL,
	[Chg_Cust_Dis_Mapping] [varchar](1) NULL,
	[pcust_update_datetime] [datetime] NULL,
	[pcust_timestamp] [binary](8) NULL,
	[###pmploy###] [varchar](15) NOT NULL,
	[e_no] [nvarchar](10) NULL,
	[e_name] [nvarchar](10) NULL,
	[E_SEX] [decimal](20, 6) NULL,
	[E_MARR] [nchar](10) NULL,
	[E_BTYPE] [decimal](20, 6) NULL,
	[E_PREV] [nchar](16) NULL,
	[E_ID] [nchar](20) NULL,
	[E_BIRTH] [datetime] NULL,
	[E_DEPT] [nchar](8) NULL,
	[E_DUTY] [nchar](20) NULL,
	[E_RDATE] [datetime] NULL,
	[E_LDATE] [datetime] NULL,
	[e_bkno] [nvarchar](20) NULL,
	[e_contact] [nvarchar](10) NULL,
	[e_tel] [nvarchar](50) NULL,
	[E_ADDR] [varchar](255) NULL,
	[E_BADDR] [varchar](255) NULL,
	[E_CRED] [varchar](255) NULL,
	[E_VITA] [varchar](255) NULL,
	[E_REM] [varchar](255) NULL,
	[e_figname] [nvarchar](60) NULL,
	[e_spday] [varchar](41) NULL,
	[e_fld1] [nvarchar](60) NULL,
	[e_fld2] [nvarchar](60) NULL,
	[e_fld3] [nvarchar](60) NULL,
	[e_fld4] [nvarchar](60) NULL,
	[e_fld5] [nvarchar](60) NULL,
	[e_fld6] [nvarchar](60) NULL,
	[e_fld7] [nvarchar](60) NULL,
	[e_fld8] [nvarchar](60) NULL,
	[e_fld9] [nvarchar](60) NULL,
	[e_fld10] [nvarchar](60) NULL,
	[e_fld11] [nvarchar](60) NULL,
	[e_fld12] [nvarchar](60) NULL,
	[E_TFG] [bit] NULL,
	[E_INBANK] [nchar](20) NULL,
	[e_mstno] [nvarchar](10) NULL,
	[EM_BIMID] [varchar](20) NULL,
	[Chg_E_BTYPE] [varchar](2) NULL,
	[Chg_e_sex] [varchar](2) NULL,
	[Chg_E_BIRTH] [varchar](6) NULL,
	[Chg_e_rdate_YY] [int] NULL,
	[Chg_e_rdate_MM] [int] NULL,
	[Chg_e_ldate_YY] [int] NULL,
	[Chg_e_ldate_MM] [int] NULL,
	[Chg_emp_day] [int] NULL,
	[Chg_emp_sumYM] [numeric](26, 1) NULL,
	[Chg_leave] [varchar](1) NULL,
	[Chg_dp_name] [nvarchar](20) NULL,
	[Chg_dp_ename] [nvarchar](40) NULL,
	[Chg_dp_whno] [nvarchar](10) NULL,
	[Chg_dp_rem] [varchar](255) NULL,
	[Chg_mst_name] [nvarchar](80) NULL,
	[Chg_Year_Tot_SaleAmt] [decimal](38, 6) NULL,
	[Chg_Master_Dept] [nvarchar](2) NULL,
	[Chg_Duty_Level] [int] NULL,
	[Chg_Dept_Level] [nvarchar](8) NULL,
	[stock_amt_Level] [int] NULL,
	[Result_e_no] [nvarchar](10) NULL,
	[Result_dept_level] [varchar](5) NULL,
	[Result_Duty_Level] [varchar](5) NULL,
	[Result_e_mstno] [nvarchar](10) NULL,
	[update_datetime] [datetime] NULL,
	[pemploy_timestamp] [binary](8) NULL,
	[sslip_update_datetime] [datetime] NOT NULL,
	[sslip_timestamp] [binary](8) NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO