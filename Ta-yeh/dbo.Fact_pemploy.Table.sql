USE [DW]
GO
/****** Object:  Table [dbo].[Fact_pemploy]    Script Date: 07/24/2017 14:43:43 ******/
DROP TABLE [dbo].[Fact_pemploy]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Fact_pemploy](
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
	[E_TFG] [bit] NOT NULL,
	[E_INBANK] [nchar](20) NULL,
	[e_mstno] [nvarchar](10) NULL,
	[EM_BIMID] [varchar](20) NULL,
	[Chg_E_BTYPE] [varchar](2) NOT NULL,
	[Chg_e_sex] [varchar](2) NOT NULL,
	[Chg_E_BIRTH] [varchar](6) NOT NULL,
	[Chg_e_rdate_YY] [int] NULL,
	[Chg_e_rdate_MM] [int] NULL,
	[Chg_e_ldate_YY] [int] NULL,
	[Chg_e_ldate_MM] [int] NULL,
	[Chg_emp_day] [int] NULL,
	[Chg_emp_sumYM] [numeric](26, 1) NULL,
	[Chg_leave] [varchar](1) NOT NULL,
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
	[update_datetime] [datetime] NOT NULL,
	[pemploy_timestamp] [binary](8) NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
