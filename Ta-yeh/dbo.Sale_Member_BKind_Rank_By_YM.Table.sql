USE [DW]
GO
/****** Object:  Table [dbo].[Sale_Member_BKind_Rank_By_YM]    Script Date: 07/24/2017 14:43:52 ******/
DROP TABLE [dbo].[Sale_Member_BKind_Rank_By_YM]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Sale_Member_BKind_Rank_By_YM](
	[Rank] [bigint] NULL,
	[year] [int] NULL,
	[month] [int] NULL,
	[kind_name_AA] [varchar](4) NOT NULL,
	[emp_no_AA] [varchar](20) NOT NULL,
	[emp_name_AA] [varchar](50) NOT NULL,
	[sale_amt_AA] [float] NOT NULL,
	[target_amt_AA] [float] NOT NULL,
	[rate_AA] [float] NOT NULL,
	[YM_Sale_amt_AA] [float] NOT NULL,
	[YM_Target_AA] [float] NOT NULL,
	[YM_rate_AA] [float] NOT NULL,
	[kind_name_AB] [varchar](2) NOT NULL,
	[emp_no_AB] [varchar](20) NOT NULL,
	[emp_name_AB] [varchar](50) NOT NULL,
	[sale_amt_AB] [float] NOT NULL,
	[target_amt_AB] [float] NOT NULL,
	[rate_AB] [float] NOT NULL,
	[YM_Sale_amt_AB] [float] NOT NULL,
	[YM_Target_AB] [float] NOT NULL,
	[YM_rate_AB] [float] NOT NULL,
	[kind_name_AC] [varchar](4) NOT NULL,
	[emp_no_AC] [varchar](20) NOT NULL,
	[emp_name_AC] [varchar](50) NOT NULL,
	[sale_amt_AC] [float] NOT NULL,
	[target_amt_AC] [float] NOT NULL,
	[rate_AC] [float] NOT NULL,
	[YM_Sale_amt_AC] [float] NOT NULL,
	[YM_Target_AC] [float] NOT NULL,
	[YM_rate_AC] [float] NOT NULL,
	[kind_name_AD] [varchar](4) NOT NULL,
	[emp_no_AD] [varchar](20) NOT NULL,
	[emp_name_AD] [varchar](50) NOT NULL,
	[sale_amt_AD] [float] NOT NULL,
	[target_amt_AD] [float] NOT NULL,
	[rate_AD] [float] NOT NULL,
	[YM_Sale_amt_AD] [float] NOT NULL,
	[YM_Target_AD] [float] NOT NULL,
	[YM_rate_AD] [float] NOT NULL,
	[kind_name_AE] [varchar](4) NOT NULL,
	[emp_no_AE] [varchar](20) NOT NULL,
	[emp_name_AE] [varchar](50) NOT NULL,
	[sale_amt_AE] [float] NOT NULL,
	[target_amt_AE] [float] NOT NULL,
	[rate_AE] [float] NOT NULL,
	[YM_Sale_amt_AE] [float] NOT NULL,
	[YM_Target_AE] [float] NOT NULL,
	[YM_rate_AE] [float] NOT NULL,
	[kind_name_OT] [varchar](4) NOT NULL,
	[emp_no_OT] [varchar](20) NOT NULL,
	[emp_name_OT] [varchar](50) NOT NULL,
	[sale_amt_OT] [float] NOT NULL,
	[target_amt_OT] [float] NOT NULL,
	[rate_OT] [float] NOT NULL,
	[YM_Sale_amt_OT] [float] NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
