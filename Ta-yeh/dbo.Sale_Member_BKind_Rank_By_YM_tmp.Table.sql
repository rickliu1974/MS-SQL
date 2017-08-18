USE [DW]
GO
/****** Object:  Table [dbo].[Sale_Member_BKind_Rank_By_YM_tmp]    Script Date: 07/24/2017 14:43:52 ******/
DROP TABLE [dbo].[Sale_Member_BKind_Rank_By_YM_tmp]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Sale_Member_BKind_Rank_By_YM_tmp](
	[Rank] [bigint] NULL,
	[year] [int] NULL,
	[month] [int] NULL,
	[kind_name_AA] [varchar](4) NOT NULL,
	[emp_no_AA] [varchar](20) NULL,
	[emp_name_AA] [varchar](50) NULL,
	[sale_amt_AA] [float] NULL,
	[target_amt_AA] [float] NULL,
	[rate_AA] [float] NULL,
	[kind_name_AB] [varchar](2) NOT NULL,
	[emp_no_AB] [varchar](20) NULL,
	[emp_name_AB] [varchar](50) NULL,
	[sale_amt_AB] [float] NULL,
	[target_amt_AB] [float] NULL,
	[rate_AB] [float] NULL,
	[kind_name_AC] [varchar](4) NOT NULL,
	[emp_no_AC] [varchar](20) NULL,
	[emp_name_AC] [varchar](50) NULL,
	[sale_amt_AC] [float] NULL,
	[target_amt_AC] [float] NULL,
	[rate_AC] [float] NULL,
	[kind_name_AD] [varchar](4) NOT NULL,
	[emp_no_AD] [varchar](20) NULL,
	[emp_name_AD] [varchar](50) NULL,
	[sale_amt_AD] [float] NULL,
	[target_amt_AD] [float] NULL,
	[rate_AD] [float] NULL,
	[kind_name_AE] [varchar](4) NOT NULL,
	[emp_no_AE] [varchar](20) NULL,
	[emp_name_AE] [varchar](50) NULL,
	[sale_amt_AE] [float] NULL,
	[target_amt_AE] [float] NULL,
	[rate_AE] [float] NULL,
	[kind_name_OT] [varchar](4) NOT NULL,
	[emp_no_OT] [varchar](20) NULL,
	[emp_name_OT] [varchar](50) NULL,
	[sale_amt_OT] [float] NULL,
	[target_amt_OT] [float] NULL,
	[rate_OT] [float] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
