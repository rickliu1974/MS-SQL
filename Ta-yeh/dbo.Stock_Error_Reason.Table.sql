USE [DW]
GO
/****** Object:  Table [dbo].[Stock_Error_Reason]    Script Date: 07/24/2017 14:43:53 ******/
DROP TABLE [dbo].[Stock_Error_Reason]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Stock_Error_Reason](
	[year] [float] NULL,
	[month] [float] NULL,
	[year_month] [varchar](8) NULL,
	[sk_no] [nvarchar](255) NULL,
	[sk_name] [nvarchar](60) NULL,
	[error_cnt] [float] NULL,
	[sk_cnt] [decimal](38, 6) NULL,
	[sk_sale_cnt] [decimal](38, 6) NULL,
	[Chg_skno_BKind] [nvarchar](2) NULL,
	[Chg_skno_Bkind_Name] [nvarchar](255) NULL,
	[Chg_skno_SKind] [nvarchar](4) NULL,
	[Chg_skno_Skind_Name] [nvarchar](255) NULL,
	[Chg_Wd_AA_first_Qty] [decimal](20, 6) NULL,
	[Chg_Wd_AA_last_Qty] [decimal](32, 6) NULL,
	[error_reason] [nvarchar](max) NULL,
	[From_Cust] [nvarchar](max) NULL,
	[error_rate] [float] NULL,
	[error_sale_rate] [float] NULL,
	[TOT_error_cnt] [float] NULL,
	[TOT_sk_cnt] [decimal](38, 6) NULL,
	[TOT_error_rate] [float] NULL,
	[TOT_sk_sale_cnt] [decimal](38, 6) NULL,
	[TOT_error_sale_rate] [float] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
