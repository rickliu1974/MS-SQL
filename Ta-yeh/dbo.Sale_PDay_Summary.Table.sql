USE [DW]
GO
/****** Object:  Table [dbo].[Sale_PDay_Summary]    Script Date: 07/24/2017 14:43:52 ******/
DROP TABLE [dbo].[Sale_PDay_Summary]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Sale_PDay_Summary](
	[sp_ctno] [nvarchar](10) NULL,
	[ct_sname] [nvarchar](12) NULL,
	[E_NO] [nvarchar](10) NULL,
	[e_name] [nvarchar](10) NULL,
	[e_dept] [nvarchar](8) NULL,
	[SP_SLIP_FG] [char](1) NOT NULL,
	[sp_year] [int] NULL,
	[sp_month] [int] NULL,
	[Sum_Tot] [numeric](38, 6) NULL,
	[Sum_SALE_AMT] [numeric](38, 6) NULL,
	[Sum_Reject_AMT] [numeric](38, 6) NULL,
	[Sum_Tax] [numeric](38, 6) NULL,
	[Sum_PAmt] [numeric](38, 6) NULL,
	[Sum_MAmt] [numeric](38, 6) NULL,
	[Sum_Dis] [numeric](38, 6) NULL,
	[ct_advance] [numeric](38, 6) NULL,
	[Sum_All] [numeric](38, 6) NULL,
	[Sum_RecAmt] [numeric](38, 6) NULL,
	[Sum_PayAmt] [numeric](38, 6) NULL,
	[Sum_RealRecAmt] [numeric](38, 6) NULL,
	[ct_rem] [nvarchar](255) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
