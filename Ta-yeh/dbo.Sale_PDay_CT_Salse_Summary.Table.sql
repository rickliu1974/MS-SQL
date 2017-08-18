USE [DW]
GO
/****** Object:  Table [dbo].[Sale_PDay_CT_Salse_Summary]    Script Date: 07/24/2017 14:43:52 ******/
DROP TABLE [dbo].[Sale_PDay_CT_Salse_Summary]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Sale_PDay_CT_Salse_Summary](
	[sp_ctno] [nvarchar](10) NULL,
	[ct_sname] [nvarchar](12) NULL,
	[Chg_BU_NO] [nvarchar](6) NULL,
	[ct_fld3] [nvarchar](60) NULL,
	[ct_no8] [nvarchar](8) NULL,
	[ct_sname8] [nvarchar](121) NULL,
	[E_NO] [nvarchar](10) NULL,
	[e_name] [nvarchar](10) NULL,
	[e_dept] [nvarchar](8) NULL,
	[SP_SLIP_FG] [char](1) NOT NULL,
	[sp_year] [int] NULL,
	[sp_month] [int] NULL,
	[Sum_Tot] [decimal](38, 9) NULL,
	[Sum_SALE_AMT] [decimal](38, 9) NULL,
	[Sum_Reject_AMT] [decimal](38, 9) NULL,
	[Sum_Tax] [decimal](38, 9) NULL,
	[Sum_PAmt] [decimal](38, 9) NULL,
	[Sum_MAmt] [decimal](38, 9) NULL,
	[Sum_Dis] [decimal](38, 9) NULL,
	[ct_advance] [decimal](38, 6) NULL,
	[Sum_All] [decimal](38, 9) NULL,
	[Sum_RecAmt] [decimal](38, 9) NULL,
	[Sum_PayAmt] [decimal](38, 9) NULL,
	[Sum_RealRecAmt] [decimal](38, 6) NULL,
	[ct_rem] [nvarchar](255) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
