USE [DW]
GO
/****** Object:  Table [dbo].[Comp_Mall_Consign_Stock_Car1]    Script Date: 07/24/2017 14:43:38 ******/
DROP TABLE [dbo].[Comp_Mall_Consign_Stock_Car1]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Comp_Mall_Consign_Stock_Car1](
	[Kind1] [varchar](10) NULL,
	[Kind2] [varchar](10) NULL,
	[sd_ctno] [varchar](50) NULL,
	[ct_sname] [nchar](50) NULL,
	[sk_no] [varchar](30) NULL,
	[sk_name] [nchar](50) NULL,
	[xls_skno] [varchar](50) NOT NULL,
	[xls_skname] [varchar](50) NOT NULL,
	[sk_bcode] [varchar](50) NULL,
	[xls_bcode] [varchar](50) NOT NULL,
	[rowid] [int] NULL,
	[Chg_sd_Qty] [numeric](18, 2) NULL,
	[Chg_sd_stot] [numeric](18, 2) NULL,
	[xls_qty] [numeric](18, 2) NOT NULL,
	[xls_amt] [numeric](18, 2) NOT NULL,
	[diff_qty] [numeric](18, 2) NULL,
	[diff_amt] [numeric](18, 2) NULL,
	[Print_Date] [varchar](10) NULL,
	[isfound] [varchar](20) NOT NULL,
	[Exec_DateTime] [datetime] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
