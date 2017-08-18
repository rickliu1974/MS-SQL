USE [DW]
GO
/****** Object:  Table [dbo].[Comp_EC_Consign_Stock_PCHOME]    Script Date: 07/24/2017 14:43:37 ******/
DROP TABLE [dbo].[Comp_EC_Consign_Stock_PCHOME]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Comp_EC_Consign_Stock_PCHOME](
	[ct_no] [varchar](50) NULL,
	[ct_sname] [varchar](50) NULL,
	[sk_no] [varchar](50) NULL,
	[sk_name] [varchar](50) NULL,
	[sk_bcode] [varchar](50) NULL,
	[xls_skno] [varchar](50) NULL,
	[xls_cono] [varchar](50) NULL,
	[xls_skname] [varchar](50) NULL,
	[fg6_qty] [numeric](38, 6) NOT NULL,
	[fg7_qty] [numeric](38, 6) NOT NULL,
	[sum_qty] [numeric](38, 6) NOT NULL,
	[xls_qty] [numeric](20, 6) NOT NULL,
	[diff_qty] [numeric](38, 6) NOT NULL,
	[fg6_amt] [numeric](38, 6) NOT NULL,
	[fg7_amt] [numeric](38, 6) NOT NULL,
	[sum_amt] [numeric](38, 6) NOT NULL,
	[Print_Date] [varchar](10) NOT NULL,
	[isfound] [varchar](100) NOT NULL,
	[Exec_DateTime] [datetime] NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
