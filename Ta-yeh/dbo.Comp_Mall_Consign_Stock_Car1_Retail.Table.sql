USE [DW]
GO
/****** Object:  Table [dbo].[Comp_Mall_Consign_Stock_Car1_Retail]    Script Date: 07/24/2017 14:43:38 ******/
DROP TABLE [dbo].[Comp_Mall_Consign_Stock_Car1_Retail]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Comp_Mall_Consign_Stock_Car1_Retail](
	[Kind1] [varchar](1) NOT NULL,
	[rowid] [int] NOT NULL,
	[ct_no] [varchar](50) NULL,
	[ct_sname] [varchar](50) NULL,
	[ct_ssname] [varchar](50) NULL,
	[sk_no] [varchar](50) NULL,
	[sk_name] [varchar](50) NULL,
	[F7] [varchar](50) NULL,
	[sk_bcode] [varchar](50) NULL,
	[F3] [varchar](50) NULL,
	[fg6_qty] [numeric](38, 6) NOT NULL,
	[fg7_qty] [numeric](38, 6) NOT NULL,
	[sum_qty] [numeric](38, 6) NOT NULL,
	[fg6_amt] [numeric](38, 6) NOT NULL,
	[fg7_amt] [numeric](38, 6) NOT NULL,
	[sum_amt] [numeric](38, 6) NOT NULL,
	[F13] [numeric](20, 6) NOT NULL,
	[F17] [numeric](20, 6) NOT NULL,
	[diff_qty] [numeric](38, 6) NOT NULL,
	[diff_amt] [numeric](38, 6) NOT NULL,
	[Print_Date] [varchar](10) NOT NULL,
	[isfound] [varchar](100) NOT NULL,
	[Exec_DateTime] [datetime] NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
