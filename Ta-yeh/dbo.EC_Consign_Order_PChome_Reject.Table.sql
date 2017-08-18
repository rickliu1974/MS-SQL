USE [DW]
GO
/****** Object:  Table [dbo].[EC_Consign_Order_PChome_Reject]    Script Date: 07/24/2017 14:43:39 ******/
DROP TABLE [dbo].[EC_Consign_Order_PChome_Reject]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[EC_Consign_Order_PChome_Reject](
	[rowid] [int] NOT NULL,
	[ct_no] [nchar](10) NULL,
	[ct_sname] [nvarchar](12) NULL,
	[ct_ssname] [nvarchar](73) NULL,
	[sk_no] [nchar](30) NULL,
	[sk_name] [nchar](60) NULL,
	[f6] [nvarchar](255) NULL,
	[sk_bcode] [nchar](30) NULL,
	[f3] [nvarchar](255) NULL,
	[sale_qty] [numeric](20, 6) NULL,
	[sale_amt] [numeric](38, 16) NULL,
	[sale_tot] [numeric](26, 10) NULL,
	[isfound] [varchar](1) NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
