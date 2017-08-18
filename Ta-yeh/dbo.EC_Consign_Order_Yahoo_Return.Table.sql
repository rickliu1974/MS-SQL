USE [DW]
GO
/****** Object:  Table [dbo].[EC_Consign_Order_Yahoo_Return]    Script Date: 07/24/2017 14:43:39 ******/
DROP TABLE [dbo].[EC_Consign_Order_Yahoo_Return]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[EC_Consign_Order_Yahoo_Return](
	[rowid] [int] NOT NULL,
	[ct_no] [nchar](10) NULL,
	[ct_sname] [nvarchar](12) NULL,
	[ct_ssname] [nvarchar](73) NULL,
	[sk_no] [nchar](30) NULL,
	[sk_name] [nchar](60) NULL,
	[f6] [nvarchar](255) NULL,
	[sk_bcode] [nchar](30) NULL,
	[f3] [nvarchar](255) NULL,
	[sale_qty] [int] NOT NULL,
	[Sale_amt] [numeric](29, 17) NOT NULL,
	[sale_tot] [numeric](18, 6) NOT NULL,
	[isfound] [varchar](1) NOT NULL,
	[import_date] [datetime] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
