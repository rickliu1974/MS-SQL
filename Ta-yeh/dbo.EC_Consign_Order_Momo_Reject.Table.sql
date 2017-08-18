USE [DW]
GO
/****** Object:  Table [dbo].[EC_Consign_Order_Momo_Reject]    Script Date: 07/24/2017 14:43:39 ******/
DROP TABLE [dbo].[EC_Consign_Order_Momo_Reject]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[EC_Consign_Order_Momo_Reject](
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
	[Sale_amt] [int] NOT NULL,
	[sale_tot] [int] NOT NULL,
	[isfound] [varchar](1) NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
