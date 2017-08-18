USE [DW]
GO
/****** Object:  Table [dbo].[EC_EZCat_List]    Script Date: 07/24/2017 14:43:40 ******/
DROP TABLE [dbo].[EC_EZCat_List]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[EC_EZCat_List](
	[UniqueID] [int] IDENTITY(1,1) NOT NULL,
	[Order_NO] [varchar](100) NULL,
	[EC_NO] [varchar](100) NULL,
	[EC_Name] [varchar](100) NULL,
	[sp_date] [varchar](10) NULL,
	[sp_date_year] [varchar](4) NULL,
	[sp_date_month] [varchar](2) NULL,
	[sp_date_day] [varchar](2) NULL,
	[sp_date_YM] [varchar](6) NULL,
	[sp_date_MD] [varchar](4) NULL,
	[sp_date_YMD] [varchar](8) NULL,
	[buyer] [varchar](100) NULL,
	[cust] [varchar](100) NULL,
	[cust_phone] [varchar](100) NULL,
	[cust_mobile] [varchar](100) NULL,
	[cust_addr1] [varchar](100) NULL,
	[cust_addr2] [varchar](100) NULL,
	[sk_no] [varchar](100) NULL,
	[sk_name] [varchar](100) NULL,
	[source_sk_name] [varchar](250) NULL,
	[isfound] [varchar](1) NULL,
	[qty] [varchar](100) NULL,
	[Price] [varchar](100) NULL,
	[MEMO] [varchar](255) NULL,
	[source_file] [varchar](255) NULL,
	[import_date] [datetime] NULL,
	[Rowid] [int] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
