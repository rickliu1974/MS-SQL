USE [DW]
GO
/****** Object:  Table [dbo].[EC_Order_SuperMarket_tmp]    Script Date: 07/24/2017 14:43:43 ******/
DROP TABLE [dbo].[EC_Order_SuperMarket_tmp]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[EC_Order_SuperMarket_tmp](
	[rowid] [int] IDENTITY(1,1) NOT NULL,
	[F1] [nvarchar](255) NULL,
	[F2] [varchar](8) NOT NULL,
	[F3] [nvarchar](255) NULL,
	[F4] [varchar](max) NULL,
	[F5] [nvarchar](255) NULL,
	[F6] [varchar](max) NULL,
	[F7] [nvarchar](255) NULL,
	[F8] [numeric](20, 6) NULL,
	[F9] [varchar](max) NULL,
	[F10] [varchar](100) NULL,
	[F11] [varchar](100) NULL,
	[F12] [varchar](100) NULL,
	[F13] [varchar](10) NOT NULL,
	[source_file] [varchar](255) NULL,
	[source_order] [varchar](max) NULL,
	[buyer] [nvarchar](255) NULL,
	[cust] [nvarchar](255) NULL,
	[cust_phone] [nvarchar](255) NULL,
	[cust_mobile] [nvarchar](255) NULL,
	[cust_addr1] [nvarchar](510) NULL,
	[cust_addr2] [nvarchar](510) NULL,
	[MEMO] [nvarchar](255) NULL,
	[print_date] [varchar](10) NOT NULL,
	[import_date] [datetime] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
