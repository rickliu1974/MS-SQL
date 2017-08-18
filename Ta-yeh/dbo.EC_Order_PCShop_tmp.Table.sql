USE [DW]
GO
/****** Object:  Table [dbo].[EC_Order_PCShop_tmp]    Script Date: 07/24/2017 14:43:42 ******/
DROP TABLE [dbo].[EC_Order_PCShop_tmp]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[EC_Order_PCShop_tmp](
	[rowid] [int] IDENTITY(1,1) NOT NULL,
	[F1] [nvarchar](255) NULL,
	[F2] [varchar](8) NOT NULL,
	[F3] [nvarchar](255) NULL,
	[F4] [varchar](max) NULL,
	[F5] [nvarchar](4000) NULL,
	[F6] [varchar](max) NULL,
	[F7] [nvarchar](255) NULL,
	[F8] [varchar](max) NULL,
	[F9] [varchar](max) NULL,
	[F10] [varchar](100) NULL,
	[F11] [varchar](100) NULL,
	[F12] [varchar](100) NULL,
	[F13] [varchar](10) NOT NULL,
	[source_file] [varchar](255) NULL,
	[source_order] [varchar](max) NULL,
	[buyer] [nvarchar](4000) NULL,
	[cust] [nvarchar](4000) NULL,
	[cust_phone] [nvarchar](4000) NULL,
	[cust_mobile] [nvarchar](4000) NULL,
	[cust_addr1] [nvarchar](4000) NULL,
	[cust_addr2] [nvarchar](4000) NULL,
	[MEMO] [nvarchar](4000) NULL,
	[print_date] [varchar](10) NOT NULL,
	[import_date] [datetime] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
