USE [DW]
GO
/****** Object:  Table [dbo].[Ori_xls#Product_Property]    Script Date: 07/24/2017 14:43:49 ******/
DROP TABLE [dbo].[Ori_xls#Product_Property]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Ori_xls#Product_Property](
	[kind] [nvarchar](255) NULL,
	[sk_no] [nvarchar](255) NULL,
	[sk_name] [nvarchar](255) NULL,
	[group_no] [nvarchar](255) NULL,
	[color] [nvarchar](255) NULL,
	[size] [nvarchar](255) NULL,
	[package] [nvarchar](255) NULL,
	[barcode_name] [nvarchar](255) NULL,
	[price] [nvarchar](255) NULL,
	[pic_6] [nvarchar](255) NULL,
	[pic_4] [nvarchar](255) NULL,
	[pic_2] [nvarchar](255) NULL,
	[main_pic] [nvarchar](255) NULL,
	[pic1] [nvarchar](255) NULL,
	[pic2] [nvarchar](255) NULL,
	[pic3] [nvarchar](255) NULL,
	[avg_price] [nvarchar](255) NULL,
	[product_property] [nvarchar](255) NULL,
	[gross_property] [nvarchar](255) NULL,
	[xlsFileName] [varchar](255) NULL,
	[SplitFileName] [varchar](255) NULL,
	[imp_date] [datetime] NOT NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
