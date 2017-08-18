USE [DW]
GO
/****** Object:  Table [dbo].[Combine_Lan_Orders]    Script Date: 07/24/2017 14:43:37 ******/
DROP TABLE [dbo].[Combine_Lan_Orders]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Combine_Lan_Orders](
	[SN] [int] IDENTITY(1,1) NOT NULL,
	[SO] [varchar](250) NULL,
	[CusName] [varchar](250) NULL,
	[TelDay] [varchar](250) NULL,
	[CellPhone] [varchar](250) NULL,
	[Address] [varchar](250) NULL,
	[Product] [varchar](250) NULL,
	[PriceInfo] [varchar](250) NULL,
	[Sender] [varchar](250) NULL,
	[SendAddr] [varchar](250) NULL,
	[Price] [int] NULL,
	[ImportDate] [datetime] NULL,
	[xlsFileName] [varchar](250) NULL,
	[Rowid] [int] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
