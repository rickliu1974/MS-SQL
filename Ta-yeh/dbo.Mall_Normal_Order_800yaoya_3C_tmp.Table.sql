USE [DW]
GO
/****** Object:  Table [dbo].[Mall_Normal_Order_800yaoya_3C_tmp]    Script Date: 07/24/2017 14:43:44 ******/
DROP TABLE [dbo].[Mall_Normal_Order_800yaoya_3C_tmp]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Mall_Normal_Order_800yaoya_3C_tmp](
	[rowid] [int] IDENTITY(1,1) NOT NULL,
	[F1] [varchar](1000) NULL,
	[F2] [varchar](100) NULL,
	[F3] [varchar](100) NULL,
	[F4] [varchar](100) NULL,
	[F5] [varchar](100) NULL,
	[F6] [varchar](100) NULL,
	[F7] [varchar](100) NULL,
	[F8] [varchar](100) NULL,
	[F9] [varchar](100) NULL,
	[F10] [varchar](100) NULL,
	[F11] [varchar](100) NULL,
	[F12] [varchar](100) NULL,
	[F13] [varchar](100) NULL,
	[print_date] [varchar](10) NOT NULL,
	[import_date] [date] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
