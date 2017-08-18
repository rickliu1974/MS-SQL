USE [DW]
GO
/****** Object:  Table [dbo].[EC_Consign_Order_PChome_tmp]    Script Date: 07/24/2017 14:43:39 ******/
DROP TABLE [dbo].[EC_Consign_Order_PChome_tmp]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[EC_Consign_Order_PChome_tmp](
	[F1] [nvarchar](255) NULL,
	[F2] [nvarchar](255) NULL,
	[F3] [nvarchar](255) NULL,
	[F4] [nvarchar](255) NULL,
	[F5] [nvarchar](255) NULL,
	[F6] [nvarchar](255) NULL,
	[F7] [nvarchar](255) NULL,
	[F8] [nvarchar](255) NULL,
	[F9] [nvarchar](255) NULL,
	[F10] [nvarchar](255) NULL,
	[F11] [nvarchar](255) NULL,
	[F12] [nvarchar](255) NULL,
	[F13] [nvarchar](255) NULL,
	[F14] [nvarchar](255) NULL,
	[XlsFileName] [varchar](255) NULL,
	[SplitFileName] [varchar](255) NULL,
	[Imp_date] [datetime] NOT NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL,
	[print_date] [varchar](10) NOT NULL,
	[SP_Exec_date] [date] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
