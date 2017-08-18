USE [DW]
GO
/****** Object:  Table [dbo].[Mall_Normal_Order_Car1_3c_tmp]    Script Date: 07/24/2017 14:43:45 ******/
DROP TABLE [dbo].[Mall_Normal_Order_Car1_3c_tmp]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Mall_Normal_Order_Car1_3c_tmp](
	[rowid] [int] IDENTITY(1,1) NOT NULL,
	[F1] [varchar](100) NULL,
	[F2] [nvarchar](255) NULL,
	[F3] [varchar](100) NULL,
	[F4] [nvarchar](255) NULL,
	[F5] [nvarchar](255) NULL,
	[F6] [nvarchar](255) NULL,
	[F7] [varchar](100) NULL,
	[F8] [varchar](100) NULL,
	[F9] [nvarchar](255) NULL,
	[F10] [nvarchar](255) NULL,
	[F11] [varchar](100) NULL,
	[F12] [varchar](10) NOT NULL,
	[F13] [nvarchar](255) NULL,
	[F14] [nvarchar](255) NULL,
	[F15] [varchar](100) NULL,
	[print_date] [varchar](10) NOT NULL,
	[xlsFileName] [varchar](255) NULL,
	[imp_date] [datetime] NOT NULL,
	[SP_Exec_date] [datetime] NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
