USE [DW]
GO
/****** Object:  Table [dbo].[Comp_Mall_Consign_Stock_Car1_Retail_tmp]    Script Date: 07/24/2017 14:43:38 ******/
DROP TABLE [dbo].[Comp_Mall_Consign_Stock_Car1_Retail_tmp]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Comp_Mall_Consign_Stock_Car1_Retail_tmp](
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
	[F11] [float] NULL,
	[F12] [nvarchar](255) NULL,
	[F13] [nvarchar](255) NULL,
	[F14] [nvarchar](255) NULL,
	[F15] [float] NULL,
	[F16] [float] NULL,
	[F17] [nvarchar](255) NULL,
	[XlsFileName] [varchar](255) NULL,
	[SplitFileName] [varchar](255) NULL,
	[Imp_date] [datetime] NOT NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL,
	[print_date] [date] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
