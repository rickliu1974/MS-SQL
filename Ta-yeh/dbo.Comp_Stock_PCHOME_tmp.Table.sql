USE [DW]
GO
/****** Object:  Table [dbo].[Comp_Stock_PCHOME_tmp]    Script Date: 07/24/2017 14:43:38 ******/
DROP TABLE [dbo].[Comp_Stock_PCHOME_tmp]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Comp_Stock_PCHOME_tmp](
	[F1] [nvarchar](255) NULL,
	[F2] [nvarchar](255) NULL,
	[F3] [nvarchar](255) NULL,
	[F4] [nvarchar](255) NULL,
	[F5] [nvarchar](255) NULL,
	[F6] [nvarchar](255) NULL,
	[F7] [nvarchar](255) NULL,
	[F8] [nvarchar](255) NULL,
	[F9] [nvarchar](255) NULL,
	[xlsFileName] [varchar](255) NULL,
	[SplitFileName] [varchar](255) NULL,
	[imp_date] [datetime] NOT NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL,
	[print_date] [date] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
