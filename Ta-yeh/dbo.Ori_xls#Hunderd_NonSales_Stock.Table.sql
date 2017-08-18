USE [DW]
GO
/****** Object:  Table [dbo].[Ori_xls#Hunderd_NonSales_Stock]    Script Date: 07/24/2017 14:43:49 ******/
DROP TABLE [dbo].[Ori_xls#Hunderd_NonSales_Stock]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Ori_xls#Hunderd_NonSales_Stock](
	[sk_no] [nvarchar](255) NULL,
	[sk_name] [nvarchar](255) NULL,
	[F3] [nvarchar](255) NULL,
	[XlsFileName] [varchar](255) NULL,
	[SplitFileName] [varchar](255) NULL,
	[Imp_date] [datetime] NOT NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
