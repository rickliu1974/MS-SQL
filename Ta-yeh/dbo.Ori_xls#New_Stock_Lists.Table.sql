USE [DW]
GO
/****** Object:  Table [dbo].[Ori_xls#New_Stock_Lists]    Script Date: 07/24/2017 14:43:49 ******/
DROP TABLE [dbo].[Ori_xls#New_Stock_Lists]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Ori_xls#New_Stock_Lists](
	[sk_no] [nvarchar](255) NULL,
	[sm_no] [nvarchar](255) NULL,
	[sk_name] [nvarchar](255) NULL,
	[arrival_date] [nvarchar](255) NULL,
	[first_qty] [float] NULL,
	[F6] [float] NULL,
	[XlsFileName] [varchar](255) NULL,
	[SplitFileName] [varchar](255) NULL,
	[Imp_date] [datetime] NOT NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
