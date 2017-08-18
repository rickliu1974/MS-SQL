USE [DW]
GO
/****** Object:  Table [dbo].[Ori_xls#Sys_StockCode]    Script Date: 07/24/2017 14:43:50 ******/
DROP TABLE [dbo].[Ori_xls#Sys_StockCode]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Ori_xls#Sys_StockCode](
	[code_level] [float] NULL,
	[code_no] [nvarchar](255) NULL,
	[up_code] [nvarchar](255) NULL,
	[code_name] [nvarchar](255) NULL,
	[print_name] [nvarchar](255) NULL,
	[update_date] [datetime] NULL,
	[XlsFileName] [varchar](255) NULL,
	[SplitFileName] [varchar](255) NULL,
	[Imp_date] [datetime] NOT NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
