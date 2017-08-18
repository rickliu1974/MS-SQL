USE [DW]
GO
/****** Object:  Table [dbo].[Ori_xls#Sys_Code]    Script Date: 07/24/2017 14:43:50 ******/
DROP TABLE [dbo].[Ori_xls#Sys_Code]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Ori_xls#Sys_Code](
	[code_class] [nvarchar](255) NULL,
	[Code_Begin] [nvarchar](255) NULL,
	[Code_End] [nvarchar](255) NULL,
	[Code_Name] [nvarchar](255) NULL,
	[Code_Remark1] [nvarchar](255) NULL,
	[Code_Remark2] [nvarchar](255) NULL,
	[XlsFileName] [varchar](255) NULL,
	[SplitFileName] [varchar](255) NULL,
	[Imp_date] [datetime] NOT NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
