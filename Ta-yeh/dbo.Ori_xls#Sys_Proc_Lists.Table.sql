USE [DW]
GO
/****** Object:  Table [dbo].[Ori_xls#Sys_Proc_Lists]    Script Date: 07/24/2017 14:43:50 ******/
DROP TABLE [dbo].[Ori_xls#Sys_Proc_Lists]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Ori_xls#Sys_Proc_Lists](
	[NO] [float] NULL,
	[Imp_Flag] [float] NULL,
	[Exec_Staut] [nvarchar](255) NULL,
	[Proc_Name] [nvarchar](255) NULL,
	[Proc_Desc] [nvarchar](255) NULL,
	[Type] [nvarchar](255) NULL,
	[Exec_Order] [nvarchar](255) NULL,
	[Folder_Name] [nvarchar](255) NULL,
	[FileName] [nvarchar](255) NULL,
	[Import_Date] [datetime] NULL,
	[Imp_FileName] [nvarchar](255) NULL,
	[Update_Date] [datetime] NULL,
	[Exec_Date] [datetime] NULL,
	[Exec_Msg] [nvarchar](255) NULL,
	[XlsFileName] [varchar](255) NULL,
	[SplitFileName] [varchar](255) NULL,
	[Imp_date] [datetime] NOT NULL,
	[rowid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
