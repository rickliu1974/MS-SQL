USE [DW]
GO
/****** Object:  Table [dbo].[Import_XLS_Schedule]    Script Date: 07/24/2017 14:43:43 ******/
DROP TABLE [dbo].[Import_XLS_Schedule]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Import_XLS_Schedule](
	[UniqueID] [int] IDENTITY(1,1) NOT NULL,
	[Folder_Name] [varchar](100) NULL,
	[xlsFileName] [varchar](100) NULL,
	[Flag] [varchar](1) NULL,
	[import_date] [datetime] NULL,
	[Update_date] [datetime] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
