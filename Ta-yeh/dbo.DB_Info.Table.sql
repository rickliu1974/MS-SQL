USE [DW]
GO
/****** Object:  Table [dbo].[DB_Info]    Script Date: 07/24/2017 14:43:39 ******/
DROP TABLE [dbo].[DB_Info]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[DB_Info](
	[Server_Name] [varchar](100) NULL,
	[DB_Name] [varchar](100) NULL,
	[File_ID] [int] NULL,
	[File_Size] [float] NULL,
	[Used_Size] [float] NULL,
	[Over_Size] [float] NULL,
	[logic_Name] [varchar](100) NULL,
	[file_Name] [varchar](1000) NULL,
	[Query_Time] [datetime] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
