USE [DW]
GO
/****** Object:  Table [dbo].[Sys_Ori_Tables]    Script Date: 07/24/2017 14:43:53 ******/
DROP TABLE [dbo].[Sys_Ori_Tables]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Sys_Ori_Tables](
	[Application_Name] [varchar](100) NULL,
	[Company_Name] [varchar](100) NULL,
	[Server_Name] [varchar](100) NULL,
	[DataBase_Name] [varchar](100) NULL,
	[Table_Name] [varchar](200) NULL,
	[Enabled] [int] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
