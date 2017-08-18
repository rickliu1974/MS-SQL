USE [DW]
GO
/****** Object:  Table [dbo].[Sys_Ori_DataBases]    Script Date: 07/24/2017 14:43:53 ******/
DROP TABLE [dbo].[Sys_Ori_DataBases]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Sys_Ori_DataBases](
	[Application_Name] [varchar](100) NOT NULL,
	[Company_Name] [varchar](100) NOT NULL,
	[Server_Name] [varchar](100) NOT NULL,
	[DataBase_Name] [varchar](100) NOT NULL,
	[Enabled] [bit] NULL,
 CONSTRAINT [PK__Ori_Data__4CEFCD515CA75062] PRIMARY KEY CLUSTERED 
(
	[Application_Name] ASC,
	[Server_Name] ASC,
	[DataBase_Name] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
