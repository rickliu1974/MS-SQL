USE [DW]
GO
/****** Object:  Table [dbo].[Scheduleu_log]    Script Date: 07/24/2017 14:43:52 ******/
DROP TABLE [dbo].[Scheduleu_log]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Scheduleu_log](
	[SchName] [varchar](50) NULL,
	[runtime] [datetime] NULL,
	[mailto] [varchar](250) NULL,
	[ParamNames] [varchar](1000) NULL,
	[msg] [text] NULL,
	[outFile] [varchar](200) NULL,
	[rptname] [varchar](50) NULL,
	[crptname] [varchar](50) NULL,
	[kind] [varchar](1) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
