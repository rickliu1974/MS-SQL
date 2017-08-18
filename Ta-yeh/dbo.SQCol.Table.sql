USE [DW]
GO
/****** Object:  Table [dbo].[SQCol]    Script Date: 07/24/2017 14:43:53 ******/
DROP TABLE [dbo].[SQCol]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[SQCol](
	[proj] [varchar](150) NULL,
	[QryName] [varchar](150) NULL,
	[TblName] [varchar](150) NULL,
	[TblCaption] [varchar](150) NULL,
	[ColName] [varchar](150) NULL,
	[ColCaption] [varchar](150) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
