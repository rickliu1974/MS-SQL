USE [DW]
GO
/****** Object:  Table [dbo].[SQRight]    Script Date: 07/24/2017 14:43:53 ******/
DROP TABLE [dbo].[SQRight]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[SQRight](
	[proj] [varchar](150) NULL,
	[pType] [varchar](1) NULL,
	[QryRptName] [varchar](150) NULL,
	[IsGroup] [varchar](1) NULL,
	[uid] [varchar](100) NULL,
	[SQRight] [varchar](10) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
