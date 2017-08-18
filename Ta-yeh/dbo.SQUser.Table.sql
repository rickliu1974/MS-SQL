USE [DW]
GO
/****** Object:  Table [dbo].[SQUser]    Script Date: 07/24/2017 14:43:53 ******/
DROP TABLE [dbo].[SQUser]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[SQUser](
	[UserID] [varchar](100) NULL,
	[Caption] [varchar](100) NULL,
	[Role] [varchar](1) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
