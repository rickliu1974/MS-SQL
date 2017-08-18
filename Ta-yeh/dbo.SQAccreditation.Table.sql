USE [DW]
GO
/****** Object:  Table [dbo].[SQAccreditation]    Script Date: 07/24/2017 14:43:53 ******/
DROP TABLE [dbo].[SQAccreditation]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[SQAccreditation](
	[UserID] [varchar](60) NULL,
	[Project] [varchar](60) NULL,
	[SubSystem] [varchar](60) NULL,
	[Query] [varchar](60) NULL,
	[Report] [varchar](60) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
