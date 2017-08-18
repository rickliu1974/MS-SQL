USE [DW]
GO
/****** Object:  Table [dbo].[Rlistlog]    Script Date: 07/24/2017 14:43:52 ******/
DROP TABLE [dbo].[Rlistlog]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Rlistlog](
	[Editer] [varchar](50) NULL,
	[EditerCaption] [varchar](50) NULL,
	[Edittime] [datetime] NULL,
	[EditType] [varchar](5) NULL,
	[UserId] [varchar](500) NULL,
	[UserIdCaption] [varchar](500) NULL,
	[Project] [varchar](500) NULL,
	[ProjectCaption] [varchar](500) NULL,
	[SubSystem] [varchar](500) NULL,
	[SubSystemCaption] [varchar](500) NULL,
	[Query] [varchar](500) NULL,
	[QueryCaption] [varchar](500) NULL,
	[Rpt] [varchar](500) NULL,
	[RptCaption] [varchar](500) NULL,
	[R] [varchar](500) NULL,
	[ValueKind] [varchar](5) NULL,
	[AuthLog] [varchar](1) NULL,
	[IsGroup] [varchar](1) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
