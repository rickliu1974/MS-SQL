USE [DW]
GO
/****** Object:  Table [dbo].[SQ_Log]    Script Date: 07/24/2017 14:43:53 ******/
DROP TABLE [dbo].[SQ_Log]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[SQ_Log](
	[UserID] [varchar](50) NULL,
	[UserName] [varchar](50) NULL,
	[IP] [varchar](50) NULL,
	[Project] [varchar](50) NULL,
	[ProjectCaption] [varchar](50) NULL,
	[Query] [varchar](50) NULL,
	[QueryCaption] [varchar](50) NULL,
	[System] [varchar](50) NULL,
	[SystemCaption] [varchar](50) NULL,
	[Report] [varchar](50) NULL,
	[ReportCaption] [varchar](50) NULL,
	[StartTime] [datetime] NULL,
	[FinishTime] [datetime] NULL,
	[IsBatch] [varchar](50) NULL,
	[Message] [text] NULL,
	[Parameters] [text] NULL,
	[IsError] [varchar](2) NULL,
	[LOGTYPE] [varchar](50) NULL,
	[RECORDCOUNT] [varchar](50) NULL,
	[ins_datetime] [datetime] NOT NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
