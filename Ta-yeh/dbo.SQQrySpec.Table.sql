USE [DW]
GO
/****** Object:  Table [dbo].[SQQrySpec]    Script Date: 07/24/2017 14:43:53 ******/
DROP TABLE [dbo].[SQQrySpec]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[SQQrySpec](
	[Proj] [varchar](150) NULL,
	[ProjCaption] [varchar](150) NULL,
	[SubSys] [varchar](150) NULL,
	[SubSysCaption] [varchar](150) NULL,
	[Query] [varchar](150) NULL,
	[QueryCaption] [varchar](150) NULL,
	[Report] [varchar](150) NULL,
	[ReportCaption] [varchar](150) NULL,
	[IsVisible] [char](1) NULL,
	[DownLoadValue] [varchar](40) NULL,
	[Project_Seq] [int] NULL,
	[SYS_Seq] [int] NULL,
	[Qry_Seq] [int] NULL,
	[Rpt_Seq] [int] NULL,
	[SQ_Type] [varchar](2) NULL,
	[sUrl] [varchar](256) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
