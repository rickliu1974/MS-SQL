USE [DW]
GO
/****** Object:  Table [dbo].[Ori_Txt#Absent]    Script Date: 07/24/2017 14:43:46 ******/
DROP TABLE [dbo].[Ori_Txt#Absent]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Ori_Txt#Absent](
	[e_no] [varchar](10) NULL,
	[ps_week] [varchar](10) NULL,
	[ps_date] [datetime] NULL,
	[ps_time] [varchar](10) NULL,
	[ps_class] [int] NULL,
	[e_name] [varchar](20) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
