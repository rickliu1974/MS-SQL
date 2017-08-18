USE [DW]
GO
/****** Object:  Table [dbo].[Apsnt_Tmp]    Script Date: 07/24/2017 14:43:37 ******/
DROP TABLE [dbo].[Apsnt_Tmp]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Apsnt_Tmp](
	[A_ps_date] [datetime] NULL,
	[ps_week] [nvarchar](1) NULL,
	[e_no] [varchar](10) NULL,
	[e_name] [varchar](20) NULL,
	[ps_time_s] [varchar](10) NULL,
	[ps_time_e] [varchar](10) NULL,
	[ps_time2_s] [varchar](10) NOT NULL,
	[ps_time2_e] [varchar](1) NOT NULL,
	[ps_time3_s] [varchar](10) NOT NULL,
	[ps_time3_e] [varchar](10) NOT NULL,
	[ps_time4_s] [varchar](5) NULL,
	[ps_time4_e] [varchar](5) NULL,
	[e_dept] [nchar](8) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
