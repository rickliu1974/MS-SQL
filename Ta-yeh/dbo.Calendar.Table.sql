USE [DW]
GO
/****** Object:  Table [dbo].[Calendar]    Script Date: 07/24/2017 14:43:37 ******/
DROP TABLE [dbo].[Calendar]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Calendar](
	[cal_date] [datetime] NOT NULL,
	[cal_year] [int] NULL,
	[cal_month] [int] NULL,
	[cal_day] [int] NULL,
	[cal_Quarter] [int] NULL,
	[cal_dayno] [int] NULL,
	[cal_week_begin] [datetime] NULL,
	[cal_week_end] [datetime] NULL,
	[cal_week_Range] [varchar](30) NULL,
	[cal_week] [int] NULL,
	[cal_week_name] [varchar](4) NULL,
	[cal_yearweek] [int] NULL,
 CONSTRAINT [PK_calender_cal_date] PRIMARY KEY CLUSTERED 
(
	[cal_date] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
