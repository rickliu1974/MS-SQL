USE [DW]
GO
/****** Object:  Table [dbo].[TB_Absent1]    Script Date: 07/24/2017 14:43:54 ******/
DROP TABLE [dbo].[TB_Absent1]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[TB_Absent1](
	[abs_day] [datetime] NULL,
	[abs_month] [int] NULL,
	[abs_week] [nvarchar](1) NULL,
	[abs_dept] [nvarchar](8) NULL,
	[abs_deptname] [nvarchar](20) NULL,
	[abs_emp_no] [nvarchar](10) NULL,
	[abs_emp_name] [nvarchar](10) NULL,
	[Chg_Duty_Level] [int] NULL,
	[Chg_Dept_Level] [nvarchar](8) NULL,
	[abs_stime] [varchar](8000) NULL,
	[abs_etime] [varchar](8000) NULL,
	[ps_tm2s] [char](5) NULL,
	[ps_tm2e] [char](5) NULL,
	[abs_seTime] [varchar](8) NULL,
	[abs_seTime_sec] [int] NOT NULL,
	[abs_ldate] [datetime] NULL,
	[ps_tm3s] [char](5) NULL,
	[ps_tm3e] [char](5) NULL,
	[hold_Name] [nvarchar](4000) NULL,
	[hold_Memo] [nvarchar](4000) NULL,
	[hold_chg] [varchar](1) NOT NULL,
	[hold_sdate] [datetime] NULL,
	[hold_edate] [datetime] NULL,
	[hold_seDay] [varchar](2) NULL,
	[hold_seTime] [varchar](8) NULL,
	[hold_seTime_sec] [int] NULL,
	[AccountID] [nvarchar](50) NULL,
	[UPdate_DateTime] [datetime] NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
