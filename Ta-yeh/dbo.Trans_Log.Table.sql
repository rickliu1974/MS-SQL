USE [DW]
GO
/****** Object:  Table [dbo].[Trans_Log]    Script Date: 07/24/2017 14:43:54 ******/
DROP TABLE [dbo].[Trans_Log]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Trans_Log](
	[RowID] [int] IDENTITY(1,1) NOT NULL,
	[Trans_Date] [datetime] NOT NULL,
	[Process] [varchar](50) NOT NULL,
	[Msg] [varchar](max) NULL,
	[SqlCmd] [varchar](max) NULL,
	[RecordCount] [int] NULL,
	[Exec_Time] [time](7) NULL,
 CONSTRAINT [PK_Trans_Log] PRIMARY KEY CLUSTERED 
(
	[RowID] DESC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
