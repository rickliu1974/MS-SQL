USE [DW]
GO
/****** Object:  Table [dbo].[Targe_Cust_Stock_Qty]    Script Date: 07/24/2017 14:43:53 ******/
DROP TABLE [dbo].[Targe_Cust_Stock_Qty]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Targe_Cust_Stock_Qty](
	[Year] [varchar](4) NOT NULL,
	[ct_no] [varchar](20) NOT NULL,
	[ct_name] [varchar](255) NULL,
	[sk_no] [varchar](20) NOT NULL,
	[sk_name] [varchar](255) NULL,
	[Qty] [int] NULL,
	[imp_date] [datetime] NULL,
 CONSTRAINT [PK_Targe_Cust_Stock_Qty] PRIMARY KEY CLUSTERED 
(
	[Year] ASC,
	[ct_no] ASC,
	[sk_no] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
