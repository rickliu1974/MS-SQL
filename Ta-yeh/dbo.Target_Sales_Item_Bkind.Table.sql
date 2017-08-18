USE [DW]
GO
/****** Object:  Table [dbo].[Target_Sales_Item_Bkind]    Script Date: 07/24/2017 14:43:54 ******/
DROP TABLE [dbo].[Target_Sales_Item_Bkind]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Target_Sales_Item_Bkind](
	[Kind] [varchar](5) NOT NULL,
	[Kind_name] [varchar](20) NULL,
	[Year] [varchar](4) NOT NULL,
	[Month] [varchar](2) NOT NULL,
	[Sales_no] [varchar](20) NOT NULL,
	[Sales_name] [varchar](20) NULL,
	[CT_no] [varchar](20) NULL,
	[CT_name] [varchar](255) NULL,
	[SK_no] [varchar](20) NULL,
	[SK_name] [varchar](255) NULL,
	[Bkind] [varchar](2) NULL,
	[Value] [float] NULL,
 CONSTRAINT [PK_Target_Sales_Item_Bkind] PRIMARY KEY CLUSTERED 
(
	[Kind] ASC,
	[Year] ASC,
	[Month] ASC,
	[Sales_no] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
