USE [DW]
GO
/****** Object:  Table [dbo].[Weather_Data_YM]    Script Date: 07/24/2017 14:43:54 ******/
DROP TABLE [dbo].[Weather_Data_YM]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Weather_Data_YM](
	[YM] [varchar](7) NOT NULL,
	[Location] [varchar](30) NOT NULL,
	[avg_Temperature] [decimal](10, 2) NULL,
	[avg_humidity] [decimal](10, 2) NULL,
	[sunshine_houre] [decimal](10, 2) NULL,
	[tot_rain_capacity] [decimal](10, 2) NULL,
	[tot_rain_days] [decimal](10, 2) NULL,
	[updated_datetime] [datetime] NULL,
PRIMARY KEY CLUSTERED 
(
	[YM] ASC,
	[Location] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
