USE [DW]
GO
/****** Object:  Table [dbo].[Weather_Data]    Script Date: 07/24/2017 14:43:54 ******/
DROP TABLE [dbo].[Weather_Data]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Weather_Data](
	[West_Date] [datetime] NULL,
	[Location] [nvarchar](4000) NULL,
	[hPa] [nvarchar](4000) NULL,
	[Temperature] [float] NULL,
	[Humidity] [float] NULL,
	[Wind_Direction] [nvarchar](4000) NULL,
	[Wind_Speed] [float] NULL,
	[Rain_Capacity] [float] NULL,
	[Sunshine_Houre] [float] NULL
) ON [PRIMARY]
GO
