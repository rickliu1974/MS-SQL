USE [DW]
GO
/****** Object:  UserDefinedFunction [dbo].[xFn_Check_EAN_Code]    Script Date: 08/18/2017 17:43:41 ******/
DROP FUNCTION [dbo].[xFn_Check_EAN_Code]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE FUNCTION [dbo].[xFn_Check_EAN_Code](@Ean varchar(max)) RETURNS INT AS
BEGIN
  /**************************************************************************************************************
    20151103 Rickliu 此函式從以下網路取得進行改寫
    http://stackoverflow.com/questions/29171859/sql-server-stored-procedure-ean-code-validation
  **************************************************************************************************************/
  DECLARE @Factor INT, @Sum INT, @Val INT, @Len INT, @CC INT, @CA INT
  DECLARE @Result NVARCHAR(MAX) 
  
  SET @Len = LEN(@Ean)
  SET @Sum = 0
  SET @Factor = 3
  Set @Val = 0
  
  IF @Len = 14 OR @Len = 13 OR @Len = 12 OR @Len = 8 
  BEGIN
    WHILE @Len > 0 
    BEGIN
       IF @Len < 13
       BEGIN
         SET @Sum = @Sum + SUBSTRING(@Ean, @Len, 1) * @Factor
         SET @Factor = 4 - @Factor
       END
       SET @Len = @Len - 1
    END

    SET @CC = ((10 - (@Sum % 10)) % 10)
    SET @CA = SUBSTRING(@Ean, LEN(@Ean), 1)

    Set @Result = @CC
  END
  ELSE
    SET @Result = -1
    
  RETURN (@Result)
END
GO
