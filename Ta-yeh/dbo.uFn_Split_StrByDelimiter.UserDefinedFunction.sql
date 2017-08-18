USE [DW]
GO
/****** Object:  UserDefinedFunction [dbo].[uFn_Split_StrByDelimiter]    Script Date: 08/18/2017 17:43:41 ******/
DROP FUNCTION [dbo].[uFn_Split_StrByDelimiter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE FUNCTION [dbo].[uFn_Split_StrByDelimiter] (@String VARCHAR(8000), @Delimiter CHAR(1), @TrimSpace Int = 0)
   RETURNS @temptable TABLE (id int, items VARCHAR(8000))
AS      
BEGIN      
   DECLARE @id Int, @idx Int
   DECLARE @slice Varchar(8000)
  
   set @idx = 1
   set @id = 0
   IF len(@String) < 1 OR @String IS NULL RETURN
  
   while @idx!= 0
   BEGIN
      SET @idx = charindex(@Delimiter, @String)
      set @id = @id + 1
      IF @idx <> 0
         SET @slice = LEFT(@String, @idx - 1)
      ELSE
         SET @slice = @String

      If @TrimSpace = 1
         Set @slice = RTrim(LTrim(@slice))

      IF (len(@slice) > 0)
         INSERT INTO @temptable(Id, Items) VALUES(@id, @slice)

      SET @String = RIGHT(@String, len(@String) - @idx)
      IF len(@String) = 0 break
   END
   RETURN
END
GO
