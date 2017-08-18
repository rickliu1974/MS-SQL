USE [DW]
GO
/****** Object:  UserDefinedFunction [dbo].[uFn_Count_Split_StrByDelimiter]    Script Date: 08/18/2017 17:43:41 ******/
DROP FUNCTION [dbo].[uFn_Count_Split_StrByDelimiter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE FUNCTION [dbo].[uFn_Count_Split_StrByDelimiter](@String VARCHAR(8000), @Delimiter CHAR(1))     
  RETURNS INT    
AS 
BEGIN     
   DECLARE @temptable TABLE (items VARCHAR(8000))   
   DECLARE @SplitCount INT
   DECLARE @idx INT      
   DECLARE @slice VARCHAR(8000)       
  
   SELECT @idx = 1       
   IF len(@String)<1 OR @String IS NULL  RETURN  0   
  
   while @idx!= 0       
   BEGIN      
      SET @idx = charindex(@Delimiter,@String)       
      IF @idx!=0       
         SET @slice = LEFT(@String,@idx - 1)       
      ELSE      
         SET @slice = @String       
  
      IF (len(@slice)>0)  
         INSERT INTO @temptable(Items) VALUES(@slice)       
  
      SET @String = RIGHT(@String,len(@String) - @idx)       
      IF len(@String) = 0 break       
   END  
   SET  @SplitCount=(SELECT COUNT(*) FROM  @temptable)
   RETURN  @SplitCount
 END
GO
