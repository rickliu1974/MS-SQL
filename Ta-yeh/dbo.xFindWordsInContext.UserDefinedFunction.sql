USE [DW]
GO
/****** Object:  UserDefinedFunction [dbo].[xFindWordsInContext]    Script Date: 08/18/2017 17:18:57 ******/
DROP FUNCTION [dbo].[xFindWordsInContext]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create FUNCTION [dbo].[xFindWordsInContext]
    (
      @words VARCHAR(255),--list of words you want searched for
      @text VARCHAR(MAX),--the text you want searched 
      @proximity INT--the maximum distance in words between specified words
    )
RETURNS @proximityList TABLE
    (
      Hit INT IDENTITY(1, 1),
      context VARCHAR(2000)
    )
AS BEGIN
    DECLARE @Pattern VARCHAR(512)
    SELECT  @Pattern = COALESCE(@pattern + '(?:\W+\w+){0,'
                                + CAST(@proximity AS VARCHAR(5)) + '}?\W+',
                                '\b') + value
    FROM    dbo.RegexFind('\b[\w]+\b', @words, 1, 1)
    INSERT  INTO @ProximityList ( context )
            SELECT  '...' + SUBSTRING(@text, Firstindex - 8, length + 16)
                    + '...'
            FROM    dbo.RegexFind(@pattern+'\b', @text, 1, 1)
    RETURN
   END
GO
