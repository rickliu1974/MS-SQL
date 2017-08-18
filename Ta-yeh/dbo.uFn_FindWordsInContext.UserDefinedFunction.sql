USE [DW]
GO
/****** Object:  UserDefinedFunction [dbo].[uFn_FindWordsInContext]    Script Date: 08/18/2017 17:43:41 ******/
DROP FUNCTION [dbo].[uFn_FindWordsInContext]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create FUNCTION [dbo].[uFn_FindWordsInContext]
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
/*
https://www.red-gate.com/simple-talk/sql/t-sql-programming/tsql-regular-expression-workbench/
Desc: 尋找字串內的任一個單字， 單字與單字之間用空白做區隔
  
SELECT * FROM dbo.FindWordsInContext('sadness farewell embark',
'Sunset and evening star,
And one clear call for me!
And may there by no moaning of the bar,
When I put out to sea,
 
But such a tide as moving seems asleep,
Too full for sound and foam,
When that which drew from out the boundless deep
Turns again home.
    
Twilight and evening bell,
And after that the dark!
And may there be no sadness of farewell,
When I embark;
 
For tho'' from out our bourne of Time and Place
The flood may bear me far,
I hope to see my Pilot face to face
When I have crost the bar.
',8)
*/

    DECLARE @Pattern VARCHAR(512)
    SELECT  @Pattern = COALESCE(@pattern + '(?:\W+\w+){0,'
                                + CAST(@proximity AS VARCHAR(5)) + '}?\W+',
                                '\b') + value
    FROM    dbo.uFn_RegexFind('\b[\w]+\b', @words, 1, 1)
    INSERT  INTO @ProximityList ( context )
            SELECT  '...' + SUBSTRING(@text, Firstindex - 8, length + 16)
                    + '...'
            FROM    dbo.uFn_RegexFind(@pattern+'\b', @text, 1, 1)
    RETURN
   END
GO
