USE [DW]
GO
/****** Object:  UserDefinedFunction [dbo].[xOARegexMatch]    Script Date: 07/24/2017 14:44:01 ******/
DROP FUNCTION [dbo].[xOARegexMatch]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create FUNCTION [dbo].[xOARegexMatch] /* very simple Function Wrapper around the call */
    (
	  @objRegexExp INT,
      @matchstring VARCHAR(max)
    )
RETURNS INT
AS BEGIN
    DECLARE @objErrorObject INT,
        @hr INT,
        @match BIT
        EXEC @hr= sp_OAMethod @objRegexExp, 'Test', @match OUT, @matchstring
    IF @hr <> 0 
        BEGIN
            RETURN NULL
        END
    RETURN @match
   END
GO