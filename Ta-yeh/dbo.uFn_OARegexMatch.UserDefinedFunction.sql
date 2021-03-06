USE [DW]
GO
/****** Object:  UserDefinedFunction [dbo].[uFn_OARegexMatch]    Script Date: 08/18/2017 17:43:41 ******/
DROP FUNCTION [dbo].[uFn_OARegexMatch]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create FUNCTION [dbo].[uFn_OARegexMatch] /* very simple Function Wrapper around the call */
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
