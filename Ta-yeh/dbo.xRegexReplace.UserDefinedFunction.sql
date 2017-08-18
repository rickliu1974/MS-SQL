USE [DW]
GO
/****** Object:  UserDefinedFunction [dbo].[xRegexReplace]    Script Date: 08/18/2017 17:18:57 ******/
DROP FUNCTION [dbo].[xRegexReplace]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE FUNCTION [dbo].[xRegexReplace]
    (
      @pattern VARCHAR(255),
      @replacement VARCHAR(255),
      @Subject VARCHAR(max),
      @global BIT = 1,
	  @Multiline bit =1
    )
RETURNS VARCHAR(max)
/*The RegexReplace function takes three string parameters. The pattern (the regular expression) the replacement expression, and the subject string to do the manipulation to.

The replacement expression is once that can cause difficulties. You can specify an empty string as the @replacement text. This will cause the Replace method to return the subject string with all regex matches deleted from it (see "strip all HTML elements out of a string" below). 
To re-insert the regex match as part of the replacement, include $& in the replacement text. (see "find a #comment and add a TSQL --" below)
 If the regexp contains capturing parentheses, you can use backreferences in the replacement text. $1 in the replacement text inserts the text matched by the first capturing group, $2 the second, etc. up to $9. To include a literal dollar sign in the replacements, put two consecutive dollar signs in the string you pass to the Replace method.*/
AS BEGIN
    DECLARE @objRegexExp INT,
        @objErrorObject INT,
        @strErrorMessage VARCHAR(255),
        @Substituted VARCHAR(8000),
        @hr INT,
        @Replace BIT

    SELECT  @strErrorMessage = 'creating a regex object'
    EXEC @hr= sp_OACreate 'VBScript.RegExp', @objRegexExp OUT
    IF @hr = 0 
        SELECT  @strErrorMessage = 'Setting the Regex pattern',
                @objErrorObject = @objRegexExp
    IF @hr = 0 
        EXEC @hr= sp_OASetProperty @objRegexExp, 'Pattern', @pattern
    IF @hr = 0 /*By default, the regular expression is case sensitive. Set the IgnoreCase property to True to make it case insensitive.*/
        SELECT  @strErrorMessage = 'Specifying the type of match' 
    IF @hr = 0 
        EXEC @hr= sp_OASetProperty @objRegexExp, 'IgnoreCase', 1
    IF @hr = 0 
        EXEC @hr= sp_OASetProperty @objRegexExp, 'MultiLine', @Multiline
    IF @hr = 0 
        EXEC @hr= sp_OASetProperty @objRegexExp, 'Global', @global
    IF @hr = 0 
        SELECT  @strErrorMessage = 'Doing a Replacement' 
    IF @hr = 0 
        EXEC @hr= sp_OAMethod @objRegexExp, 'Replace', @Substituted OUT,
            @subject, @Replacement
     /*If the RegExp.Global property is False (the default), Replace will return the @subject string with the first regex match (if any) substituted with the replacement text. If RegExp.Global is true, the @Subject string will be returned with all matches replaced.*/   
    IF @hr <> 0 
        BEGIN
            DECLARE @Source VARCHAR(255),
                @Description VARCHAR(255),
                @Helpfile VARCHAR(255),
                @HelpID INT
	
            EXECUTE sp_OAGetErrorInfo @objErrorObject, @source OUTPUT,
                @Description OUTPUT, @Helpfile OUTPUT, @HelpID OUTPUT
            SELECT  @strErrorMessage = 'Error whilst '
                    + COALESCE(@strErrorMessage, 'doing something') + ', '
                    + COALESCE(@Description, '')
            RETURN @strErrorMessage
        END
    EXEC sp_OADestroy @objRegexExp
    RETURN @Substituted
   END
GO
