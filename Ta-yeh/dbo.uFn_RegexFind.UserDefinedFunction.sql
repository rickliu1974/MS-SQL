USE [DW]
GO
/****** Object:  UserDefinedFunction [dbo].[uFn_RegexFind]    Script Date: 07/24/2017 14:44:01 ******/
DROP FUNCTION [dbo].[uFn_RegexFind]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create function [dbo].[uFn_RegexFind](
    @pattern VARCHAR(255),
    @matchstring VARCHAR(max),
    @global BIT = 1,
	@Multiline bit =1)
returns
    @result TABLE
        (
        Match_ID INT,
          FirstIndex INT ,
          length INT ,
          Value VARCHAR(2000),
          Submatch_ID INT,
          SubmatchValue VARCHAR(2000),
		  Error Varchar(255)
        )


AS -- columns returned by the function
	begin
    DECLARE @objRegexExp INT,
        @objErrorObject INT,
        @objMatch INT,
        @objSubMatches INT,
        @strErrorMessage VARCHAR(255),
		@error varchar(255),
        @Substituted VARCHAR(8000),
        @hr INT,
        @matchcount INT,
        @SubmatchCount INT,
        @ii INT,
        @jj INT,
        @FirstIndex INT,
        @length INT,
        @Value VARCHAR(2000),
        @SubmatchValue VARCHAR(2000),
        @objSubmatchValue INT,
        @command VARCHAR(8000),
        @Match_ID INT
        
    DECLARE @match TABLE
        (
          Match_ID INT IDENTITY(1, 1)
                       NOT NULL,
          FirstIndex INT NOT NULL,
          length INT NOT NULL,
          Value VARCHAR(2000)
        )    
    DECLARE @Submatch TABLE
        (
          Submatch_ID INT IDENTITY(1, 1),
          match_ID INT NOT NULL,
          SubmatchNo INT NOT NULL,
          SubmatchValue VARCHAR(2000)
        )
		


    SELECT  @strErrorMessage = 'creating a regex object',@error=''
    EXEC @hr= sp_OACreate 'VBScript.RegExp', @objRegexExp OUT
    IF @hr = 0 
        SELECT  @strErrorMessage = 'Setting the Regex pattern',
                @objErrorObject = @objRegexExp
    IF @hr = 0 
        EXEC @hr= sp_OASetProperty @objRegexExp, 'Pattern', @pattern
    IF @hr = 0 
        SELECT  @strErrorMessage = 'Specifying a case-insensitive match' 
    IF @hr = 0 
        EXEC @hr= sp_OASetProperty @objRegexExp, 'IgnoreCase', 1
    IF @hr = 0 
        EXEC @hr= sp_OASetProperty @objRegexExp, 'MultiLine', @Multiline
    IF @hr = 0 
        EXEC @hr= sp_OASetProperty @objRegexExp, 'Global', @global
    IF @hr = 0 
        SELECT  @strErrorMessage = 'Doing a match' 
    IF @hr = 0 
        EXEC @hr= sp_OAMethod @objRegexExp, 'execute', @objMatch OUT,
            @matchstring
    IF @hr = 0 
        SELECT  @strErrorMessage = 'Getting the number of matches'     
    IF @hr = 0 
        EXEC @hr= sp_OAGetProperty @objmatch, 'count', @matchcount OUT
    SELECT  @ii = 0 
    WHILE @hr = 0
        AND @ii < @Matchcount
        BEGIN
/*The Match object has four read-only properties. 
The FirstIndex property indicates the number of characters in the string to the left of the match. 
 The Length property of the Match object indicates the number of characters in the match. 
 The Value property returns the text that was matched.*/
            SELECT  @strErrorMessage = 'Getting the FirstIndex property',
                    @command = 'item(' + CAST(@ii AS VARCHAR) + ').FirstIndex'    
            IF @hr = 0 
                EXEC @hr= sp_OAGetProperty @objmatch, @command,
                    @Firstindex OUT
            IF @hr = 0 
                SELECT  @strErrorMessage = 'Getting the length property',
                        @command = 'item(' + CAST(@ii AS VARCHAR) + ').Length'    
            IF @hr = 0 
                EXEC @hr= sp_OAGetProperty @objmatch, @command, @Length OUT
            IF @hr = 0 
                SELECT  @strErrorMessage = 'Getting the value property',
                        @command = 'item(' + CAST(@ii AS VARCHAR) + ').Value'    
            IF @hr = 0 
                EXEC @hr= sp_OAGetProperty @objmatch, @command, @Value OUT
            INSERT  INTO @match
                    (
                      Firstindex,
                      [Length],
                      [Value]
                    )
                    SELECT  @firstindex + 1,
                            @Length,
                            @Value
            SELECT  @Match_ID = @@Identity			
/*The SubMatches property of the Match object is a collection of strings. It will only hold values if your regular expression has capturing groups. The collection will hold one string for each capturing group. The Count property indicates the number of string in the collection. The Item property takes an index parameter, and returns the text matched by the capturing group. The Item property is the default member, so you can write SubMatches(7) as a shorthand to SubMatches.Item(7). Unfortunately, VBScript does not offer a way to retrieve the match position and length of capturing groups.

*/
            IF @hr = 0 
                SELECT  @strErrorMessage = 'Getting the SubMatches collection',
                        @command = 'item(' + CAST(@ii AS VARCHAR)
                        + ').SubMatches'    
            IF @hr = 0 
                EXEC @hr= sp_OAGetProperty @objmatch, @command,
                    @objSubmatches OUT
            IF @hr = 0 
                SELECT  @strErrorMessage = 'Getting the number of submatches'     
            IF @hr = 0 
                EXEC @hr= sp_OAGetProperty @objSubmatches, 'count',
                    @submatchCount OUT
            SELECT  @jj = 0 
            WHILE @hr = 0
                AND @jj < @submatchCount
                BEGIN
                    IF @hr = 0 
                        SELECT  @strErrorMessage = 'Getting the submatch value property',
                                @command = 'item(' + CAST(@jj AS VARCHAR)
                                + ')' ,@submatchValue=null   
                    IF @hr = 0 
                        EXEC @hr= sp_OAGetProperty @objSubmatches, @command,
                            @SubmatchValue OUT
                    INSERT  INTO @Submatch
                            (
                              Match_ID,
                              SubmatchNo,
                              SubmatchValue
                            )
                            SELECT  @Match_ID,
                                    @jj+1,
                                    @SubmatchValue
                    SELECT  @jj = @jj + 1
                END		
            SELECT  @ii = @ii + 1
        END
    IF @hr <> 0 
        BEGIN
            DECLARE @Source VARCHAR(255),
                @Description VARCHAR(255),
                @Helpfile VARCHAR(255),
                @HelpID INT
	
            EXECUTE sp_OAGetErrorInfo @objErrorObject, @source OUTPUT,
                @Description OUTPUT, @Helpfile OUTPUT, @HelpID OUTPUT
            SELECT  @Error = 'Error whilst '
                    + COALESCE(@strErrorMessage, 'doing something') + ', '
                    + COALESCE(@Description, '')
        END
    EXEC sp_OADestroy @objRegexExp
     EXEC sp_OADestroy        @objMatch
     EXEC sp_OADestroy        @objSubMatches

insert into @result
          (Match_ID,
          FirstIndex,
          [length],
          [Value],
          Submatch_ID,
          SubmatchValue,
		  error)


    SELECT  m.[Match_ID],
			[FirstIndex],
			[length],
			[Value],[SubmatchNo],
			[SubmatchValue],@error
  FROM    @match m
    LEFT OUTER JOIN   @submatch	s
    ON m.match_ID=s.match_ID	
if @@rowcount=0 and len(@error)>0
insert into @result(error) select @error
 return 
end
GO
