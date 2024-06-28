CREATE PROCEDURE [dbo].[scriptTableData]
   @tableName nvarchar(64) = NULL,
   @whereClause nvarchar(2000) = NULL,
   @insertIdentityValues bit = 0,
   @executeScript bit = 0
AS
DECLARE @name nvarchar(64)
DECLARE @coltype nvarchar(64)
DECLARE @length int
DECLARE @isnullable int
DECLARE @isidentity int
DECLARE @temp nvarchar(max)
DECLARE @sql nvarchar(max)
DECLARE @field nvarchar(255)
DECLARE @isfirst tinyint
DECLARE @tableID int

SELECT @tableID = OBJECT_ID(@tableName)

IF (@tableID IS NULL) BEGIN
   DECLARE @myName nvarchar(128)

   SELECT @myName = OBJECT_NAME(@@PROCID)

   PRINT CASE WHEN COALESCE(@tableName, '') = '' THEN 'Missing' ELSE 'Invalid' END + ' @tableName'
   PRINT ''

   PRINT 'Generates a list of INSERT statements from rows from a table. These INSERT statements can then be used'
   PRINT 'to insert the data in to the same structured table in another database.'
   PRINT ''
   PRINT 'Usage: EXEC ' + @myName + ' @tableName, @whereClause, @insertIdentityValues, @executeScript'
   PRINT ''
   PRINT 'Where'
   PRINT '   @tableName: The name of database table whose data will be scripted.'
   PRINT '   @whereClause: Optional SQL WHERE clause used to filter the data to be scripted.'
   PRINT '      Do not include "WHERE" at the start of the clause.'
   PRINT '   @insertIdentityValues: Should be set to 1 if you wish to insert the same numeric values into identity columns.'
   PRINT '      If this is 1, the script will also print out SET IDENTITY_INSERT statements.'
   PRINT '   @executeScript: If set to 0, the SQL that would be used to generate the INSERT statuements will be printed.'
   PRINT '      If set to 1, the script is executed, and the INSERT statements will be output.'
   RETURN 1
END   
SET NOCOUNT ON

DECLARE curcol CURSOR READ_ONLY FOR
   SELECT sc.name, 
      coltype = st.name, 
      sc.length, 
      IsNullable = COLUMNPROPERTY(@tableID, sc.name, 'ALLOWSNULL'),
      IsIdentity = COLUMNPROPERTY(@tableID, sc.name, 'ISIDENTITY')
   FROM syscolumns sc
      JOIN sysobjects so ON sc.id = so.id
      JOIN systypes st ON st.xtype = sc.xtype
   WHERE (so.id = @tableID)
      AND (so.sysstat & 0xf = 3)
      AND st.xusertype != 256
   ORDER BY sc.colid

SELECT @temp = 'SELECT N''INSERT INTO ' + QUOTENAME(@tableName, '[') + ' ('
SELECT @sql = ') VALUES (''' + CHAR(13) + CHAR(9)
SELECT @isfirst = 1

OPEN curcol
FETCH NEXT FROM curcol INTO @name, @coltype, @length, @isnullable, @isidentity
WHILE @@fetch_status <> -1 BEGIN
   IF @@fetch_status <> -2 BEGIN
      SELECT @field = 'NULL'

      IF (@coltype IN ('varchar', 'char', 'nvarchar', 'nchar')) BEGIN
         IF @length = -1 BEGIN
            SELECT @field = N'CASE WHEN ' + QUOTENAME(@name, '[') + N' IS NULL THEN N''NULL'' ELSE N''N'''''' + REPLACE(' + QUOTENAME(@name, '[') + N', N'''''''', N'''''''''''') + N'''''''' END'
         END
         ELSE BEGIN
            SELECT @field = N'N''N'''''' + REPLACE(RTRIM(' + QUOTENAME(@name, '[') + N'), N'''''''', N'''''''''''') + N'''''''''
            IF (@isnullable <> 0)
               SELECT @field = N'COALESCE(' + @field + N', N''NULL'')'
         END
      END
      ELSE IF (@coltype IN ('bigint', 'bit', 'decimal', 'float', 'int', 'numeric', 'real', 'smallint', 'tinyint', 'money', 'smallmoney')) BEGIN
         SELECT @field = N'CONVERT(nvarchar, [' + @Name + N'])'
         IF (@isnullable <> 0)
            SELECT @field = N'COALESCE(' + @field + N', N''NULL'')'
      END
      ELSE IF (@coltype IN ('datetime', 'smalldatetime')) BEGIN
         SELECT @field = N'''N'' + QUOTENAME(CONVERT(nvarchar, [' + @name + N'], 121), '''''''')'
         IF (@isnullable <> 0)
            SELECT @field = N'COALESCE(' + @field + N', N''NULL'')'
      END
      ELSE IF (@coltype IN ('text', 'ntext')) BEGIN
         SELECT @field = N'QUOTENAME(REPLACE(CONVERT(nvarchar(max), [' + @name + N']), N'''''''', N''''''''''''), N'''''''')'
         IF (@isnullable <> 0)
            SELECT @field = N'COALESCE(' + @field + N', N''NULL'')'
      END
      ELSE IF (@coltype IN ('uniqueidentifier')) BEGIN
         SELECT @field = N'QUOTENAME(CONVERT(varchar, [' + @Name + N']), '''''''')'
         IF (@isnullable <> 0)
            SELECT @field = N'COALESCE(' + @field + N', ''NULL'')'
      END
      ELSE IF (@coltype IN ('binary', 'image', 'varbinary')) BEGIN
         SELECT @field = 'NULL'
      END

      IF (@isidentity = 0 OR @insertIdentityValues = 1) BEGIN
         IF (@isfirst = 0) BEGIN
            SELECT @temp = @temp + N', '
            SELECT @sql = @sql + N' + N'', ''' + CHAR(13) + CHAR(9)
         END
 
         SELECT @temp = @temp + N'[' + @Name + N']'
         SELECT @sql = @sql + N'+ ' + @field
         SELECT @isfirst = 0
      END
   END
   FETCH NEXT FROM curcol INTO @name, @coltype, @length, @isnullable, @isidentity
END
CLOSE curcol
DEALLOCATE curcol

SELECT @sql = @temp + @sql + N' + '')''' + CHAR(13) + N'FROM ' + @tableName + CHAR(13)

IF DATALENGTH(@whereClause) > 0 BEGIN
   SELECT @sql = @sql + N' WHERE (' + @whereClause + N')' + CHAR(13)
END

IF @insertIdentityValues = 1 BEGIN
   SELECT @temp = N'PRINT ''SET IDENTITY_INSERT ' + QUOTENAME(@tableName, '[')
   SELECT @sql = @temp + N' ON''' + CHAR(13) + @sql + @temp + N' OFF''' + CHAR(13)
END

IF @executeScript = 0 BEGIN
   SELECT @sql = N'SET NOCOUNT ON' + CHAR(13) + @sql + N'SET NOCOUNT OFF' + CHAR(13)
END

SELECT @temp = N'PRINT ''/* Table: ' + QUOTENAME(@tableName, '[') + ' */''' + CHAR(13)
SELECT @sql = @temp + @sql

IF @executeScript = 1 BEGIN
   EXEC(@sql)
END
ELSE BEGIN
   DECLARE @idx int

   -- If the resultant SQL is long, PRINT won't be able to print the whole
   -- thing, so we will print the SQL in chunks of 4000 characters.

   WHILE DATALENGTH(@sql) > 4000 BEGIN
      SELECT @idx = 4000

      WHILE @idx > 1 AND SUBSTRING(@sql, @idx, 1) != CHAR(13)
         SELECT @idx = @idx - 1

      IF @idx <= 1 BEGIN
         SELECT @idx = 1

         WHILE SUBSTRING(@sql, @idx, 1) != CHAR(13)
            SELECT @idx = @idx + 1
      END

      PRINT SUBSTRING(@sql, 1, @idx - 1)

      SELECT @sql = SUBSTRING(@sql, @idx + 1, DATALENGTH(@sql) - @idx)
   END

   PRINT @sql
END

SET NOCOUNT OFF
RETURN 0
