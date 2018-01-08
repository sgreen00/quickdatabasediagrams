/*TODO: support foreign key relationships. This was built against a schema without any*/
DECLARE @fkOverride TABLE (colFkName NVARCHAR(128) NOT NULL, colFkNameFull NVARCHAR(255) NOT NULL)
INSERT INTO @fkOverride VALUES
    (N'Field1', N''), --to prevent a match
    (N'Field2', N' FK >- table.field') --to override
;WITH colFkInfer AS (
    --Returns the PK field name that foreign references would like go against for joining
    --If table has a field 'id' and it is a pk or identity, then use [TableName]Id as likely join name
    --fkDedup needs to be filtered to 1 to ensure that colFkName is unique
    SELECT c.object_id, c.column_id
         , ROW_NUMBER() OVER (
             PARTITION BY 
               IIF(c.name = N'id' AND (SELECT count(*) FROM sys.columns sysc WHERE sysc.object_id = c.object_id AND sysc.name=CONCAT(OBJECT_NAME(c.object_id), c.name)) = 0
                   , CONCAT(OBJECT_NAME(c.object_id), c.name)
                   , c.name) 
             ORDER BY CASE WHEN c.name = N'id' THEN 1 WHEN PATINDEX(OBJECT_NAME(c.object_id) + N'%', c.name) > 0 THEN 2 ELSE 3 END, c.column_id, c.object_id) fkDedup
         , IIF(c.name = N'id' AND (SELECT count(*) FROM sys.columns sysc WHERE sysc.object_id = c.object_id AND sysc.name=CONCAT(OBJECT_NAME(c.object_id), c.name)) = 0
               , CONCAT(OBJECT_NAME(c.object_id), c.name)
               , c.name) colFkName
         , CONCAT(N' FK >- ', OBJECT_NAME(c.object_id), N'.', c.name) colFkNameFull
    FROM sys.columns c
      INNER JOIN sys.tables t ON c.object_id = t.object_id
      LEFT JOIN (SELECT i.object_id, ic.column_id, pkc.is_pk_compound
                 FROM sys.indexes i
                   INNER JOIN sys.index_columns ic ON i.object_id = ic.object_id AND i.index_id = ic.index_id
                   INNER JOIN (SELECT object_id, index_id, CAST(CASE WHEN COUNT(*) = 1 THEN 0 ELSE 1 END AS BIT) is_pk_compound
                               FROM sys.index_columns
                               GROUP BY object_id, index_id) pkc ON i.object_id = pkc.object_id AND i.index_id = pkc.index_id
                 WHERE i.is_primary_key = 1) pk ON c.object_id = pk.object_id AND c.column_id = pk.column_id
    WHERE (pk.object_id IS NOT NULL OR c.is_identity = 1)
      AND ISNULL(pk.is_pk_compound, 0) = 0
      AND t.type = N'U' AND t.is_ms_shipped = 0
), meta AS (
    SELECT t.object_id, t.name tblName, c.name colName, c.column_id
         , CASE WHEN ty.name in(N'varchar', N'varbinary', N'nchar', N'float', N'char', N'binary') THEN CONCAT(ty.name, N'(', IIF(c.max_length < 0, N'max', CAST(c.max_length AS varchar(50))) , N')')
           WHEN ty.name in(N'numeric', N'decimal') THEN CONCAT(ty.name, N'(', c.precision, N',', c.scale, N')') 
           ELSE ty.name END colType
         , IIF(c.is_nullable = 1, N' NULL', N'') colNull
         , IIF(pk.object_id IS NOT NULL, N' PK', N'') colPk
         , IIF(c.is_identity = 1, N' IDENTITY', N'') colIdentity
         , pk.is_pk_compound
         , COALESCE(fkOver.colFkNameFull, fk.colFkNameFull, N'') colFkNameFull
    FROM sys.tables t
      INNER JOIN sys.columns c ON t.object_id = c.object_id
      INNER JOIN sys.types ty ON c.user_type_id = ty.user_type_id
      LEFT JOIN @fkOverride fkOver ON c.name = fkOver.colFkName
      LEFT JOIN colFkInfer fk ON c.name = fk.colFkName AND c.object_id != fk.object_id AND fk.fkDedup = 1
      LEFT JOIN (SELECT i.object_id, ic.column_id, pkc.is_pk_compound
                 FROM sys.indexes i
                   INNER JOIN sys.index_columns ic ON i.object_id = ic.object_id AND i.index_id = ic.index_id
                   INNER JOIN (SELECT object_id, index_id, CAST(CASE WHEN COUNT(*) = 1 THEN 0 ELSE 1 END AS BIT) is_pk_compound
                               FROM sys.index_columns
                               GROUP BY object_id, index_id) pkc ON i.object_id = pkc.object_id AND i.index_id = pkc.index_id
                 WHERE i.is_primary_key = 1) pk ON t.object_id = pk.object_id AND c.column_id = pk.column_id
    WHERE t.type = N'U' AND t.is_ms_shipped = 0
), result AS (
    SELECT DISTINCT tblName, 1 lvl, 0 column_id, CAST(tblName AS NVARCHAR(255)) ddl FROM meta
    UNION
    SELECT DISTINCT tblName, 2 lvl, 0 column_id, N'-' FROM meta
    UNION
    SELECT tblName, 3 lvl, column_id, CONCAT(colName, N' ', colType, colNull, colPk, colIdentity, colFkNameFull) FROM meta
    UNION
    SELECT DISTINCT tblName, 4 lvl, 0 column_id, N'' FROM meta
)
SELECT ddl FROM result ORDER BY tblName, lvl, column_id
