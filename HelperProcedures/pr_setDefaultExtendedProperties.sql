
/****** Object:  StoredProcedure [dbo].[pr_setDefaultExtendedProperty]    Script Date: 24.02.2021 13:32:30 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



IF OBJECT_ID('dbo.pr_setDefaultExtendedProperty') IS NOT NULL
DROP PROCEDURE [dbo].[pr_setDefaultExtendedProperty];
GO


CREATE PROCEDURE [dbo].[pr_setDefaultExtendedProperty]
(
	@paramSchemaNameDestination		sysname,			/* Schema des ZielObjektes */
	@paramObjectNameDestination		sysname,			/* Name des ZielObjektes */
	@paramObjectTypeDestination		VARCHAR(2),			/* Typ des Objektes */
	@paramExtendedPropertyAuthor	NVARCHAR(100)		/* Name des Authors, welcher in Extendend Properties gesetzt wird */
	
)
AS
	/* Diese Prozedur erzeugt vollautomatisiert zu einem übergebenen Objekt vom Typ
	 * Type des Objektes, 'U' für User Table, 'V' für View, 'P' für DQL_STORED_PROCEDURE, 'FN' für SQL_SCALAR_FUNCTION, 'TF' für SQL_TABLE_VALUED_FUNCTION, 
	 * 'IN' für SQL_INLINE_TABLE_VALUED_FUNCTION, 'TR' für TRIGGER, 'TT' für Tabellentyp
	 * die Extendend Properties mit Default Werten.
	 */
BEGIN
	DECLARE @objectNameDestination		NVARCHAR(1000),		/* Tabellenname, aus der die Spalten ermittelt werden sollen, welche dokumentiert werden sollen */
			@columnNameDestination		NVARCHAR(1000),		/* Spaltenname, für welche die Extended PRoteerties gesetzt werden sollen */
			@schemaNameDestination		NVARCHAR(1000),		/* Schemaname des zu dokumentierenden Objektes  */
			
			@objectTypeDestination		VARCHAR(2),			/* Type des Objektes, 'U' für User Table, 'V' für View, 'P' für DQL_STORED_PROCEDURE, 'FN' für SQL_SCALAR_FUNCTION, 'TF' für SQL_TABLE_VALUED_FUNCTION, 'IN' für SQL_INLINE_TABLE_VALUED_FUNCTION, 'TR' für TRIGGER	 */
			
			@objectTypeSourceName		NVARCHAR(100),		/* Name des Objektyps für die ExtendedProperty z.B. 'View', 'Table', 'Function', 'Procedure' */
			
			@propertyName				NVARCHAR(100),		/* Entended Property Author */
			@propertyValue				NVARCHAR(1000),		/* Entended Property Author */
			
			@columnForCursor			NVARCHAR(1000),		/* Hilfsvariable zur Verwendung im Cursor */
			@countForWhile				INT,				/* Hilfsvariable für Schleifendurchläufe */
			@columnCountForCursor		INT					/* Hilfsvariable für die Durchläufe des Cursors */
	;	

	/* Tabellenvariable mit Spaltennamen und Reihenfolge der jeweiligen Tabelle/View */
	DECLARE @columnsDestination AS TABLE 
			(
				columnName				sysname NOT NULL,		/* Spaltenname */
				columnOrdinalPosition	INT	NOT NULL,			/* Reihenfolge der Tabellen/Viewspalten */
				objectIdColumnBelogsTo  BIGINT NOT NULL			/* ObjectID der zugehörigen Tabelle */
			)	
	;
			
	/* Tabellenvariable mit der Liste der Erweiterten Eigenschaften, die für ein Objekt/Objektspalte gesetzt werden sollen*/
	DECLARE @extendedPropertyList AS TABLE 
			(
				extendedPropertyName			NVARCHAR(100),	/* Name der Erweiterten Eigenschaft */
				extendedPropertyValue			NVARCHAR(1000),	/* Wert der Erweiterten Eigenschaft */
				extendedPropertyOrdinalPosition	INT NOT NULL	/* Reihenfolgenposition der Erweiterten Eigenschaft */
			)
									
	/* Cursordefinition für die Spalten der Tabelle, in welche die ExtendedProterties geschrieben werden sollen */					
	DECLARE curAddDefaultExtendedPropertyToColumns CURSOR FOR
		SELECT 
			columnName,
			columnOrdinalPosition 
		FROM 
			@columnsDestination
	;	
	/* ================================================================================================================ */
	/* Parameterinhalte prüfen*/
	IF (CHARINDEX('.',LTRIM(RTRIM(@paramSchemaNameDestination))) <> 0)
	BEGIN
		PRINT 'Parameter @paramSchemaNameDestination hat ein falsches Format. Bitte im Format Objekt ohne Sonderzeichen "." angeben.';
		RETURN 1;
	END	
	;	
	IF (CHARINDEX('.',LTRIM(RTRIM(@paramObjectNameDestination))) <> 0)
	BEGIN
		PRINT 'Parameter @paramObjectNameDestination hat ein falsches Format. Bitte im Format Objekt ohne Sonderzeichen "." angeben.';
		RETURN 1;
	END	
	;
	/* Type des Objektes, 'U' für User Table, 'V' für View, 'P' für SQL_STORED_PROCEDURE, 'FN' für SQL_SCALAR_FUNCTION, 'TF' für SQL_TABLE_VALUED_FUNCTION, 'IN' für SQL_INLINE_TABLE_VALUED_FUNCTION, 'TR' für TRIGGER	 */
	IF (LTRIM(RTRIM(@paramObjectTypeDestination)) NOT IN ('U', 'V', 'P', 'FN', 'TF', 'IN', 'TR', 'TT') OR @paramObjectTypeDestination IS NULL)
	BEGIN
		PRINT 'Parameter @paramObjectTypeDestination hat eine falsche Eingabe.';
		PRINT 'Unterstützte Eingaben sind: ';
		PRINT 'U für User Table, '; 
		PRINT 'V für View, ';
		PRINT 'P für SQL_STORED_PROCEDURE, ';
		PRINT 'FN für SQL_SCALAR_FUNCTION, ';
		PRINT 'TF für SQL_TABLE_VALUED_FUNCTION, ';
		PRINT 'IN für SQL_INLINE_TABLE_VALUED_FUNCTION, ';
		PRINT 'TR für TRIGGER angeben.';
		PRINT 'TT für Tabellentyp angeben.';
		RETURN 1;
	END	
	
	/* ================================================================================================================ */	
	/* Setzen der DEFAULT Werte und Namen für die Erweiterten Eigenschaften */
	INSERT INTO @extendedPropertyList
			(
				extendedPropertyName,		
				extendedPropertyValue,			
				extendedPropertyOrdinalPosition	
			)
			SELECT 
				'Author', @paramExtendedPropertyAuthor, 1
			UNION ALL
			SELECT 
				'ChangeDate', '', 2
			UNION ALL
			SELECT 
				'ChangeHistory', '', 3
			UNION ALL
			SELECT
				'CreationDate', 
				RIGHT('0' + CAST(DAY(GETDATE()) AS VARCHAR(2)), 2) 
				+ '.' 
				+RIGHT('0' + CAST(MONTH(GETDATE()) AS VARCHAR(2)), 2) 
				+ '.' 
				+ CAST(YEAR(GETDATE()) AS VARCHAR(4)),
				4 
			UNION ALL
			SELECT
				'MS_Description', '', 5
	;
			
	SET @objectNameDestination =			@paramObjectNameDestination;
	SET @schemaNameDestination =			@paramSchemaNameDestination;
	SET @objectTypeDestination =			LTRIM(RTRIM(@paramObjectTypeDestination));

	SET @objectTypeSourceName =				CASE 
												WHEN @objectTypeDestination  = 'U'					THEN 'TABLE'
												WHEN @objectTypeDestination  = 'V'					THEN 'VIEW'
												WHEN @objectTypeDestination  IN ('IN', 'FN', 'TF')	THEN 'FUNCTION'
												WHEN @objectTypeDestination  = 'P'					THEN 'PROCEDURE'
												WHEN @objectTypeDestination  = 'TR'					THEN 'TRIGGER'
												WHEN @objectTypeDestination  = 'TT'					THEN 'TYPE'
											END
	;

	/* ================================================================================================================ */

	/* Wenn es sich bei dem Objekt um eine Tabelle oder eine View handelt, werden die Spalten ausgelesen */
	IF (@objectTypeDestination IN ('U','V'))
	BEGIN 
		/* Spalten der Quelltabelle/View auslesen */
		INSERT INTO @columnsDestination 
					(
					columnName,
					columnOrdinalPosition,
					objectIdColumnBelogsTo
					)
					SELECT  
						c.name, 
						c.column_id, 
						subObj.object_id
					FROM    
						sys.columns c INNER JOIN
						(
							SELECT object_id
							FROM	sys.objects obj INNER JOIN
									sys.schemas sch	ON obj.schema_id = sch.schema_id
							WHERE	(obj.name = @objectNameDestination OR (obj.type = 'TT' AND obj.name LIKE '%' + @objectNameDestination + '%'))
									AND (sch.name = @schemaNameDestination OR  (obj.type = 'TT' AND sch.name = 'sys'))
									AND obj.type = @objectTypeDestination
						) subObj ON c.object_id = subObj.object_id
	END				
					
	/* SELECT * FROM @columnsDestination */ 

	/* ================================================================================================================ */
	/* In diesem Block werden die vorgefinierten Extended Properties für das Objekt gesetzt. Im Vorfeld wird auf das Vorhandensein geprüft, wenn
	 * vorhanden, dann wird nichts verändert.
	 */
	
	PRINT 'Schema: ' + @schemaNameDestination;
	PRINT 'Objekt: ' + @objectNameDestination;
	PRINT 'Typ:    ' + @objectTypeSourceName;
	PRINT ' ';
	 
	SET @countForWhile = 1;	/* Anzahl der zu setzenden Properties */
	/* Es wird solange widerholt, bis jede der gesetzten Extended Properties gesetzt ist */
	WHILE @countForWhile <= (SELECT COUNT(extendedPropertyName) FROM @extendedPropertyList epl)  
	BEGIN
		SET @propertyName = (SELECT epl.extendedPropertyName FROM @extendedPropertyList epl WHERE epl.extendedPropertyOrdinalPosition = @countForWhile);
		SET @propertyValue = (SELECT epl.extendedPropertyValue FROM @extendedPropertyList epl WHERE epl.extendedPropertyOrdinalPosition = @countForWhile);
		
		IF NOT EXISTS
		(	
			SELECT * 
			FROM sys.extended_properties ep 
			WHERE 
					ep.minor_id = 0 
				AND ep.name = (SELECT epl.extendedPropertyName FROM @extendedPropertyList epl WHERE epl.extendedPropertyOrdinalPosition	= @countForWhile)
				AND ep.major_id = (
									SELECT object_id 
									FROM sys.objects obj INNER JOIN
									sys.schemas sch	ON obj.schema_id = sch.schema_id 
									WHERE	obj.name = @objectNameDestination
										AND sch.name = @schemaNameDestination
									)
		)
		BEGIN
			EXEC sys.sp_addextendedproperty	@name = @propertyName,			@value = @propertyValue , 
										@level0type = N'SCHEMA',			@level0name = @schemaNameDestination, 
										@level1type = @objectTypeSourceName,@level1name = @objectNameDestination
		END
		ELSE
		BEGIN
			PRINT 'Extended Property ''' + @propertyName + '''' +  REPLICATE(' ',20 - LEN(@propertyName)) 
				  + ' ist bereits gesetzt. Es wurde keine Aktion durchgeführt.';
		END
		
		SET @countForWhile = @countForWhile + 1;
		
		/*PRINT 'Extended Property ''' + @propertyName + '''' +  REPLICATE(' ',20 - LEN(@propertyName)) 
				  + ' ist bereits gesetzt. Es wurde keine Aktion durchgeführt.';
		*/
	END
	;
	PRINT '========================================================================================================================';

	/* ================================================================================================================ */	

	/* Nur wenn es sich bei dem übegebenen Objekt um eine Tabelle oder View oder Tabellentyp handelt, werden die Extended Propertier des Spalten gesetzt */
	IF (@objectTypeDestination IN ('U','V', 'TT'))
	BEGIN 	

		/* Cursor durchläuft die Spaltenliste und setzt die 4 vorgefinierten Extended Properties Author, ChangeDate, ChangeHistory, CreationDate, MS_Description*/
		OPEN curAddDefaultExtendedPropertyToColumns;
		FETCH NEXT FROM curAddDefaultExtendedPropertyToColumns INTO @columnForCursor, @columnCountForCursor WHILE @@FETCH_STATUS = 0	
		BEGIN
			PRINT 'Schema:  ' + @schemaNameDestination;
			PRINT 'Objekt:  ' + @objectNameDestination;
			PRINT 'Typ:     ' + @objectTypeSourceName;
			PRINT 'Spalte:  ' + @columnForCursor;
			PRINT ' ';
			
			SET @countForWhile = 1;	/* Anzahl der zu setzenden Properties */

			/* Schleife durchläuft genau die Anzahl der zu setzenden Erweiterten Properties (vorgegeben durch Tabellenvariable) und setzt diese */
			WHILE @countForWhile <= (SELECT COUNT(extendedPropertyName) FROM @extendedPropertyList epl) 
			BEGIN
				SET @propertyName = (SELECT epl.extendedPropertyName FROM @extendedPropertyList epl WHERE epl.extendedPropertyOrdinalPosition = @countForWhile);
				SET @propertyValue = (SELECT epl.extendedPropertyValue FROM @extendedPropertyList epl WHERE epl.extendedPropertyOrdinalPosition = @countForWhile);
		
				IF NOT EXISTS
				(	
					SELECT * 
					FROM sys.extended_properties ep 
					WHERE 
							ep.minor_id = @columnCountForCursor
						AND ep.name = @propertyName
						AND ep.major_id =	(
												SELECT object_id 
												FROM sys.objects obj INNER JOIN
												sys.schemas sch	ON obj.schema_id = sch.schema_id 
												WHERE	obj.name = @objectNameDestination
													AND sch.name = @schemaNameDestination
											)
				)
				BEGIN
					EXEC sys.sp_addextendedproperty	@name = @propertyName,				@value = @propertyValue, 
													@level0type = N'SCHEMA',			@level0name = @schemaNameDestination, 
													@level1type = @objectTypeSourceName,@level1name = @objectNameDestination,
													@level2type = N'COLUMN',			@level2name = @columnForCursor;
					PRINT 'Extended Property ''' + @propertyName + '''' + REPLICATE(' ',20 - LEN(@propertyName)) 
							+ ' Neuer Wert ''' + @propertyValue + '''';


				END
				ELSE
				BEGIN
					PRINT 'Extended Property ''' + @propertyName + '''' +  REPLICATE(' ',20 - LEN(@propertyName)) 
						+ ' ist bereits gesetzt. Es wurde keine Aktion durchgeführt.';
				END /* OF IF */
				
			SET @countForWhile = @countForWhile + 1;	
			
			/*PRINT 'Extended Property ''' + @propertyName + '''' + REPLICATE(' ',20 - LEN(@propertyName)) 
					+ ' Neuer Wert ''' + @propertyValue + '''';
			*/
			
			END /* OF WHILE */
			
			PRINT '========================================================================================================================';
			
			FETCH NEXT FROM curAddDefaultExtendedPropertyToColumns INTO @columnForCursor, @columnCountForCursor    	
		END
		CLOSE curAddDefaultExtendedPropertyToColumns;
		DEALLOCATE curAddDefaultExtendedPropertyToColumns;
	END /* OF IF */
END /* OF PROCEDURE */



GO

EXEC sys.sp_addextendedproperty @name=N'Author', @value=N'Roman Anschiz' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'PROCEDURE',@level1name=N'pr_setDefaultExtendedProperty'
GO

EXEC sys.sp_addextendedproperty @name=N'ChangeDate', @value=N'24.02.2020' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'PROCEDURE',@level1name=N'pr_setDefaultExtendedProperty'
GO

EXEC sys.sp_addextendedproperty @name=N'ChangeHistory', @value=N'24.02.2020, Martin Guth: Erweiterung für Tabellentypen. 27.02.2013, Roman Anschiz: Anpassung der PRINT-Ausgabe bei Veränderung der ExtendedProperties. 08.02.2013 Erläuternde Ausgaben über PRINT hinzugefügt, um die gesetzten Extended Properties im Meldung-Fenster nachvollziehen zu können' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'PROCEDURE',@level1name=N'pr_setDefaultExtendedProperty'
GO

EXEC sys.sp_addextendedproperty @name=N'CreationDate', @value=N'21.06.2013' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'PROCEDURE',@level1name=N'pr_setDefaultExtendedProperty'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Die Prozedur setzt unter Angabe des Autors und des betroffenen Objektes die DEFAULT Extended Properties: Author, ChangeDate, Change History, CreationDate, MS_Description' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'PROCEDURE',@level1name=N'pr_setDefaultExtendedProperty'
GO


