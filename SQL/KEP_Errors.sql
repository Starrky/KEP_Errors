DECLARE @errordate_1 as date
DECLARE @errordate_2 as date


set @errordate_1 = "$(Date_1)" -- Enter start date  for error checking; date format: yyyy-mm-dd
set @errordate_2 = "$(Date_2)" -- Enter end date for error checking; date format: yyyy-mm-dd

-- KEP ERRORS
SELECT TOP(1000)
   [ErrorID],
   [UserName],
   [ErrorNumber],
   [ErrorState],
   [ErrorSeverity],
   [ErrorLine],
   [ErrorProcedure],
   [ErrorMessage],
   ((SELECT CAST(ErrorDateTime as date)
   [ErrorDateTime])
   ) AS 'ErrorDateTime'
   
   FROM [TPStockDB].[dbo].[DB_Errors] WHERE ErrorProcedure IS NOT NULL AND (ErrorDateTime >= @errordate_1 AND ErrorDateTime <= @errordate_2)
   