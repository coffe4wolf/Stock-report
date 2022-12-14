SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO






ALTER PROCEDURE [TM3].[DBSUB_UpdateSafetyStock]
	 @MaterialCode		 bigint
	,@SafetyStockEntered decimal(18, 3) = null
AS

SET NOCOUNT ON;

BEGIN TRY

	IF @SafetyStockEntered IS NOT NULL
		BEGIN

			UPDATE 
				[dbo].[TM3_WSStockReport_MasterMaterialsList]
			SET
				[Safety stock (Entered)] = @SafetyStockEntered
			WHERE
				[Material code] = @MaterialCode;

		END;

END TRY

BEGIN CATCH
	
	SELECT  
		 ERROR_NUMBER()		AS ErrorNumber  
		,ERROR_SEVERITY()	AS ErrorSeverity  
		,ERROR_STATE()		AS ErrorState  
		,ERROR_PROCEDURE()	AS ErrorProcedure  
		,ERROR_LINE()		AS ErrorLine  
		,ERROR_MESSAGE()	AS ErrorMessage;

END CATCH

