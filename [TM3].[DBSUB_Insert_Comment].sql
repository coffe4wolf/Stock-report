SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [TM3].[DBSUB_Insert_Comment]
	 @MaterialCode  bigint
	,@MaterialType  nvarchar(120)
	,@Comment		nvarchar(max) 
	,@Batch			nvarchar(120) = NULL
AS

	SET NOCOUNT ON;

	BEGIN TRY

	DECLARE 
		 @MaterialTypeID int = null;

	-- Turn mm type description into id.
	SELECT
		@MaterialTypeID = [Id]
	FROM
		[dbo].[TM3_WSStockReport_MasterMaterialTypes]
	WHERE
		[Type] = @MaterialType;


	INSERT INTO [TM3].[WSStockReport_UserComments] (
		 [Material code]
		,[Material type id]
		,[Comment]
		,[Batch]
	)
	VALUES (
		 @MaterialCode
		,@MaterialTypeID
		,@Comment
		,@Batch
	);


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
