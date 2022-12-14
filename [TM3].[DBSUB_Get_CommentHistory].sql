SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


ALTER PROCEDURE [TM3].[DBSUB_Get_CommentHistory]
	 @materialCode	bigint
	,@batch			bigint
	,@materialType	nvarchar(max)
AS

SET NOCOUNT ON;


DECLARE 
		@MaterialTypeID int = null;

SELECT
	@MaterialTypeID = [Id]
FROM
	[dbo].[TM3_WSStockReport_MasterMaterialTypes]
WHERE
	[Type] = @MaterialType;

SELECT STUFF((
			SELECT ' | ' + uc.Comment + '(' + CAST(CAST(uc.[Created datetime] as date) as nvarchar(10)) + ', ' + uc.[User] + ')'
			FROM [TM3].[WSStockReport_UserComments] uc
			WHERE uc.[Material code] = @materialCode 
			AND ISNULL(uc.Batch, '') = ISNULL(@batch, '')
			AND uc.[Material type id] = @MaterialTypeID
			ORDER BY uc.[Created datetime] DESC
			FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')