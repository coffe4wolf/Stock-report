SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [TM3].[DBSUB_WSStockReport_LoadReport]
	@MaterialType nvarchar(40) = NULL
AS 

SET NOCOUNT ON;

DECLARE 
	 @StoragePlant					nvarchar(4) = '2100'
	,@ContractorsPlant				nvarchar(4) = '2000'
	,@SecondVirtualPlant			nvarchar(4) = '2VRT'
	,@MovementTypeIncome			nvarchar(3)	= '101'
	,@MovementTypeIncomeCancel		nvarchar(3)	= '102'
	,@PurchaseDocMove				nvarchar(3) = '44%'
	,@PurchaseDocPurchOrd			nvarchar(3) = '45%'
	,@StockFlagSpecial				nvarchar(1) = 'O'
	,@StockFlagProject				nvarchar(1) = 'M'
	,@FreeWBSTemplate				nvarchar(6) = 'F.%'
	,@DateFormat					nvarchar(10) = 'dd.MM.yyyy'
	,@UsdRubRateConstantName		nvarchar(120) = N'Корпоративный курс доллара'
	,@CurrencyRub					nvarchar(20)  = N'Рубль'
	,@CurrencyEur					nvarchar(20)  = N'Евро%'
	,@CurrencyUsd					nvarchar(20)  = N'Доллар США%'
	,@CurrencyFunt					nvarchar(20)  = N'Фунт%'
	,@UsdRubRate					int;

SET @UsdRubRate				= (SELECT [Value] FROM dbo.Master_DbConstants WHERE [Description] = @UsdRubRateConstantName)

DECLARE
	 @EurRubRate			as decimal(18,2) = @USDRubRate * 1.38
	,@FuntRubRate			as decimal(18,2) = @USDRubRate * 1.5;

	IF OBJECT_ID('tempdb..#POQty')	IS NOT NULL  DROP TABLE #POQty;

	SELECT
		 [Материал]
		,[ЗКЗП: Дата поставки]
		,[Quantity in POs] = SUM([ЗКЗП: Количество]) - ISNULL(SUM([Qty in unit of entry]), 0)
	INTO #POQty
	FROM
		dbo.BEX_Analytical_Report b
		LEFT JOIN (
					SELECT
						 Material
						,[Delivery Order]
						,[Position]
						,[Qty in unit of entry]	= SUM([Qty in unit of entry])
					FROM
						dbo.Materials_Movements_SAP4HANA_new
					WHERE
						[Movement type] IN (@MovementTypeIncome, @MovementTypeIncomeCancel)
						AND Plant IN (@StoragePlant, @SecondVirtualPlant)
					GROUP BY
							Material
						,[Delivery Order]
						,[Position]
					) mv
			ON b.[ЗКЗ: Номер заказа на поставку] = mv.[Delivery Order]
			AND b.Материал = mv.Material
			AND b.[ЗКЗП: Позиция заказа на поставку] = mv.position
	WHERE 
			b.[ЗКЗ: Номер заказа на поставку] IS NOT NULL 
		AND b.[ЗКЗ: Номер заказа на поставку] NOT LIKE @PurchaseDocMove
		AND b.[ЗКЗП: Индикатор удаления] IS NULL 
		AND b.[ЗКЗП: Конечная поставка] = 0
	GROUP BY
		 [Материал]
		,[ЗКЗП: Дата поставки];


WITH 
	materials 
	AS (
		SELECT
			l.[Material code]
		FROM
			dbo.TM3_WSStockReport_MasterMaterialsList l
			LEFT JOIN dbo.TM3_WSStockReport_MasterMaterialTypes t
				ON l.[Type id] = t.Id
		),

	POQuantity
	AS (
		SELECT
			 [Материал]
			,[ЗКЗП: Дата поставки]
			,[Quantity in POs] = SUM([ЗКЗП: Количество]) - ISNULL(SUM([Qty in unit of entry]), 0)
		FROM
			dbo.BEX_Analytical_Report b
			LEFT JOIN (
						SELECT
							 Material
							,[Delivery Order]
							,[Position]
							,[Qty in unit of entry]	= SUM([Qty in unit of entry])
						FROM
							dbo.Materials_Movements_SAP4HANA_new
						WHERE
							[Movement type] IN (@MovementTypeIncome, @MovementTypeIncomeCancel)
							AND Plant IN (@StoragePlant, @SecondVirtualPlant)
						GROUP BY
							 Material
							,[Delivery Order]
							,[Position]
						) mv
				ON b.[ЗКЗ: Номер заказа на поставку] = mv.[Delivery Order]
				AND b.Материал = mv.Material
				AND b.[ЗКЗП: Позиция заказа на поставку] = mv.position
		WHERE 
				b.[ЗКЗ: Номер заказа на поставку] IS NOT NULL 
			AND b.[ЗКЗ: Номер заказа на поставку] NOT LIKE @PurchaseDocMove
			AND b.[ЗКЗП: Индикатор удаления] IS NULL 
			AND b.[ЗКЗП: Конечная поставка] = 0
		GROUP BY
			 [Материал]
			,[ЗКЗП: Дата поставки]
	),

manufacturers
	AS (
		SELECT 
			 Материал
			,[Название поставщика]
			,[Цена]
		FROM 
			(
			SELECT 
				 Материал
				,[Название поставщика]
				,[ЗКЗ: Номер заказа на поставку]
				,[ЗКЗ: Дата создания]
				,[ЗКЗП: Дата поставки]
				,[Цена]								= CASE 
															WHEN [ЗКЗ: Валюта] LIKE @CurrencyRub THEN [ЗКЗП:Стоимость] / [ЗКЗП: Количество]
															WHEN [ЗКЗ: Валюта] LIKE @CurrencyUsd THEN ([ЗКЗП:Стоимость] / [ЗКЗП: Количество]) * @UsdRubRate
															WHEN [ЗКЗ: Валюта] LIKE @CurrencyUsd THEN ([ЗКЗП:Стоимость] / [ЗКЗП: Количество]) * @EurRubRate
															WHEN [ЗКЗ: Валюта] LIKE @CurrencyFunt THEN ([ЗКЗП:Стоимость] / [ЗКЗП: Количество]) * @FuntRubRate
													   END
				,[Rank]								= ROW_NUMBER() OVER (PARTITION BY Материал ORDER BY [ЗКЗП: Дата поставки] DESC)
			FROM
				dbo.BEX_Analytical_Report
			WHERE 
				[ЗКЗ: Номер заказа на поставку] IS NOT NULL
				AND [ЗКЗП: Индикатор удаления] IS NULL 
				AND [Материал] IS NOT NULL
				AND [Название поставщика] IS NOT NULL
				AND [ЗКЗ: Номер заказа на поставку] IS NOT NULL
				AND [ЗКЗ: Номер заказа на поставку] LIKE @PurchaseDocPurchOrd
			) a
		WHERE
			[Rank] = 1
		),

stock	
	AS (
		SELECT 
			 [Material code]
			,[Batch]
			,[Free stock (Warehouse)]	= SUM(CASE WHEN Plant = @StoragePlant AND (WBS LIKE @FreeWBSTemplate OR WBS IS NULL) AND ([Special stock] <> @StockFlagSpecial OR [Special stock] IS NULL) THEN [Free stock] ELSE 0 END)
			,[Proj stock (Warehouse)]	= SUM(CASE WHEN Plant = @StoragePlant AND (WBS IS NOT NULL AND WBS NOT LIKE @FreeWBSTemplate) AND ([Special stock] <> @StockFlagSpecial OR [Special stock] IS NULL) THEN [Free stock] ELSE 0 END)
			,[Stock (Contractors)]		= SUM(CASE WHEN Plant = @ContractorsPlant AND ([Special stock] <> @StockFlagSpecial OR [Special stock] IS NULL) THEN [Free stock] ELSE 0 END)
			,[Stock (Refactor)]			= SUM(CASE WHEN Plant = @StoragePlant AND [Special stock] = @StockFlagSpecial THEN [Free stock] ELSE 0 END)
		FROM dbo.Warehouse_stocks_MB52
		WHERE [Import timestamp] = (SELECT MAX([Import timestamp]) FROM dbo.Warehouse_stocks_MB52)
		GROUP BY
			 [Material code]
			,Batch
	),

result
	AS (
		SELECT
		     t.[Type]
		      ,[Material Group] = mg.[User type]
			,l.[Material code]
			,s.[Batch]
			  ,[Batch description]	= CASE WHEN s.Batch LIKE 'SURPLUS%' THEN RIGHT(s.Batch, LEN(s.Batch) - 7) ELSE NULL END
		   ,ss.[Short Description]
		    ,m.Цена
		   ,ss.[Safety Stock]
			,l.[Safety stock (Entered)]
			,s.[Free stock (Warehouse)]
			,s.[Proj stock (Warehouse)]
			,s.[Stock (Contractors)]
			,s.[Stock (Refactor)]
			,p.[Quantity in POs]
			  ,[Delivery time] = STUFF((
								SELECT ' | '  + CAST(FORMAT(r.[ЗКЗП: Дата поставки], @DateFormat) as nvarchar(10)) + ' (' + CAST(CAST(r.[Quantity in POs] as float) as nvarchar(12)) + ') '
								FROM #POQty r
								WHERE r.Материал = l.[Material code] AND r.[Quantity in POs] > 0 
								ORDER BY r.[ЗКЗП: Дата поставки] ASC
								FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
			,m.[Название поставщика]
			  ,[Comment]			= ''
			  ,[Comment info]		= STUFF((
										SELECT ' | ' + uc.Comment + '(' + CAST(CAST(uc.[Created datetime] as date) as nvarchar(10)) + ', ' + uc.[User] + ')'
										FROM [TM3].[WSStockReport_UserComments] uc
										WHERE uc.[Material code] = l.[Material code] 
										AND ISNULL(uc.Batch, '') = ISNULL(s.Batch, '')
										AND uc.[Material type id] = l.[Type id]
										ORDER BY uc.[Created datetime] DESC
										FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
			,l.[Deleted]
		FROM
			dbo.TM3_WSStockReport_MasterMaterialsList l
			LEFT JOIN (
				SELECT 
					 Материал
					,[Quantity in POs] = SUM([Quantity in POs])
				FROM
					#POQty
				GROUP BY
					Материал
				) p
				ON l.[Material code] = p.Материал
			LEFT JOIN stock s
				ON l.[Material code] = s.[Material code]
			LEFT JOIN manufacturers m
				ON l.[Material code] = m.Материал
			LEFT JOIN dbo.Materials_Settings_new ss
				ON l.[Material code] = ss.[Material Code]
			LEFT JOIN dbo.TM3_WSStockReport_MasterMaterialTypes t
				ON l.[Type id] = t.Id
			LEFT JOIN dbo.[TM3_WSStockReport_UserMaterialGroup] mg
				ON l.[Material code] = mg.[Material code]
				AND l.[Type id] = mg.[Material type id]
		WHERE
			(@MaterialType IS NULL OR t.[Type] = @MaterialType)
	)
SELECT
	 [Type]
	,[Material Group]
	,[Material code]
	,[Batch]
	,[Batch description]
	,[Short Description]
	,Цена							= CAST([Цена] as decimal(18,2))	
	,[Safety stock (SAP)]			= ISNULL([Safety Stock], 0)
	,[Safety stock (Entered)]		= ISNULL([Safety stock (Entered)], 0)
	,[Stock]						= SUM(ISNULL([Free stock (Warehouse)], 0)) OVER (PARTITION BY [Material code], [Material Group]) + SUM(ISNULL([Proj stock (Warehouse)], 0)) OVER (PARTITION BY [Material code], [Material Group]) 
	,[Free stock (Warehouse)]		= SUM(ISNULL([Free stock (Warehouse)], 0)) OVER (PARTITION BY [Material code], [Material Group])
	,[Free stock (Warehouse) Batch]	= ISNULL([Free stock (Warehouse)], 0)
	,[Proj stock (Warehouse) Batch]	= ISNULL([Proj stock (Warehouse)], 0)
	,[Stock (Contractors)]			= ISNULL([Stock (Contractors)], 0)
	,[Stock (Refactor)]				= ISNULL([Stock (Refactor)], 0)
	,Ordered						= ISNULL([Quantity in POs], 0)
	,[Delivery time]				= [Delivery time]
	,Manufacturer					= [Название поставщика]
	,[Comment]
	,[Comment info]
	,[Deleted material]		= [Deleted]
FROM 
	result
order by 
	 [Deleted] ASC
	,[Type] DESC
	,[Material Group] ASC
	,[Material code] DESC
	,[Batch] ASC;


