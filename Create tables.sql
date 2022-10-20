
CREATE SCHEMA [TM3]
GO

CREATE TABLE [dbo].[TM3_WSStockReport_UserMaterialGroup](
	[RowID] [bigint] IDENTITY(1,1) NOT NULL,
	[Material code] [bigint] NOT NULL,
	[Material type id] [int] NOT NULL,
	[User type] [nvarchar](120) NULL,
 CONSTRAINT [PK_TM3_WSStockReport_UserMaterialGroup_MaterialCode] PRIMARY KEY CLUSTERED 
(
	[Material code] ASC,
	[Material type id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO


CREATE TABLE [dbo].[TM3_WSStockReport_UserComments](
	[RowID] [bigint] IDENTITY(1,1) NOT NULL,
	[Material code] [bigint] NOT NULL,
	[Comment] [nvarchar](max) NOT NULL,
	[User] [nvarchar](120) NOT NULL,
	[Created datetime] [datetime2](7) NOT NULL,
	[Batch] [nvarchar](120) NULL,
	[Material type id] [int] NOT NULL,
 CONSTRAINT [PK_TM3_WSStockReport_UserComments_MaterialcodeCreateddatetime] PRIMARY KEY CLUSTERED 
(
	[Material code] ASC,
	[Material type id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

ALTER TABLE [dbo].[TM3_WSStockReport_UserComments] ADD  DEFAULT (user_name()) FOR [User]
GO

ALTER TABLE [dbo].[TM3_WSStockReport_UserComments] ADD  DEFAULT (getdate()) FOR [Created datetime]
GO

CREATE TABLE [dbo].[TM3_WSStockReport_MasterMaterialTypes](
	[Id] [bigint] IDENTITY(1,1) NOT NULL,
	[Type] [nvarchar](120) NULL,
	[Department]  AS (case when [Type] like '%WE%' then 'WE' when [Type] like '%CWI%' then 'CWI'  end),
 CONSTRAINT [PK_TM3_WSStockReport_MasterMaterialsTypes_ID] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO


CREATE TABLE [dbo].[TM3_WSStockReport_MasterMaterialsList](
	[Material code] [bigint] NOT NULL,
	[Type id] [int] NOT NULL,
	[Deleted] [bit] NULL,
	[Safety stock (Entered)] [decimal](18, 3) NULL,
 CONSTRAINT [PK_TM3_WSStockReport_MasterMaterialsTypes_MaterialCode] PRIMARY KEY CLUSTERED 
(
	[Material code] ASC,
	[Type id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[TM3_WSStockReport_MasterMaterialsList] ADD  DEFAULT ((0)) FOR [Deleted]
GO

CREATE TABLE [TM3].[WSStockReport_UserComments](
	[RowID] [bigint] IDENTITY(1,1) NOT NULL,
	[Material code] [bigint] NOT NULL,
	[Comment] [nvarchar](max) NOT NULL,
	[User] [nvarchar](120) NOT NULL,
	[Created datetime] [datetime2](7) NOT NULL,
	[Batch] [nvarchar](120) NULL,
	[Material type id] [int] NOT NULL,
 CONSTRAINT [PK_WSStockReport_UserComments_MaterialcodeCreatedDT] PRIMARY KEY CLUSTERED 
(
	[RowID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

ALTER TABLE [TM3].[WSStockReport_UserComments] ADD  DEFAULT (user_name()) FOR [User]
GO

ALTER TABLE [TM3].[WSStockReport_UserComments] ADD  DEFAULT (getdate()) FOR [Created datetime]
GO


SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[BEX_Analytical_Report](
	[RowID] [int] IDENTITY(1,1) NOT NULL,
	[���: ����� ������] [bigint] NULL,
	[���: ���-�������] [nvarchar](120) NULL,
	[���: ����� ���-��������] [nvarchar](max) NULL,
	[��������] [bigint] NULL,
	[���: ������������ �������� ���] [nvarchar](40) NULL,
	[���: ���� �������������� �������� ���] [date] NULL,
	[���: ���� �������������� ������� ���� ������] [date] NULL,
	[���: ������������ ������������ ���� ������] [nvarchar](40) NULL,
	[���: ���� �������������� ������� ���� ��� ����] [date] NULL,
	[���������] [nvarchar](20) NULL,
	[�������� ����������] [nvarchar](max) NULL,
	[���: ����� ������ �� ��������] [bigint] NULL,
	[����: ���� ��������] [date] NULL,
	[���: ���� �����������] [date] NULL,
	[���: ������������ ��������� ������] [nvarchar](40) NULL,
	[���: ���� �������������� ��������] [date] NULL,
	[���: ������������ ��������] [nvarchar](40) NULL,
	[���: ����������] [decimal](18, 3) NULL,
	[���: ���������] [decimal](18, 2) NULL,
	[����: ����������] [decimal](18, 3) NULL,
	[����:���������] [decimal](18, 2) NULL,
	[���: ����������] [decimal](18, 3) NULL,
	[���: ���������] [decimal](18, 2) NULL,
	[���: �����] [nvarchar](4) NULL,
	[Import timestamp] [datetime2](7) NULL,
	[���: ������� ������] [int] NULL,
	[����: ������� ������ �� ��������] [int] NULL,
	[���: ��������� ��������] [bit] NULL,
	[����: ��������� ��������] [nvarchar](1) NULL,
	[���: ������� �������] [int] NULL,
	[���: ��� ���������� ������] [nvarchar](1) NULL,
	[���: ���� ��������] [date] NULL,
	[���: ������] [nvarchar](20) NULL,
	[����: ����] [nvarchar](255) NULL,
	[����: ���-�������] [nvarchar](120) NULL,
	[���: ��������� �������������� ������] [nvarchar](1) NULL,
	[����: �����] [nvarchar](30) NULL,
	[���: ��� ���������] [nvarchar](3) NULL,
	[���: ������] [nvarchar](20) NULL,
	[���: ���� ��������] [date] NULL,
	[���: ������������ ��������] [nvarchar](40) NULL,
	[���: ���� �������������� ��������] [date] NULL,
	[���: ���� �������������� ������������] [date] NULL,
	[���: ������ ���������] [bit] NULL,
	[����: �������� ��������] [bit] NULL,
	[���: �����] [nvarchar](15) NULL,
	[���: ��� ���������] [nvarchar](10) NULL,
	[����: �����] [nvarchar](4) NULL,
	[���: �����] [nvarchar](3) NULL,
	[�����# 2] [nvarchar](max) NULL,
	[���: �����] [nvarchar](4) NULL,
	[����: �����] [nvarchar](4) NULL,
 CONSTRAINT [PK__BEX_Analytical_Report_RowID] PRIMARY KEY CLUSTERED 
(
	[RowID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

ALTER TABLE [dbo].[BEX_Analytical_Report] ADD  DEFAULT (getdate()) FOR [Import timestamp]
GO

ALTER TABLE [dbo].[BEX_Analytical_Report] ADD  DEFAULT ((0)) FOR [���: ��������� ��������]
GO

CREATE TABLE [dbo].[Materials_Settings_new](
	[RowID] [bigint] IDENTITY(1,1) NOT NULL,
	[Material Code] [bigint] NULL,
	[3d] [nvarchar](4) NULL,
	[Short Description] [nvarchar](40) NULL,
	[Pl] [nvarchar](10) NULL,
	[Safety Stock] [decimal](18, 3) NULL,
	[Material Type] [nvarchar](2) NULL,
	[Order Point] [decimal](18, 3) NULL,
	[Max Level] [decimal](18, 3) NULL,
	[PP] [nvarchar](10) NULL,
	[Lead Time] [decimal](18, 2) NULL,
	[Request Processing Time] [decimal](18, 2) NULL,
	[Material Group] [nvarchar](10) NULL,
	[Unit of entry] [nvarchar](10) NULL,
	[Criticality] [bit] NOT NULL,
	[Long Description] [nvarchar](max) NULL,
	[Unit of entry storaging] [nvarchar](10) NULL,
	[Unit of entry storaging alt] [nvarchar](10) NULL,
	[ID sap accounting status] [smallint] NULL,
	[Import] [bit] NULL,
	[Manufacturer code] [nvarchar](20) NULL,
	[Material status plant] [nvarchar](2) NULL,
	[Material status all plants] [nvarchar](2) NULL,
	[Material code replacer] [bigint] NULL,
	[Material code old] [bigint] NULL,
	[Batch size calculation type] [nvarchar](2) NULL,
	[Moved during migration] [tinyint] NULL,
PRIMARY KEY CLUSTERED 
(
	[RowID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

ALTER TABLE [dbo].[Materials_Settings_new] ADD  DEFAULT ((0)) FOR [Criticality]
GO