-- 20261112 GROUPSALES 新增 公司行號 欄位 (frmGROUPSALESBYTA008 使用)
-- 比照 旅行社：GROUPSALES.TRAVELNAMES + 代碼表 GROUPTRAVELS
--             => GROUPSALES.COMPANYNAMES + 代碼表 GROUPCOMPANYS
USE [TKMK];
GO

-- 1. GROUPSALES 新增 COMPANYNAMES
IF COL_LENGTH('[dbo].[GROUPSALES]', 'COMPANYNAMES') IS NULL
BEGIN
    ALTER TABLE [dbo].[GROUPSALES] ADD [COMPANYNAMES] NVARCHAR(100) NULL;
END
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.extended_properties
    WHERE major_id = OBJECT_ID('[dbo].[GROUPSALES]')
      AND minor_id = COLUMNPROPERTY(OBJECT_ID('[dbo].[GROUPSALES]'), 'COMPANYNAMES', 'ColumnId')
      AND name = 'MS_Description'
)
BEGIN
    EXEC sys.sp_addextendedproperty
        @name = N'MS_Description', @value = N'公司行號',
        @level0type = N'SCHEMA', @level0name = N'dbo',
        @level1type = N'TABLE',  @level1name = N'GROUPSALES',
        @level2type = N'COLUMN', @level2name = N'COMPANYNAMES';
END
GO

-- 2. 公司行號 代碼表 (自動完成清單；畫面輸入新名稱時程式會自動 INSERT [NAMES])
IF OBJECT_ID('[dbo].[GROUPCOMPANYS]', 'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[GROUPCOMPANYS]
    (
        [ID]    INT IDENTITY(1,1) NOT NULL CONSTRAINT [PK_GROUPCOMPANYS] PRIMARY KEY,
        [NAMES] NVARCHAR(100) NOT NULL
    );
    CREATE UNIQUE INDEX [UX_GROUPCOMPANYS_NAMES] ON [dbo].[GROUPCOMPANYS]([NAMES]);

    EXEC sys.sp_addextendedproperty
        @name = N'MS_Description', @value = N'團務 公司行號',
        @level0type = N'SCHEMA', @level0name = N'dbo',
        @level1type = N'TABLE',  @level1name = N'GROUPCOMPANYS';
END
GO
