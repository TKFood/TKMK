-- 20261112 GROUPSALES 新增 公司行號 欄位 (frmGROUPSALESBYTA008 使用)
USE [TKMK];
GO

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
