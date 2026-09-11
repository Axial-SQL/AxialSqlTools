-- Run in a disposable test database. Existing objects are never dropped.
-- Select dbo-style schema AxialDataCompareTest in Compare Table Data after this script finishes.
IF SCHEMA_ID(N'AxialDataCompareTest') IS NOT NULL
    THROW 51100, 'The test schema already exists. Use another disposable database.', 1;
GO
CREATE SCHEMA AxialDataCompareTest;
GO
CREATE TABLE AxialDataCompareTest.SourceRows
(
    Id int IDENTITY(1,1) NOT NULL PRIMARY KEY,
    Label nvarchar(100) NULL,
    Amount decimal(38,10) NULL,
    Payload varbinary(max) NULL,
    EventDate datetime2(7) NULL,
    LegacyDate datetime NULL,
    LocalTime time(7) NULL,
    ZonedDate datetimeoffset(7) NULL,
    Document xml NULL,
    LabelLength AS LEN(Label),
    Version rowversion
);
CREATE TABLE AxialDataCompareTest.TargetRows
(
    Id int IDENTITY(1,1) NOT NULL PRIMARY KEY,
    Label nvarchar(100) NULL,
    Amount decimal(38,10) NULL,
    Payload varbinary(max) NULL,
    EventDate datetime2(7) NULL,
    LegacyDate datetime NULL,
    LocalTime time(7) NULL,
    ZonedDate datetimeoffset(7) NULL,
    Document xml NULL,
    LabelLength AS LEN(Label),
    Version rowversion
);
SET IDENTITY_INSERT AxialDataCompareTest.SourceRows ON;
INSERT AxialDataCompareTest.SourceRows
    (Id, Label, Amount, Payload, EventDate, LegacyDate, LocalTime, ZonedDate, Document)
VALUES
    (1, N'Identical', 1234567890123456789012345678.1234567890, 0x00FF01,
     '2026-09-11T12:34:56.1234567', '2026-09-11T12:34:56.003', '12:34:56.1234567',
     '2026-09-11T12:34:56.1234567-06:00', CONVERT(xml, N'<root> <x>1</x> </root>', 1)),
    (2, N'Updated Unicode: Москва', 42.1234567890, 0x010203,
     '2026-09-11T12:34:56.1234567', '2026-09-11T12:34:56.007', '12:34:56.1234567',
     '2026-09-11T12:34:56.1234567+05:30', N'<root a="new"/>'),
    (3, N'Only source', NULL, 0x, NULL, NULL, NULL, NULL, NULL);
SET IDENTITY_INSERT AxialDataCompareTest.SourceRows OFF;
SET IDENTITY_INSERT AxialDataCompareTest.TargetRows ON;
INSERT AxialDataCompareTest.TargetRows
    (Id, Label, Amount, Payload, EventDate, LegacyDate, LocalTime, ZonedDate, Document)
SELECT Id, Label, Amount, Payload, EventDate, LegacyDate, LocalTime, ZonedDate, Document
FROM AxialDataCompareTest.SourceRows WHERE Id IN (1,2);
UPDATE AxialDataCompareTest.TargetRows SET Label=N'Old value', Amount=41, Document=N'<root a="old"/>' WHERE Id=2;
INSERT AxialDataCompareTest.TargetRows (Id, Label) VALUES (4, N'Only target');
SET IDENTITY_INSERT AxialDataCompareTest.TargetRows OFF;

CREATE TABLE AxialDataCompareTest.DuplicateRows (Id int NULL, Label nvarchar(100) NULL);
INSERT AxialDataCompareTest.DuplicateRows VALUES (1, N'First'), (1, N'Duplicate');

-- Expected initial results: Different=1, OnlySource=1, OnlyTarget=1, Identical=1.
-- LabelLength and Version are excluded automatically.
-- Default synchronization: update Id=2, insert Id=3, retain Id=4.
-- Select the Only in target category for deletion, then apply: all three source rows match.
-- Conflict test: compare, then execute the following statement in another query before applying:
-- UPDATE AxialDataCompareTest.TargetRows SET Label=N'Concurrent change' WHERE Id=2;
-- Apply must stop and leave all selected actions uncommitted.
-- Duplicate test: compare DuplicateRows to SourceRows, include only Id/Label, and choose Id as
-- the custom key; the duplicate source key must stop comparison.
-- Cleanup, only after finishing the tests:
-- DROP TABLE AxialDataCompareTest.DuplicateRows;
-- DROP TABLE AxialDataCompareTest.TargetRows;
-- DROP TABLE AxialDataCompareTest.SourceRows;
-- DROP SCHEMA AxialDataCompareTest;
