Stored procedure for Microsoft SQL Server that generates INSERT statements for rows from a table.

The INSERT statements can then be copied to another database to insert the data into a similarly structured table.

It will also allow inserts in to IDENTITY columns.


Example:

exec cms_ScriptTable 'documentTypes', 'id BETWEEN 2 and 7', 1, 1

Output:

SET IDENTITY_INSERT [documentTypes] ON
INSERT INTO [documentTypes] ([id], [label], [iconURL], [extension], [mimeType], [isActive]) VALUES (2, N'Word Document', N'/_assets/document-icons/word-document.png', N'.doc', N'application/msword', 1)
INSERT INTO [documentTypes] ([id], [label], [iconURL], [extension], [mimeType], [isActive]) VALUES (3, N'Excel Document', N'/_assets/document-icons/excel-document.png', N'.xsl', N'application/vnd.ms-excel', 1)
INSERT INTO [documentTypes] ([id], [label], [iconURL], [extension], [mimeType], [isActive]) VALUES (4, N'PowerPoint (Open XML)', N'', N'.pptx', N'application/vnd.openxmlformats-officedocument.presentationml.presentation', 1)
INSERT INTO [documentTypes] ([id], [label], [iconURL], [extension], [mimeType], [isActive]) VALUES (5, N'Plain Text', N'', N'.txt', N'text/plain', 1)
INSERT INTO [documentTypes] ([id], [label], [iconURL], [extension], [mimeType], [isActive]) VALUES (6, N'Word Document (Open XML)', N'/_assets/document-icons/word-document.png', N'.docx', N'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 1)
INSERT INTO [documentTypes] ([id], [label], [iconURL], [extension], [mimeType], [isActive]) VALUES (7, N'Excel Document (Open XML)', N'/_assets/document-icons/excel-document.png', N'.xslx', N'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 1)
SET IDENTITY_INSERT [documentTypes] OFF
