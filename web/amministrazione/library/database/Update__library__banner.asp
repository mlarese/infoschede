<%
'...........................................................................................
'........................................................................................... 
'libreria di funzioni che contiene tutti gli aggiornamenti per il NEXT-banner
'...........................................................................................
'...........................................................................................


'*******************************************************************************************
'INSTALLAZIONE NEXT-BANNER
'...........................................................................................
function Install__NEXTBANNER(conn)
	Select case DB_Type(conn)
		case DB_Access
			Install__NEXTBANNER = _
					""
		case DB_SQL
			Install__NEXTBANNER = _
				"CREATE TABLE [dbo].[tb_applicativi](" + vbCrLf + _
				"	[sito_id] [int] IDENTITY(1,1) NOT NULL," + vbCrLf + _
				"	[sito_nome] [nvarchar](250) NULL," + vbCrLf + _
				"	[sito_url] [nvarchar](250) NULL," + vbCrLf + _
				" CONSTRAINT [PK_tb_applicativi] PRIMARY KEY CLUSTERED " + vbCrLf + _
				"(" + vbCrLf + _
				"	[sito_id] ASC " + vbCrLf + _
				") ON [PRIMARY] " + vbCrLf + _
				") ON [PRIMARY] " + vbCrLf + _
				"CREATE TABLE [dbo].[tb_tipiBanner]( " + vbCrLf + _
					"	[tipoB_id] [int] IDENTITY(1,1) NOT NULL, " + vbCrLf + _
					"	[tipoB_nome] [nvarchar](50) NULL, " + vbCrLf + _
					"	[tipoB_note] [ntext] NULL, " + vbCrLf + _
					" CONSTRAINT [PK_tb_tipiBanner] PRIMARY KEY CLUSTERED  " + vbCrLf + _
					"( " + vbCrLf + _
					"	[tipoB_id] ASC " + vbCrLf + _
					") ON [PRIMARY] " + vbCrLf + _
					") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]; " + vbCrLf + _
					"CREATE TABLE [dbo].[tb_banner]( " + vbCrLf + _
					"	[ban_id] [int] IDENTITY(1,1) NOT NULL, " + vbCrLf + _
					"	[ban_nome] [nvarchar](50) NULL, " + vbCrLf + _
					"	[ban_image] [nvarchar](50) NULL, " + vbCrLf + _
					"	[ban_link] [nvarchar](250) NULL, " + vbCrLf + _
					"	[ban_alt] [ntext] NULL, " + vbCrLf + _
					"	[ban_tipo] [int] NOT NULL, " + vbCrLf + _
					"	[ban_az] [int] NOT NULL, " + vbCrLf + _
					"	[ban_param] [nvarchar](100) NULL, " + vbCrLf + _
					"	[ban_value] [nvarchar](250) NULL, " + vbCrLf + _
					" CONSTRAINT [PK_tb_banner] PRIMARY KEY CLUSTERED " + vbCrLf + _ 
					"( " + vbCrLf + _
					"	[ban_id] ASC " + vbCrLf + _
					") ON [PRIMARY] " + vbCrLf + _
					") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]; " + vbCrLf + _
					"CREATE TABLE [dbo].[rel_banner_pagine]( " + vbCrLf + _
					"	[rbp_id] [int] IDENTITY(1,1) NOT NULL, " + vbCrLf + _
					"	[rbp_impress_iniz] [int] NULL, " + vbCrLf + _
					"	[rbp_impress] [int] NULL, " + vbCrLf + _
					"	[rbp_data_iniz] [smalldatetime] NULL, " + vbCrLf + _
					"	[rbp_data_fine] [smalldatetime] NULL, " + vbCrLf + _
					"	[rbp_click_iniz] [int] NULL, " + vbCrLf + _
					"	[rbp_click] [int] NULL, " + vbCrLf + _
					"	[rbp_pag] [int] NOT NULL, " + vbCrLf + _
					"	[rbp_banner] [int] NOT NULL, " + vbCrLf + _
					" CONSTRAINT [PK_rel_banner_pagine] PRIMARY KEY CLUSTERED  " + vbCrLf + _
					"( " + vbCrLf + _
					"	[rbp_id] ASC " + vbCrLf + _
					") ON [PRIMARY] " + vbCrLf + _
					") ON [PRIMARY]; " + vbCrLf + _
					"CREATE TABLE [dbo].[tb_pagine]( " + vbCrLf + _
					"	[pag_id] [int] IDENTITY(1,1) NOT NULL, " + vbCrLf + _
					"	[pag_url] [nvarchar](250) NULL, " + vbCrLf + _
					"	[pag_cat] [nvarchar](250) NULL, " + vbCrLf + _
					"	[pag_sito] [int] NOT NULL, " + vbCrLf + _
					" CONSTRAINT [PK_tb_pagine] PRIMARY KEY CLUSTERED  " + vbCrLf + _
					"( " + vbCrLf + _
					"	[pag_id] ASC " + vbCrLf + _
					") ON [PRIMARY] " + vbCrLf + _
					") ON [PRIMARY]; " + vbCrLf + _
					"ALTER TABLE [dbo].[tb_banner]  WITH CHECK ADD  CONSTRAINT [FK_tb_banner_tb_Indirizzario] FOREIGN KEY([ban_az]) " + vbCrLf + _
					"REFERENCES [dbo].[tb_Indirizzario] ([IDElencoIndirizzi]) " + vbCrLf + _
					"ON UPDATE CASCADE " + vbCrLf + _
					"ON DELETE CASCADE; " + vbCrLf + _
					"ALTER TABLE [dbo].[tb_banner]  WITH CHECK ADD  CONSTRAINT [FK_tb_banner_tb_tipiBanner] FOREIGN KEY([ban_tipo]) " + vbCrLf + _
					"REFERENCES [dbo].[tb_tipiBanner] ([tipoB_id]) " + vbCrLf + _
					"ON UPDATE CASCADE " + vbCrLf + _
					"ON DELETE CASCADE; " + vbCrLf + _
					"ALTER TABLE [dbo].[rel_banner_pagine]  WITH CHECK ADD  CONSTRAINT [FK_rel_banner_pagine_tb_banner] FOREIGN KEY([rbp_banner]) " + vbCrLf + _
					"REFERENCES [dbo].[tb_banner] ([ban_id]) " + vbCrLf + _
					"ON UPDATE CASCADE " + vbCrLf + _
					"ON DELETE CASCADE; " + vbCrLf + _
					"ALTER TABLE [dbo].[rel_banner_pagine]  WITH CHECK ADD  CONSTRAINT [FK_rel_banner_pagine_tb_pagine] FOREIGN KEY([rbp_pag]) " + vbCrLf + _
					"REFERENCES [dbo].[tb_pagine] ([pag_id]) " + vbCrLf + _
					"ON UPDATE CASCADE " + vbCrLf + _
					"ON DELETE CASCADE; " + vbCrLf + _
					"ALTER TABLE [dbo].[tb_pagine]  WITH CHECK ADD  CONSTRAINT [FK_tb_pagine_tb_applicativi] FOREIGN KEY([pag_sito]) " + vbCrLf + _
					"REFERENCES [dbo].[tb_applicativi] ([sito_id]) " + vbCrLf + _
					"ON UPDATE CASCADE " + vbCrLf + _
					"ON DELETE CASCADE "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'DISINSTALLAZIONE NEXT-BANNER
'...........................................................................................
function Uninstall__NEXTBANNER(conn)
	Select case DB_Type(conn)
		case DB_Access
			Uninstall__NEXTBANNER = _
					""
		case DB_SQL
			Uninstall__NEXTBANNER = _
                " ALTER TABLE tb_pagine DROP CONSTRAINT FK_tb_pagine_tb_applicativi; " + _
                " ALTER TABLE rel_banner_pagine DROP CONSTRAINT FK_rel_banner_pagine_tb_pagine; " + _
                " ALTER TABLE rel_banner_pagine DROP CONSTRAINT FK_rel_banner_pagine_tb_banner; " + _
                " ALTER TABLE tb_banner DROP CONSTRAINT FK_tb_banner_tb_tipiBanner; " + _
                " ALTER TABLE tb_banner DROP CONSTRAINT FK_tb_banner_tb_Indirizzario; " + _
                DropObject(conn, "tb_pagine", "TABLE") + _
                DropObject(conn, "rel_banner_pagine", "TABLE") + _
                DropObject(conn, "tb_banner", "TABLE") + _
                DropObject(conn, "tb_tipiBanner", "TABLE") + _
                DropObject(conn, "tb_applicativi", "TABLE")
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-BANNER  1
'...........................................................................................
'
'...........................................................................................
function Aggiornamento__NEXTBANNER__1(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__INFO__1 = ""
		case DB_SQL
			Aggiornamento__INFO__1 = ""
	end select
end function
'*******************************************************************************************

%>