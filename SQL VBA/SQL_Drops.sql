USE master;
GO
ALTER DATABASE Tempt SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
GO
DROP DATABASE Temp;
GO

--

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [Tempt].[dbo].[tempt_report](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Datum] [date] NOT NULL,
	[Zeit] [varchar](10) NOT NULL,
	[Benutzer] [varchar](50) NOT NULL,
	[Objekt] [varchar](10) NOT NULL,
	[Grundstueckskaufpreis] [money] NOT NULL,
	[Objektankaufskosten] [money] NOT NULL,
	[Baukosten] [money] NOT NULL,
	[Planerkosten] [money] NOT NULL,
	[Sicherheit] [money] NOT NULL,
	[Herstellkosten] [money] NOT NULL,
	[Vertriebskosten] [money] NOT NULL,
	[SonstigeKosten] [money] NOT NULL,
	[Gesamtkosten] [money] NOT NULL,
	[VerkaufspreisEinheiten] [money] NOT NULL,
	[VerkaufspreisTG] [money] NOT NULL,
	[Gesamterloes] [money] NOT NULL,
	[IRR] [float] NOT NULL,
	[ObjektReturn] [float] NOT NULL,
	[EKmax] [money] NOT NULL,
 CONSTRAINT [PK_Tempt_Report] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [tempt].[dbo].[tempt_report_test](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Datum] [date] NOT NULL,
	[Zeit] [varchar](10) NOT NULL,
	[Benutzer] [varchar](50) NOT NULL,
	[Objekt] [varchar](10) NOT NULL,
	[Grundstueckskaufpreis] [money] NOT NULL,
	[Objektankaufskosten] [money] NOT NULL,
	[Baukosten] [money] NOT NULL,
	[Planerkosten] [money] NOT NULL,
	[Sicherheit] [money] NOT NULL,
	[Herstellkosten] [money] NOT NULL,
	[Vertriebskosten] [money] NOT NULL,
	[SonstigeKosten] [money] NOT NULL,
	[Gesamtkosten] [money] NOT NULL,
	[VerkaufspreisEinheiten] [money] NOT NULL,
	[VerkaufspreisTG] [money] NOT NULL,
	[Gesamterloes] [money] NOT NULL,
	[IRR] [float] NOT NULL,
	[ObjektReturn] [float] NOT NULL,
	[EKmax] [money] NOT NULL,
 CONSTRAINT [PK_Tempt_Report_test] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO
