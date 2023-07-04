
if OBJECT_ID('importTest') is not null
  drop table importTest
go
create table importTest
(ID                    numeric(18, 0) identity
,Datum                 datetime       null -- Datum	
,Portfolioname         nvarchar(256)  null -- Portfolioname	
,Wahrung               nvarchar(32)   null -- Währung	
,KampagnenName         nvarchar(256)  null -- Kampagnen-Name	
,Anzeigengruppenname   nvarchar(256)  null -- Anzeigengruppenname	
,SKU	               nvarchar(256)  null -- Beworbene SKU	
,ASIN                  nvarchar(256)  null -- Beworbene ASIN	
,Impressionen          int            null -- Impressionen	
,Klicks                int            null -- Klicks	
,Klickrate             float          null -- Klickrate (CTR)	
,KlickCPC              money          null -- Kosten pro Klick (CPC)	
,Ausgaben              money          null -- Ausgaben	
,UmsatzGesamt          money          null -- 7 Tage, Umsatz gesamt (€)	
,ACOS                  float          null -- Gesamtumsatzkosten für Werbung (ACOS) 	
,ROAS                  float          null -- Gesamtrendite von Werbeausgaben (Return on Advertising Spend, ROAS)	
,AuftrageGesamt        int            null -- 7 Tage, Aufträge gesamt (#)	
,EinheitenGesamt       int            null -- 7 Tage, Einheiten gesamt (#)	
,Konversionsrate       float          null -- 7-Tage-Konversionsrate	
,BeworbeneSKUEinheiten int            null -- 7 Tage, Beworbene SKU-Einheiten (#)	
,AndereSKUEinheiten    int            null -- 7-Tage, Andere SKU-Einheiten (#)	
,BeworbeneSKUUmsatze   money          null -- 7 Tage, Beworbene SKU-Umsätze (€)	
,AndereSKUUmsatze      money          null -- 7-Tage, Andere SKU-Umsätze (€)
)
go
--create unique index rdb1 on importTest(Datum, Portfolioname, Wahrung, KampagnenName, Anzeigengruppenname, SKU, ASIN)
go
grant all on importTest to public




