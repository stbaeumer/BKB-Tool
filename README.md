# BKB-Tool
```
                                                       ____    _  __  ____            _____                   _ 
                                                       | __ )  | |/ / | __ )          |_   _|   ___     ___   | |
                                                       |  _ \  | ' /  |  _ \   _____    | |    / _ \   / _ \  | |
                                                       | |_) | | . \  | |_) | |_____|   | |   | (_) | | (_) | | |
                                                       |____/  |_|\_\ |____/            |_|    \___/   \___/  |_|
                                                                                                                 
╭────────────────────────────────────────────────────── https://github.com/stbaeumer/BKB-Tool | GPLv2 | v1.0.51 ───────────────────────────────────────────────────────╮
│ 1952 Schüler*innen: 1933 aktive, 19 extern                                                                                                                           │
╰──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────╯
 1  Mailadressen           Mo, Mi, Fr      Fehlende schulinterne Mailadressen in den Individualdaten I ergänzen                                          
 2  Webuntis & Co.         Mo, Mi, Fr      Importdateien für Webuntis, Littera, Netman erstellen                                                         
 3  Fotos aus SchILD       Mo, Mi, Fr      Schüler*innenfotos aus SchILD für Webuntis, Geevoo und Netman bereitstellen                                   
 4  Absentismus            Mo              Klassenleitungen über schulpflichtverletzende Schüler*innen informieren                                       
 5  Klassenbuchpflege      Mo              Säumige Lehrer*innen auf fehlende Klassenbucheinträge hinweisen                                               
 6  Gruppen & Organigramm  Mo              Gruppen & Organigramm aus Untisanrechnungen und Unterrichten für Wiki-Import erstellen                        
 7  Outlook                Mo              CSV-Terminexporte für Wiki aufbereiten                                                                        
 8  Stammdaten             Mo              Stammdaten zwischen SchILD und Untis abgleichen                                                               
 9  Zeugnis #1 Unterricht  14 Tage vor ZK  Unterrichte und Kurse schülergenau von Webuntis nach SchILD übertragen                                        
10  Zeugnis #2 Belegung    13 Tage vor ZK  In den Klassen des Beruflichen Gymnasiums die Klausurbelegung aus Wiki in die SchILD-Leistungsdaten übernehmen
11  Zeugnis #3 Fehlzeiten  3  Tage vor ZK  Die Fehlzeiten werden in bestehende Lernabschnittsdaten eingefügt                                             
12  Zeugnis #4 Noten       1  Tag  vor ZK  Noten aus Webuntis (MarksPerLesson) nach SchILD (SchuelerLeistungsdaten) schreiben                            
13  Zeugnis #5 Listen      1  Tag  vor ZK  Notenlisten werden in Wiki veröffentlicht. Fehlende Noten werden markiert                                     
14  Wiki                                   Diverse SQLite-Dateien (Praktikum etc.) erstellen                                                             
15  Fotos (a)              1. Schultag     Schüler*innen klassenweise fotografieren, kopieren, umbenennen, ablegen                                       
16  Fotos (b)              1. Schultag     Schüler*innenfotos nach SchILD2 hochladen                                                                     
17  Klausurbelegung        1. Schultag     Wiki-Seite (mit Zuordnung der SuS zu allen Fächern) erstellen / auslesen                                      
18  Schnellmeldung         September       Relationsgruppen im September aufbereiten                                                                     
19  Statistik              September       Unterrichtsverteilung und Anrechnungen nach SchILD importieren                                                
20  Klassen anlegen        vor Sommer      Neue Klassen von Untis nach SchILD übergeben und Eigenschaften anpassen                                       
21  Altersermäßigung       vor Sommer      Altersermäßigung berechnen für 2025/2026 und 2026/2027                                                        
22  Sonderzeiten           Mai/Juni        Lehrkräftesonderzeiten (Anrechnungen) nach SchILD importieren                                                 
23  PDF verschlüsseln                      Von PDF-Dateien in /home/stefan/Downloads verschlüsselte Kopien erstellen                                     
24  PDF-Seiten mailen                      PDF-Seiten an darauf enthaltene E-Mail-Adressen mailen                                                        
25  PDF-Zeugnisse                          Auslesen und in die SchILD-Dokumentenverwaltung einsortieren                                                  
26  Zeugnis-Chat                           Säumige Lehrer*innen im Teams-Chat an die Noten-Eintragung erinnern                                           
27  Zeugnisse                              Leistungsdaten (Unterrichte, Zeugnisnoten, Fehlzeiten, ...) nach SchILD importieren                           
28  Teams-Chat                             Teams-Chat mit Gruppe von Lehrkräften beginnen                                                                
29  Teilleistungen                         SchuelerTeilleistungen.dat erstellen                                                                          
30  Sprechtag              Januar          Lehrerübersichtsseite im Wiki veröffentlichen                                                                 
31  SVWS-Server                            managen                                                                                                       
┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ Geben Sie eine Zahl ein oder e für Einstellungen oder h für Onlinehilfe                                                                                              │
└──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘
 Ihre Auswahl (25): 


```

<a name="funktionen"></a>
## Funktionen in BKB-Tool

<a name="webuntis"></a>
### Schüler*innen-Importdatei für Webuntis erstellen

<a name="littera"></a>
### Schüler*innen-Importdatei für Littera erstellen

<a name="netman"></a>
### Schüler*innen-Importdatei für Netman erstellen

<a name="mailadressen"></a>
### Schuelerzusatzdaten.dat werden um schulinterne Mailadressen ergänzt (falls leer)

Eine eindeutige Mailadresse in SchILD ist unbedingt erstrebenswert. Die Adresse wird dann von SchILD an Drittproramme (z.B. O365) übergeben.

Die Mailaddressen der Schüler*innen und Schüler können wahlweise manuell in den Individualdaten I eingetippt werden oder über die Datei Schuelerzusatzdaten.dat hochgalden werden. 



Die Mailadressen werden 

<a name="statistik"></a>
### Unterrichtsverteilung und Anrechnungen nach SchILD importieren
