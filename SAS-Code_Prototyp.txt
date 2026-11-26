/*************************************************************
 Prototyp Pumpwerk GmbH – Schritt 1: Import der drei Quellen
 und erste Prüfung der Datenqualität
*************************************************************/

/* Pfad zum Ordner mit den drei Quelldateien */
%let pfad = ~/sasuser.v94;

/* ---------- Quelle 1: historische Fertigungsdaten (Excel) ---------- */
proc import datafile="&pfad/fertigung_historisch.xlsx"
            out=work.fertigung dbms=xlsx replace;
   sheet="Fertigung";
   getnames=yes;
run;

/* ---------- Quelle 2: aktuelle Sensordaten (CSV, Semikolon) ---------- */
/* Einlesen per DATA-Step, damit Trennzeichen und deutsches
   Datumsformat kontrolliert verarbeitet werden. */
data work.sensor_roh;
   infile "&pfad/sensordaten_oktober2026.csv"
          dlm=';' dsd firstobs=2 truncover;
   length Geraete_ID $10;
   input Geraete_ID $
         Datum :ddmmyy10.
         Betriebsstunden
         Auslastung_Prozent
         Drehzahl_U_min
         Druck_bar
         Temperatur_C
         Vibration_mm_s;
   format Datum ddmmyy10.;
run;

/* ---------- Quelle 3: Servicemeldungen (CSV mit Freitext) ---------- */
proc import datafile="&pfad/servicemeldungen.csv"
            out=work.service dbms=csv replace;
   getnames=yes;
   guessingrows=max;
run;

/* ---------- Erste Sichtprüfung: je 5 Zeilen ---------- */
title "Quelle 1: Fertigung (historisch)";
proc print data=work.fertigung(obs=5); run;
title "Quelle 2: Sensordaten (aktuell)";
proc print data=work.sensor_roh(obs=5); run;
title "Quelle 3: Servicemeldungen";
proc print data=work.service(obs=5); run;

/* ---------- Datenqualität der Sensordaten ---------- */
title "Datenqualität Sensordaten: Anzahl, fehlende Werte, Wertebereich";
proc means data=work.sensor_roh n nmiss min max mean maxdec=2;
   var Betriebsstunden Auslastung_Prozent Drehzahl_U_min
       Druck_bar Temperatur_C Vibration_mm_s;
run;
title;


/*************************************************************
 Prototyp Pumpwerk GmbH – Schritt 2: Bereinigung und
 Zusammenführung historischer und aktueller Daten
*************************************************************/

/* ---------- 2a Fertigung: Excel-Datum in SAS-Datum umrechnen ----------
   Excel zählt Tage ab 1900, SAS ab 1.1.1960. Differenz = 21916 Tage. */
data work.fertigung_clean;
   set work.fertigung;
   Fertigungsdatum    = Fertigungsdatum    - 21916;
   Auslieferungsdatum = Auslieferungsdatum - 21916;
   format Fertigungsdatum Auslieferungsdatum ddmmyy10.;
run;

title "Fertigung nach Datumsumrechnung";
proc print data=work.fertigung_clean(obs=5);
   var SerienNr Baureihe Fertigungsdatum Auslieferungsdatum Kundenbranche;
run;

/* ---------- 2b Sensordaten bereinigen ---------- */
/* Dubletten (gleiche Pumpe, gleicher Tag) entfernen */
proc sort data=work.sensor_roh out=work.sensor_sort
          nodupkey dupout=work.dubletten;
   by Geraete_ID Datum;
run;

/* Fehlercode 999.9 als fehlenden Wert kennzeichnen */
data work.sensor_clean;
   set work.sensor_sort;
   if Temperatur_C = 999.9 then Temperatur_C = .;
run;

title "Entfernte Dubletten";
proc print data=work.dubletten; run;

title "Datenqualität Sensordaten nach Bereinigung";
proc means data=work.sensor_clean n nmiss min max mean maxdec=2;
   var Betriebsstunden Auslastung_Prozent Drehzahl_U_min
       Druck_bar Temperatur_C Vibration_mm_s;
run;

/* ---------- 2c Verdichten und Zusammenführen ---------- */
proc sql;
   /* aktuelle Sensordaten: ein Datensatz je Pumpe */
   create table work.sensor_agg as
   select Geraete_ID               as SerienNr,
          count(*)                 as Anzahl_Tage,
          mean(Betriebsstunden)    as BStd_Tag    format=8.1,
          mean(Auslastung_Prozent) as Auslastung  format=8.1,
          mean(Temperatur_C)       as Temp_Mittel format=8.1,
          max(Temperatur_C)        as Temp_Max    format=8.1,
          mean(Vibration_mm_s)     as Vib_Mittel  format=8.2,
          max(Vibration_mm_s)      as Vib_Max     format=8.2
   from work.sensor_clean
   group by Geraete_ID;

   /* Servicemeldungen: Anzahl je Pumpe, plus Stichwortsuche im Freitext */
   create table work.service_agg as
   select SerienNr,
          count(*) as Anzahl_Meldungen,
          sum( find(Beschreibung, 'Vibration', 'i') > 0
            or find(Beschreibung, 'Geräusch',  'i') > 0
            or find(Beschreibung, 'heiß',      'i') > 0
            or find(Beschreibung, 'Übertemperatur', 'i') > 0 )
                   as Hinweis_Lager_Hitze,
          /* Eingangsdatum liegt als Text (JJJJ-MM-TT) vor -> in SAS-Datum umwandeln */
          max(input(Eingangsdatum, yymmdd10.)) as Letzte_Meldung format=ddmmyy10.
   from work.service
   group by SerienNr;

   /* Analysetabelle: historisch + aktuell + Service */
   create table work.analyse as
   select f.*,
          s.Anzahl_Tage, s.BStd_Tag, s.Auslastung,
          s.Temp_Mittel, s.Temp_Max, s.Vib_Mittel, s.Vib_Max,
          coalesce(v.Anzahl_Meldungen, 0)    as Anzahl_Meldungen,
          coalesce(v.Hinweis_Lager_Hitze, 0) as Hinweis_Lager_Hitze,
          v.Letzte_Meldung
   from work.fertigung_clean as f
        inner join work.sensor_agg  as s on f.SerienNr = s.SerienNr
        left  join work.service_agg as v on f.SerienNr = v.SerienNr
   order by f.SerienNr;
quit;

title "Zusammengeführte Analysetabelle (Auszug)";
proc print data=work.analyse(obs=10);
   var SerienNr Baureihe Kundenbranche Fertigungsdatum
       BStd_Tag Auslastung Temp_Mittel Vib_Mittel Vib_Max Anzahl_Meldungen;
run;
title;


/*************************************************************
 Prototyp Pumpwerk GmbH – Schritt 3: Analysen und Grafiken
*************************************************************/

/* ===========================================================
   3.1 Ausreißererkennung: 1,5 x IQR-Regel je Baureihe
       auf der maximalen Vibration im Oktober 2026
   =========================================================== */
proc means data=work.analyse noprint nway;
   class Baureihe;
   var Vib_Max;
   output out=work.grenzen(drop=_type_ _freq_) q1=Q1 q3=Q3;
run;

data work.grenzen;
   set work.grenzen;
   IQR        = Q3 - Q1;
   Obergrenze = Q3 + 1.5*IQR;
   Untergrenze= Q1 - 1.5*IQR;
   format Q1 Q3 IQR Obergrenze Untergrenze 8.2;
run;

proc sql;
   create table work.analyse_flag as
   select a.*, g.Obergrenze,
          case when a.Vib_Max > g.Obergrenze or a.Vib_Max < g.Untergrenze
               then 1 else 0 end as Ausreisser
   from work.analyse as a
        inner join work.grenzen as g on a.Baureihe = g.Baureihe
   order by a.SerienNr;
quit;

title "Grenzwerte der Ausreißererkennung je Baureihe (1,5 x IQR)";
proc print data=work.grenzen noobs; run;

title "Erkannte Ausreißer";
proc print data=work.analyse_flag noobs;
   where Ausreisser = 1;
   var SerienNr Baureihe Kundenbranche Vib_Mittel Vib_Max Obergrenze
       Temp_Max Hinweis_Lager_Hitze Letzte_Meldung;
run;

title "Abgleich: Ausreißer vs. Servicemeldungen mit Hinweis auf Lager/Hitze";
proc freq data=work.analyse_flag;
   tables Ausreisser*Hinweis_Lager_Hitze / norow nocol nopercent;
run;

/* ===========================================================
   3.2 Clusteranalyse (k-Means) auf dem Nutzungsverhalten
   =========================================================== */
/* Nur die beiden Nutzungsmerkmale clustern: Betriebsstunden pro Tag
   und Auslastung. Temperatur und Vibration hängen stark von der Baureihe
   (Pumpengröße) ab und würden sonst die Baureihen statt der
   Nutzungsprofile trennen (Ergebnis eines ersten Laufs mit 4 Merkmalen).
   Kopien anlegen und standardisieren (Mittelwert 0, Std.abw. 1), damit
   Stunden (0-24) und Prozent (0-100) gleich gewichtet sind. */
data work.cl_input;
   set work.analyse;
   z_BStd = BStd_Tag;
   z_Ausl = Auslastung;
run;

proc stdize data=work.cl_input out=work.cl_std method=std;
   var z_BStd z_Ausl;
run;

/* Wahl der Clusterzahl: Ellbogen-Kriterium für k = 2 bis 6 */
%macro ellbogen;
   %do k = 2 %to 6;
      proc fastclus data=work.cl_std maxclusters=&k maxiter=100
                    out=work.tmp noprint;
         var z_BStd z_Ausl;
      run;
      proc sql;
         create table work.e&k as
         select &k as k, sum(distance**2) as Fehlerquadratsumme
         from work.tmp;
      quit;
   %end;
   data work.ellbogen;
      set work.e2-work.e6;
   run;
%mend;
%ellbogen

title "Ellbogen-Kriterium zur Wahl der Clusterzahl";
proc sgplot data=work.ellbogen;
   series x=k y=Fehlerquadratsumme / markers;
   xaxis label="Anzahl Cluster k" integer;
   yaxis label="Fehlerquadratsumme innerhalb der Cluster";
run;

/* Finale Clusteranalyse mit k = 3 */
title "k-Means-Clusteranalyse (k = 3)";
proc fastclus data=work.cl_std maxclusters=3 maxiter=100 out=work.cluster;
   var z_BStd z_Ausl;
run;

title "Clusterprofile (Mittelwerte in Originaleinheiten)";
proc means data=work.cluster mean maxdec=1;
   class CLUSTER;
   var BStd_Tag Auslastung Temp_Mittel Vib_Mittel;
run;

title "Cluster nach Kundenbranche und Baureihe";
proc freq data=work.cluster;
   tables CLUSTER*Kundenbranche CLUSTER*Baureihe / norow nocol nopercent;
run;

/* ===========================================================
   3.3 Grafiken
   =========================================================== */
/* Boxplot: maximale Vibration je Baureihe, Ausreißer beschriftet */
title "Maximale Vibration im Oktober 2026 je Baureihe";
proc sgplot data=work.analyse_flag;
   vbox Vib_Max / category=Baureihe datalabel=SerienNr;
   xaxis label="Baureihe";
   yaxis label="maximale Vibration (mm/s)";
run;

/* Vergleich: mittlere Vibration – hier fallen die Pumpen kaum auf */
title "Mittlere Vibration im Oktober 2026 je Baureihe";
proc sgplot data=work.analyse_flag;
   vbox Vib_Mittel / category=Baureihe datalabel=SerienNr;
   xaxis label="Baureihe";
   yaxis label="mittlere Vibration (mm/s)";
run;

/* Streudiagramm: Nutzungsprofile nach Cluster */
title "Nutzungsprofile: Betriebsstunden und Auslastung nach Cluster";
proc sgplot data=work.cluster;
   scatter x=BStd_Tag y=Auslastung / group=CLUSTER
           markerattrs=(symbol=circlefilled size=8);
   xaxis label="mittlere Betriebsstunden pro Tag";
   yaxis label="mittlere Auslastung (%)";
run;

/* Verlaufsdiagramm: tägliche Vibration der Ausreißer-Pumpen */
proc sql;
   create table work.verlauf as
   select s.Geraete_ID, s.Datum, s.Vibration_mm_s
   from work.sensor_clean as s
   where s.Geraete_ID in (select SerienNr from work.analyse_flag
                          where Ausreisser = 1)
   order by s.Geraete_ID, s.Datum;
quit;

title "Tägliche Vibration der auffälligen Pumpen im Oktober 2026";
proc sgplot data=work.verlauf;
   series x=Datum y=Vibration_mm_s / group=Geraete_ID markers;
   xaxis label="Datum";
   yaxis label="Vibration (mm/s)";
run;
title;
