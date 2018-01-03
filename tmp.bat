@echo off
echo Um Ihre Datenbankstruktur zu aktualisieren, muessen einige Kommandos
echo auf Ihre Datenbank angewendet werden.
echo -------------------------------------------------------------------
echo SICHERN SIE IHRE DATENBANK! (kein Scherz)
echo UND
echo lassen Sie diese Datei von Ihrem technischen Support ueberpruefen,
echo -------------------------------------------------------------------
echo !BEVOR! Sie sie durch Doppelklick vom Arbeisplatz aus starten.
echo   Agencyprof kann erst aktualisiert werden wenn diese Aenderungen
echo   durchgefuehrt wurden.
echo -------------------------------------------------------------------
echo Danach koennen Sie Agencyprof über die Verwaltungsfunktionen
echo aktualisieren.
echo -------------------------------------------------------------------
echo Druecken Sie Return um fortzufahren, oder
echo             Strg c um abzubrechen.
pause
echo -------------------------------------------------------------------
echo Aktualisierung laeuft ...
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `access_log` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `alarmliste` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `anreden` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `artikel` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `artikellieferant` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `auftritthigru` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `auftrittsfelder` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `auftrittstypen` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `aut_werke` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `b_loc` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `benutzerdaten` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `benutzergruppen` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `bkfirmen` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `bkurse` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `bplan` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `dictionary` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `dictionary_taboo` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `dochist` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `finanzen` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `gruppennamen` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `kurse` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `mailip` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `mailsafe` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `mysql.columns_priv` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `mysql.db` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `mysql.func` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `mysql.host` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `mysql.tables_priv` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `mysql.user` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `poplist` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `preisgruppen` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `q4dek` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `q4gms` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `sysvars` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `telefonbuch` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `todolist` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `tpwernoch` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `waehrung` ADD `tstamp` TIMESTAMP(14)"
d:\mysql\bin\mysql.exe -h 192.168.10.44 -u root -pokdrgz -D sks -e "ALTER TABLE `Warengruppe` ADD `tstamp` TIMESTAMP(14)"
