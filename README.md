# CMon

CMon ist ein einfacher, kleiner Anrufmonitor für die FRITZ!Box-Familie, der als Symbol im Infobereich der Windows-Taskleiste läuft.

## Wichtige Vorabbemerkungen

**HINWEIS:** Dieses Projekt ist Teil einer Reihe von Veröffentlichungen sehr alter Projekte, die ich bisher nur in Form fertig kompilierter Windows-EXE-Tools als Freeware auf fabianswebworld.de veröffentlicht hatte. Es handelt sich dabei um teilweise über 20 Jahre alten Visual-Basic-6-Code. Die Veröffentlichung erfolgt hier ausschließlich zu Bastel- und Inspirationszwecken für Interessierte, da ich im Laufe der letzten 10 Jahre bereits mehrfach Nachfragen hierzu erhalten habe.
Der Code ist - zumindest in Teilen - weder schön noch elegant, aber auch heute noch zumindest funktional. Alle hier veröffentlichten Quellcodes wurden zumindest einmal unter Windows 11 erfolgreich ans Laufen gebracht. Es gilt aber: Der Code wird ohne jeglichen Anspruch auf Funktionalität, Sinnhaftigkeit oder Verständlichkeit veröffentlicht, und ich gebe auch keinen Support beim Nutzen oder Kompilieren des Codes - ich bin mir sicher, ihr habt dafür Verständnis. Der Code ist einfach zu alt und wurde zu lange nicht mehr gewartet.
Kurzum: der Code wird hier einfach veröffentlicht - tut damit, was ihr wollt, aber "don't blame me". Ich lizenziere ihn bewusst nicht unter einer GPL o.ä., da ich hierfür die diversen Copyright-Hinweise im Code konsequenterweise anpassen oder entfernen müsste. Insofern bleibt das Urheberrecht in gewissem Sinne auch weiterhin bei mir, aber ihr könnt natürlich den Code als Inspiration nehmen oder gern auch einfach so wie er ist in anderen Projekten verwenden; wenn mein Name als Quelle irgendwo dabeistehen bliebe, wär nett. Keine Ahnung, welcher Lizenz das am nächsten käme - vielleicht sowas wie "CC-BY-SA". Jedenfalls: have fun! 😉

## Was tut CMon?

Das Tool dient einfach nur dazu, bei einem eingehenden (und auch bei einem ausgehenden) Anruf eine kleine Sprechblase im Tray (Infobereich, Benachrichtigungsbereich) der Windows-Taskleiste anzuzeigen mit der Rufnummer des Anrufers bzw. des Angerufenen - mehr kann und tut es nicht. Kein Telefonbuch-Abgleich, weder intern noch mit dem Telefonbuch der Box, keine sonstigen Optionen.

Die IP-Adresse oder der Hostname der FRITZ!Box muss beim ersten Start des Programms eingegeben werden und wird dann in einer INI-Datei gespeichert. Bei folgenden Aufrufen kann das Programm mit dem Parameter **/tray** aufgerufen werden, dabei wird
der Konfigurationsdialog übersprungen und das Programm minimiert sich direkt ins Tray.
Hierbei wird dann die in der INI-Datei gespeicherte Adresse der Box direkt übernommen.

Voraussetzung für die Funktion von CMon ist, dass der Anrufmonitor der FRITZ!Box aktiviert ist. Falls dies noch nicht der Fall ist, kann das durch Eingabe der Zahlenfolge

    #96*5*

an einem an die Box angeschlossenen Telefon erledigt werden.

Mehr Informationen zu CMon gibt es in folgendem Thread im IP-Phone-Forum:

  https://www.ip-phone-forum.de/showthread.php?t=167903

## Binaries

Die jeweils aktuelle Version kann als Binary hier heruntergeladen werden:

  https://www.fabianswebworld.de/downloads/tools/cmon/


Viel Spaß!

Das Programm ist kostenlos und darf (im Binary) unverändert gerne weitergegeben werden; für die Nutzung des Quellcodes gelten die oben unter "Wichtige Vorabbemerkungen" angegebenen Vereinbarungen.

Ich übernehme keinerlei Haftung für die Funktion des Programms oder für aus dem Gebrauch des Programms entstehende Schäden jeglicher Art.

(c) 2008, 2009, 2013 Fabian Schneider - www.fabianswebworld.de
