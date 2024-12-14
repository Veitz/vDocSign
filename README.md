# vDocSign

## Windows

Ein Python3-Programm, das .docx-Dateien einliest und einen eingescannten Stempel mit Unterschrift anhängt und das Dokument anschließend als PDF speichert. 

Bibliotheken

    python-docx: Zum Lesen und Bearbeiten von .docx-Dateien.
    Pillow: Zum Bearbeiten von Bilddateien (Stempel und Unterschrift).
    reportlab: Zum Erstellen und Konvertieren von PDFs.
    docx2pdf: Zum einfachen Konvertieren von .docx in PDF (nur auf Windows und macOS).

Benötigte Pakete:

    pip install python-docx Pillow reportlab docx2pdf


Erklärung des Codes

    Einlesen des .docx-Dokuments:
        Die Datei wird mit Document(docx_path) geöffnet.

    Stempel/Unterschrift hinzufügen:
        Mit add_picture(image_path, width=Inches(2.5)) wird das Bild des Stempels eingefügt. 
        Die Breite kann angepasst werden.

    Seite hinzufügen:
        add_page_break() fügt eine neue Seite hinzu, damit der Stempel am Ende platziert wird.

    Speichern des modifizierten .docx-Dokuments:
        Das Dokument wird unter output_docx_path gespeichert.

    Konvertieren in PDF:
        Die Funktion convert() aus docx2pdf wandelt das .docx-Dokument in ein PDF um.  


Wichtige Hinweise

    Kompatibilität: docx2pdf funktioniert nur unter Windows und macOS. Für Linux kannst du pandoc 
    oder LibreOffice zur Konvertierung verwenden.

    Bildformat: Der Stempel mit Unterschrift sollte im Format .png mit transparentem Hintergrund sein, 
    um ein sauberes Ergebnis zu erzielen.

    Pfad-Anpassungen!


## Linux


Benötigten Python-Pakete und LibreOffice:

    pip install python-docx Pillow
    sudo apt-get install libreoffice

python-docx: Zum Bearbeiten von .docx-Dateien.  
Pillow: Zum Arbeiten mit Bildern.  
LibreOffice: Zum Konvertieren von .docx in PDF.  

Codes

    Einlesen des .docx-Dokuments:
        Mit Document(docx_path) wird das .docx-Dokument geladen.

    Stempel hinzufügen:
        Mit add_picture(image_path, width=Inches(2.5)) wird das Bild eingefügt. Die Breite kann angepasst werden.

    Speichern des .docx-Dokuments:
        Das bearbeitete .docx wird unter output_docx_path gespeichert.

    Konvertieren mit LibreOffice:
        subprocess.run() führt LibreOffice im Headless-Modus aus, um das .docx in ein PDF zu konvertieren:

        libreoffice --headless --convert-to pdf <dateiname> --outdir <zielverzeichnis>

        Der Headless-Modus erlaubt die Ausführung ohne grafische Oberfläche.

Wichtige Hinweise

    Bildformat: Das Bild für den Stempel sollte vorzugsweise ein .png mit transparentem Hintergrund sein.  
    Pfad-Anpassungen: Stelle sicher, dass die Pfade zu den Dateien korrekt sind.  
    LibreOffice-Verfügbarkeit: Überprüfe mit libreoffice --version, ob LibreOffice korrekt installiert ist.  

    Um den Stempel direkt am Ende des vorhandenen Textes auf derselben Seite einzufügen, 
    müssen wir den Seitenumbruch entfernen und stattdessen leere Absätze einfügen. 
    Damit wird sichergestellt, dass der Stempel nur dann nach unten geschoben wird, 
    wenn tatsächlich Platz vorhanden ist.

Blockimport

    Der Blockimport wird per vDocSign_linux_Blockimport realisiert.
    Der Code liest alle .docx-Dateien aus dem Ordner "input-docxs" ein, 
    fügt den Stempel entsprechend ein und speichert die bearbeiteten Dateien 
    mit dem gleichen Namen im Ausgabeordner "output-docxs". Anschließend werden 
    die Dateien in PDFs umgewandelt und ebenfalls gespeichert.

Ausführung des Programms

    Führe das gewünschte Skript aus:

    python3 xxx.py



