# vDocSign

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
        Mit add_picture(image_path, width=Inches(2.5)) wird das Bild des Stempels eingefügt. Die Breite kann angepasst werden.

    Seite hinzufügen:
        add_page_break() fügt eine neue Seite hinzu, damit der Stempel am Ende platziert wird.

    Speichern des modifizierten .docx-Dokuments:
        Das Dokument wird unter output_docx_path gespeichert.

    Konvertieren in PDF:
        Die Funktion convert() aus docx2pdf wandelt das .docx-Dokument in ein PDF um.  


Wichtige Hinweise

    Kompatibilität: docx2pdf funktioniert nur unter Windows und macOS. Für Linux kannst du pandoc oder LibreOffice zur Konvertierung verwenden.

    Bildformat: Der Stempel mit Unterschrift sollte im Format .png mit transparentem Hintergrund sein, um ein sauberes Ergebnis zu erzielen.

    Pfad-Anpassungen!

