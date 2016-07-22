# Aufgabenbereich-Add-In-Beispiel f�r Arbeitsblattschnellzugriff f�r Excel 2016

_Gilt f�r: Excel 2016_

Mit diesem Aufgabenbereich-Add-In k�nnen Sie neue Arbeitsbl�tter zu einer Arbeitsmappe �ber den Aufgabebereich hinzuf�gen und in Excel 2016 zu den verschiedenen Arbeitsbl�ttern navigieren. Es ist in zwei Versionen verf�gbar: Text-Editor und Visual Studio.

![Arbeitsblatt-Schnellzugriff-Beispiel](../images/SheetSwitcher_taskpane.PNG)

## Probieren Sie es aus
### Text-Editor-Version

Am einfachsten k�nnen Sie Ihr Add-In bereitstellen und testen, indem Sie die Dateien in eine Netzwerkfreigabe kopieren.

1.  Erstellen Sie einen Ordner in einer Netzwerkfreigabe (z.�B. \\\MyShare\SheetSwitcher), und kopieren Sie alle Dateien im Ordner �Text-Editor�. 
2.  Bearbeiten Sie das <SourceLocation>-Element der Manifestdatei, damit es auf den Freigabepfad aus Schritt�1 zeigt. 
3.  Kopieren Sie das Manifest (SheetSwitcherManifest.xml) in eine Netzwerkfreigabe (z.�B. \\\MyShare\MyManifests).
4.  F�gen Sie den Freigabepfad, unter dem das Manifest enthalten ist, als vertrauensw�rdigen App-Katalog in Excel hinzu.

    a. Starten Sie Excel, und �ffnen Sie ein leeres Arbeitsblatt.  
    
    b. Klicken Sie auf die Registerkarte **Datei**, und klicken Sie dann auf **Optionen**.
    
    c. W�hlen Sie **Trust Center** aus, und klicken Sie dann auf die Schaltfl�che **Einstellungen f�r das Trust Center**.
    
  d. W�hlen Sie **Vertrauensw�rdige App-Kataloge** aus.
    
  e. Geben Sie im Feld **Katalog-URL** den Pfad zu der in Schritt�1 erstellten Netzwerkfreigabe ein, und klicken Sie auf **Katalog hinzuf�gen**.
    
   f. Aktivieren Sie das Kontrollk�stchen **Im Men� anzeigen**, und w�hlen Sie dann **OK**. Eine Meldung wird angezeigt, dass Ihre Einstellungen angewendet werden, wenn Office das n�chste Mal gestartet wird. 
        
5.  Testen und f�hren Sie das Add-In aus. 

  a. Klicken Sie auf der Registerkarte **Einf�gen** in Excel�2016 auf **Meine-Add-Ins**. 
    
  b. W�hlen Sie im Dialogfenster **Office-Add-Ins** die Option **Freigegebener Ordner** aus.
    
  c. W�hlen Sie **Arbeitsblatt-Schnellzugriff-Beispiel**>**Einf�gen**. The add-in opens in a task pane to the right of the current worksheet, as shown in the following figure. 
        
  ![Arbeitsblatt-Schnellzugriff-Beispiel](../images/SheetSwitcher_taskpane.PNG)

  d. Klicken Sie auf die Schaltfl�che **Bl�tter hinzuf�gen**. This will add fourteen new sheets to the spreadsheet. Click any of the sheet buttons rendered in the task pane to navigate to that sheet in the workbook.
        

### Visual Studio-Version
1.  Kopieren Sie das Projekt in einen lokalen Ordner, und �ffnen Sie die Datei �Excel-Add-in-Sheet-Switcher.sln� in Visual Studio.
2.  Dr�cken Sie F5, um das Beispiel-Add-In zu erstellen und bereitzustellen. Excel wird gestartet und das Add-In wird in einem Aufgabenbereich rechts neben einem leeren Arbeitsblatt ge�ffnet, wie in der folgenden Abbildung dargestellt. 
        
  ![Arbeitsblatt-Schnellzugriff-Beispiel](../images/SheetSwitcher_taskpane.PNG)

3. Klicken Sie auf die Schaltfl�che **Bl�tter hinzuf�gen**. Dabei werden vierzehn neue Arbeitsbl�tter zur Arbeitsmappe hinzugef�gt. Klicken Sie auf die Blattschaltfl�chen im Aufgabebereich, um zum entsprechenden Arbeitsblatt in der Arbeitsmappe zu navigieren.



### Weitere Informationen

Die Excel-JavaScript-APIs haben viel mehr bei der Entwicklung von Add-Ins zu bieten. Im Folgenden werden nur einige der verf�gbaren Ressourcen aufgef�hrt. 

1.  [Programmierungs�bersicht f�r Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Codeausschnitt-Explorer f�r Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Codebeispiele zu Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md) 
4.  [JavaScript-API-Referenz zu Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Erstellen Ihres ersten Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)
