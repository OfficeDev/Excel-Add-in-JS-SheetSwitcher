# <a name="sheet-switcher-task-pane-add-in-sample-for-excel-2016"></a>Aufgabenbereich-Add-In-Beispiel für Arbeitsblattschnellzugriff für Excel 2016

_Gilt für: Excel 2016_

Mit diesem Aufgabenbereich-Add-In können Sie neue Arbeitsblätter zu einer Arbeitsmappe über den Aufgabebereich hinzufügen und in Excel 2016 zu den verschiedenen Arbeitsblättern navigieren. Es ist in zwei Versionen verfügbar: Text-Editor und Visual Studio.

![Arbeitsblatt-Schnellzugriff-Beispiel](../Images/SheetSwitcher_taskpane.PNG)

## <a name="try-it-out"></a>Probieren Sie es aus
### <a name="text-editor-version"></a>Text-Editor-Version

Am einfachsten können Sie Ihr Add-In bereitstellen und testen, indem Sie die Dateien in eine Netzwerkfreigabe kopieren.

1.  Kopieren Sie die Dateien im Ordner „Text-Editor“, und hosten Sie sie unter Verwendung eines Servers Ihrer Wahl.
2.  Bearbeiten Sie die Elemente \<SourceLocation\> und \<URL\> der Manifestdatei (SheetSwitcherManifest.xml) so, dass sie auf den gehosteten Speicherort aus Schritt 1 zeigt (z. B. https://localhost/SheetSwitcher/home.html).
3.  Kopieren Sie das Manifest (SheetSwitcherManifest.xml) in eine Netzwerkfreigabe (z. B. \\\MyShare\MyManifests).
4.  Fügen Sie den Freigabepfad, unter dem das Manifest enthalten ist, als vertrauenswürdigen App-Katalog in Excel hinzu.

    a.  Starten Sie Excel, und öffnen Sie ein leeres Arbeitsblatt.

    b.  Klicken Sie auf die Registerkarte **Datei**, und klicken Sie dann auf **Optionen**.

    c.  Wählen Sie **Sicherheitscenter** aus, und klicken Sie dann auf die Schaltfläche **Einstellungen für das Sicherheitscenter**.

    d.  Wählen Sie **Vertrauenswürdige App-Kataloge** aus.

    e.  Geben Sie im Feld  **Katalog-URL** den Pfad zu der in Schritt 1 erstellten Netzwerkfreigabe ein, und klicken Sie auf **Katalog hinzufügen**.

   f. Aktivieren Sie das Kontrollkästchen **Im Menü anzeigen**, und wählen Sie dann **OK**. Eine Meldung wird angezeigt, dass Ihre Einstellungen angewendet werden, wenn Office das nächste Mal gestartet wird.

5.  Testen und führen Sie das Add-In aus.

    a.  Klicken Sie auf der Registerkarte **Einfügen** in Excel 2016 auf **Meine-Add-Ins**.

    b.  Wählen Sie im Dialogfenster **Office-Add-Ins** die Option **Freigegebener Ordner** aus.

    c.  Wählen Sie **Arbeitsblatt-Schnellzugriff-Beispiel**>**Einfügen** aus. Das Add-In wird in einem Aufgabenbereich rechts neben dem aktuellen Arbeitsblatt geöffnet, wie in der folgenden Abbildung dargestellt.

  ![Arbeitsblatt-Schnellzugriff-Beispiel](../Images/SheetSwitcher_taskpane.PNG)

    d.  Klicken Sie auf die Schaltfläche **Blätter hinzufügen**. Dabei werden vierzehn neue Arbeitsblätter zur Arbeitsmappe hinzugefügt. Klicken Sie auf die Blattschaltflächen im Aufgabebereich, um zum entsprechenden Arbeitsblatt in der Arbeitsmappe zu navigieren.


### <a name="visual-studio-version"></a>Visual Studio-Version
1.  Kopieren Sie das Projekt in einen lokalen Ordner, und öffnen Sie die Datei „Excel-Add-in-Sheet-Switcher.sln“ in Visual Studio.
2.  Drücken Sie F5, um das Beispiel-Add-In zu erstellen und bereitzustellen. Excel wird gestartet und das Add-In wird in einem Aufgabenbereich rechts neben einem leeren Arbeitsblatt geöffnet, wie in der folgenden Abbildung dargestellt.

  ![Arbeitsblatt-Schnellzugriff-Beispiel](../Images/SheetSwitcher_taskpane.PNG)

3. Klicken Sie auf die Schaltfläche **Blätter hinzufügen**. Dabei werden vierzehn neue Arbeitsblätter zur Arbeitsmappe hinzugefügt. Klicken Sie auf die Blattschaltflächen im Aufgabebereich, um zum entsprechenden Arbeitsblatt in der Arbeitsmappe zu navigieren.



### <a name="learn-more"></a>Weitere Informationen

Die Excel-JavaScript-APIs haben viel mehr bei der Entwicklung von Add-Ins zu bieten. Im Folgenden werden nur einige der verfügbaren Ressourcen aufgeführt.

1.  [Programmierungsübersicht für Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Codeausschnitt-Explorer für Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Codebeispiele zu Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md)
4.  [JavaScript-API-Referenz zu Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Erstellen Ihres ersten Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)