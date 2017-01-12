# <a name="sheet-switcher-task-pane-add-in-sample-for-excel-2016"></a>Ejemplo del complemento del panel de tareas del conmutador de hojas para Excel 2016

_Se aplica a: Excel 2016_

Este complemento del panel de tareas proporciona una forma de agregar nuevas hojas a un libro mediante el panel de tareas y navegar fácilmente a las diferentes hojas en Excel 2016. Hay dos tipos: editor de texto y Visual Studio.

![Ejemplo de conmutador de hojas](../Images/SheetSwitcher_taskpane.PNG)

## <a name="try-it-out"></a>Pruébelo
### <a name="text-editor-version"></a>Versión del editor de texto

La forma más sencilla de implementar y probar el complemento consiste en copiar los archivos en un recurso compartido de red.

1.  Copie los archivos en la carpeta del Editor de texto y hospédelos mediante su servidor favorito.
2.  Edite los elementos \<SourceLocation\> y \<Url\> del archivo de manifiesto (SheetSwitcherManifest.xml) para que apunte a la ubicación hospedada del paso 1 (p. ej., https://hostlocal/ConmutadorHojas/home.html).
3.  Copie el manifiesto (SheetSwitcherManifest.xml) en un recurso compartido de red (por ejemplo, \\\MiRecursoCompartido\ConmutadorHojas).
4.  Agregue la ubicación del recurso compartido que contiene el manifiesto como un catálogo de aplicaciones de confianza en Excel.

    a.  Inicie Excel y abra una hoja de cálculo en blanco.

    b.  Elija la pestaña **Archivo** y, a continuación, **Opciones**.

    c.  Elija **Centro de confianza** y, a continuación, el botón **Configuración del Centro de confianza**.

    d.  Elija **Catálogos de aplicaciones de confianza**.

    e.  En el cuadro **URL del catálogo**, escriba la ruta de acceso al recurso compartido de red que creó en el paso 1 y, a continuación, elija **Agregar catálogo**.

   f. Active la casilla **Mostrar en el menú** y elija **Aceptar**. Aparecerá un mensaje para informarle de que la configuración se aplicará la próxima vez que inicie Office.

5.  Pruebe y ejecute el complemento.

    a.  En la pestaña **Insertar** de Excel 2016, elija **Mis complementos**.

    b.  En el cuadro de diálogo **Complementos de Office**, elija **Carpeta compartida**.

    c.  Elija **Ejemplo de conmutador de hojas**>**Insertar**. El complemento se abrirá en un panel de tareas a la derecha de la hoja de cálculo actual, como se muestra en la ilustración siguiente.

  ![Ejemplo de conmutador de hojas](../Images/SheetSwitcher_taskpane.PNG)

    d.  Haga clic en el botón **Agregar hojas**. Esto agregará catorce hojas nuevas en la hoja de cálculo. Haga clic en cualquiera de los botones de la hoja representados en el panel de tareas para ir a esa hoja del libro.


### <a name="visual-studio-version"></a>Versión de Visual Studio
1.  Copie el proyecto en una carpeta local y abra Excel-Add-in-Sheet-Switcher.sln en Visual Studio.
2.  Pulse F5 para crear e implementar el complemento de ejemplo. Excel se inicia y se abre el complemento en un panel de tareas a la derecha de una hoja de cálculo en blanco, como se muestra en la siguiente ilustración.

  ![Ejemplo de conmutador de hojas](../Images/SheetSwitcher_taskpane.PNG)

3. Haga clic en el botón **Add Sheets!** (Agregar hojas). Esto agregará catorce hojas nuevas en la hoja de cálculo. Haga clic en cualquiera de los botones de la hoja representados en el panel de tareas para ir a esa hoja en el libro.



### <a name="learn-more"></a>Más información

Las API de JavaScript de Excel tienen mucho que ofrecer para el desarrollo de complementos. A continuación se muestran algunos de los recursos disponibles.

1.  [Introducción a la programación de complementos de Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Explorador de fragmentos de código para Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Ejemplos de código de complementos de Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md)
4.  [Referencia de la API de JavaScript de complementos de Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Compilar el primer complemento de Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)