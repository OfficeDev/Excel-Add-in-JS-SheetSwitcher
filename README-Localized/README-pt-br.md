# Exemplo de Suplemento do Painel de Tarefas do Seletor de Planilhas para o Excel 2016

_Aplica-se ao: Excel 2016_

Este suplemento do painel de tarefas oferece uma maneira de adicionar novas planilhas a uma pasta de trabalho usando o painel de tarefas e navegar facilmente até as diversas planilhas do Excel 2016. Há dois tipos: o editor de texto e o Visual Studio.

![Exemplo de Seletor de Planilhas](../images/SheetSwitcher_taskpane.PNG)

## Experimente
### Versão do editor de texto

A maneira mais fácil de implantar e testar o suplemento é copiar os arquivos em um compartilhamento de rede.

1.  Crie uma pasta em um compartilhamento de rede (por exemplo, \\\MyShare\SheetSwitcher) e copie todos os arquivos na pasta do Editor de Texto. 
2.  Edite o elemento <SourceLocation> do arquivo de manifesto de modo que ele aponte para o local de compartilhamento da etapa 1. 
3.  Copie o manifesto (SheetSwitcherManifest.xml) em um compartilhamento de rede (por exemplo, \\\MyShare\MyManifests).
4.  Adicione o local de compartilhamento que contém o manifesto como um catálogo de aplicativos confiáveis no Excel.

    a. Inicie o Excel e abra uma planilha em branco.  
    
    b. Escolha a guia **Arquivo** e escolha **Opções**.
    
    c. Escolha **Central de Confiabilidade** e, em seguida, escolha o botão **Configurações da Central de Confiabilidade**.
    
    d. Escolha **Catálogos de Aplicativos Confiáveis**.
    
    e. Na caixa  **URL de Catálogo**, insira o caminho para o compartilhamento de rede que você criou na etapa 1 e escolha **Adicionar Catálogo**.
    
   f.  Marque a caixa de seleção **Mostrar no Menu** e escolha **OK**. Será exibida uma mensagem para informá-lo de que suas configurações serão aplicadas na próxima vez que você iniciar o Office. 
        
5.  Teste e execute o suplemento. 

    a. Na **guia Inserir** no Excel 2016, escolha **Meus Suplementos**. 
    
    b. Na caixa de diálogo **Suplementos do Office**, escolha **Pasta Compartilhada**.
    
    c. Escolha **Amostra do Seletor de Planilhas**>**Inserir**. O suplemento abre em um painel de tarefas à direita da planilha atual, conforme mostrado na figura a seguir. 
        
  ![Exemplo de Seletor de Planilhas](../images/SheetSwitcher_taskpane.PNG)

    d. Clique no botão **Adicionar Planilhas!**. Isso adicionará 14 novas folhas à planilha. Clique em qualquer um dos botões da planilha renderizados no painel de tarefas para navegar até essa planilha na pasta de trabalho.
        

### Versão do Visual Studio
1.  Copie o projeto para uma pasta local e abra o Excel-Add-in-Sheet-Switcher.sln no Visual Studio.
2.  Pressione F5 para criar e implantar o suplemento de exemplo. O Excel inicia e o suplemento abre em um painel de tarefas à direita da planilha em branco, conforme mostrado na figura a seguir. 
        
  ![Exemplo de Seletor de Planilhas](../images/SheetSwitcher_taskpane.PNG)

3. Clique no botão **Adicionar Planilhas!**. Isso adicionará 14 novas folhas à planilha. Clique em qualquer um dos botões da planilha renderizados no painel de tarefas para navegar até essa planilha na pasta de trabalho.



### Saiba mais

As APIs JavaScript para Excel têm muito mais a oferecer à medida que você desenvolve suplementos. Confira a seguir alguns dos recursos disponíveis. 

1.  [Visão geral da programação de Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Explorador de trecho para Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Exemplos de código de Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md) 
4.  [Referência da API JavaScript para Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Crie seu primeiro Suplemento do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)
