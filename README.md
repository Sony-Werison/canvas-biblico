# Canvas Bíblico para web

Editor independente que funciona no navegador, sem Obsidian, Excalidraw ou servidor de aplicação. Cole um texto, organize as palavras por arrasto, selecione várias com Shift e aplique cores e realces. Adicione notas e alterne entre as posições originais e as organizadas.

A barra lateral organiza as ferramentas nas abas Texto, Realces, Setas, Anotações, Verso e Layout. O quadro mostra os limites de cada slide 16:9; linhas de texto adicionadas ou reorganizadas são ajustadas para não atravessar esses limites.

Use Ctrl+Z para desfazer alterações do quadro e Ctrl+Shift+Z ou Ctrl+Y para refazê-las. Na aba Layout, selecione uma palavra e use **Dividir antes da palavra** para iniciar outro trecho na exportação; a divisão fica salva no projeto. Se um trecho ultrapassar um slide, a divisão automática continua em funcionamento.

**Salvar projeto** baixa um arquivo `.json`; use **Abrir projeto** para continuar editando em outro momento ou dispositivo. Arquivos `.canvasbiblico` criados anteriormente também podem ser abertos. O navegador guarda automaticamente uma cópia local, que pode ser apagada ao limpar os dados do navegador. Mantenha um arquivo salvo como cópia de segurança.

**Exportar PowerPoint** divide o texto longo em trechos, criando para cada trecho um slide com as posições originais e outro com o quadro organizado. Esta versão ainda não implementa todos os recursos do plugin original e não importa arquivos de desenhos do Excalidraw.

## Desenvolvimento

Execute `npm install` e `npm run build`. Publique `index.html`, `style.css` e `app.bundle.js` juntos; caminhos relativos permitem hospedar em um subdiretório no GitHub Pages. O pacote gerado inclui a biblioteca de PowerPoint e não requer servidor para processar os arquivos.

Para publicar no GitHub Pages, crie um repositório com o conteúdo desta pasta na raiz, envie a branch `main` e selecione **Settings → Pages → Build and deployment → Source → GitHub Actions**. O fluxo incluído publica automaticamente cada atualização da branch. O repositório ainda precisa ser criado e configurado no GitHub.
