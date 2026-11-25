# Canvas Bíblico para web

Editor independente que funciona no navegador, sem Obsidian, Excalidraw ou servidor de aplicação. Cole um texto, organize as palavras por arrasto, selecione várias com Shift e aplique cores e realces. Adicione notas e alterne entre as posições originais e as organizadas.

**Salvar projeto** baixa um arquivo `.canvasbiblico`; use **Abrir projeto** para continuar editando em outro momento ou dispositivo. O navegador também guarda automaticamente uma cópia local, que pode ser apagada ao limpar os dados do navegador. Mantenha um arquivo salvo como cópia de segurança.

**Exportar PowerPoint** cria dois slides com textos editáveis: as posições originais e o quadro organizado. Esta primeira versão não implementa todos os recursos e animações do plugin original; também não importa arquivos de desenhos do Excalidraw.

## Desenvolvimento

Execute `npm install` e `npm run build`. Publique `index.html`, `style.css` e `app.bundle.js` juntos; caminhos relativos permitem hospedar em um subdiretório no GitHub Pages. O pacote gerado inclui a biblioteca de PowerPoint e não requer servidor para processar os arquivos.

Para publicar no GitHub Pages, crie um repositório com o conteúdo desta pasta na raiz, envie a branch `main` e selecione **Settings → Pages → Build and deployment → Source → GitHub Actions**. O fluxo incluído publica automaticamente cada atualização da branch. O repositório ainda precisa ser criado e configurado no GitHub.
