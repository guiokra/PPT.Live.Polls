# PowerPoint Live Polls & Quizzes – Suplemento Microsoft PowerPoint

Este projeto é um aplicativo web completo e pronto para produção, desenvolvido em **React, TypeScript e Tailwind CSS**, preparado para funcionar como um suplemento (Add-In) para o **Microsoft PowerPoint**. 

O sistema permite que instrutores, palestrantes e professores criem enquetes ou quizzes de múltipla escolha diretamente no slide lateral do PowerPoint, projetem as perguntas com QR Code em tempo real na tela principal de exibição e permitam que os participantes acessem e respondam instantaneamente pelos seus celulares sem a necessidade de cadastros.

---

## 🎨 Principais Recursos

1. **🔌 Suplemento PowerPoint (Taskpane):** Painel compacto otimizado para a barra lateral do Office. Contém comandos nativos via `Office.js` para inserir textos de perguntas e o QR Code de votações diretamente no slide atual.
2. **⚙️ Painel do Instrutor:** Painel de controle completo para criar novas aulas interativas, configurar modos de respostas (Anônimo vs. Identificado) e quizzes (com peso de pontuação XP), gerenciar o banco de questões e exportar os relatórios completos em formato **CSV (Excel)** ou **JSON**.
3. **🖥️ Tela do Projetor (Apresentação):** Visual com layout limpo e de alto contraste (Dark Premium), ideal para projetores e compartilhamento de telas. Atualiza-se dinamicamente conforme os participantes votam, com suporte opcional a **Leaderboard / Placar de Líderes** no fim de quizzes competitivos.
4. **📱 Celular do Participante (Mobile):** Experiência de toque confortável e fluida, construída sob filosofia mobile-first. Sincroniza em tempo real com as mudanças de slides comandadas pelo apresentador.

---

## 🚀 Como Inicializar Localmente

### Pré-requisitos
1. **Node.js** (versão 18 ou superior)
2. **npm** (instalador nativo de pacotes do Node)

### Passo a Passo
1. Na raiz do projeto, instale as dependências:
   ```bash
   npm install
   ```
2. Inicialize o servidor de desenvolvimento rápido por meio do Vite:
   ```bash
   npm run dev
   ```
3. O servidor Vite rodará na porta `3000`. Acesse em seu navegador:
   * **Visualizador Central (com Simulador Integrado):** `http://localhost:3000`

---

## ⚙️ Banco de Dados em Tempo Real (Preparado para Produção)

A versão inicial utiliza uma arquitetura híbrida inteligente que sincroniza as ações instantaneamente através de `localStorage` e canais de transmissão de navegador (`BroadcastChannel`). Isso significa que, ao abrir duas abas no mesmo computador (uma em modo Aluno e outra em modo Projetor), **a comunicação em tempo real acontece de forma local e 100% offline.**

### Como Preparar para Produção (Firebase Firestore)
Para distribuir este aplicativo para usuários remotos em redes de celulares distintas, basta habilitar um banco real-time em nuvem. O Firestore do Firebase é o ideal.

#### Estrutura Recomendada do Firestore:
* `sessions/{sessionId}`
  * `id` (string)
  * `name` (string)
  * `currentQuestionIndex` (number)
  * `showResults` (boolean)
  * `revealAnswer` (boolean)
  * `isAnonymous` (boolean)
  * `isQuizMode` (boolean)
  * `status` (string)
* `sessions/{sessionId}/questions/{questionId}`
  * `text` (string)
  * `options` (array de strings)
  * `correctOptionIndex` (number ou null)
* `sessions/{sessionId}/votes/{voteId}`
  * `questionId` (string)
  * `participantId` (string)
  * `participantName` (string)
  * `selectedOptionIndex` (number)
  * `timestamp` (string)
* `sessions/{sessionId}/participants/{participantId}`
  * `name` (string)
  * `score` (number)

#### Configuração Simples:
Insira sua chave nas variáveis de ambiente na raiz do projeto (arquivo `.env`) e inicialize a escuta Firestore substituindo os callbacks em `/src/services/store.ts` por `onSnapshot` do SDK do Firebase.

---

## 🔌 Sideload: Como Instalar o Suplemento no PowerPoint

Para testar o aplicativo diretamente de dentro do Microsoft PowerPoint em seu computador, siga os passos de **sideload local** do arquivo de manifesto:

### Passo 1: Atualizar o arquivo `manifest.xml`
Abra `/public/manifest.xml` com qualquer editor de texto e as seguintes linhas devem apontar para a sua URL final HTTPS (uma vez feito o deploy):
* `<SourceLocation DefaultValue="SUA_URL_AQUI?role=addin" />`
* `<AppDomain>SUA_URL_AQUI</AppDomain>`

*Nota: Se estiver testando localmente, configure para `https://localhost:3000?role=addin` (necessita de túnel HTTPS ou certificado SSL local configurado por meio do comando `office-addin-dev-certs install`).*

### Passo 2: Sideload no PowerPoint para Windows
1. Compartilhe uma pasta na sua rede de arquivos e insira o arquivo `manifest.xml` dentro dela.
2. No PowerPoint, vá em **Arquivo > Opções > Central de Confiabilidade > Configurações da Central de Confiabilidade > Catálogos de Suplementos Confiáveis**.
3. No campo **URL do Catálogo**, coloque a rota de pasta de rede compartilhada e clique em **Adicionar Catálogo**.
4. Marque a caixa **Mostrar no Menu** e salve.
5. Reinicie o PowerPoint. Crie uma apresentação vazia, vá em **Inserir > Meus Suplementos > Pasta de Compartilhamento Compartilhada** e dê duplo clique no **PowerPoint Live Polls**.

### Passo 3: Sideload no PowerPoint para Web (Navegador)
1. Acesse o [Office Online](https://office.com) de sua conta profissional ou acadêmica e abra uma apresentação no PowerPoint para Web.
2. Vá no menu superior **Inserir > Suplementos do Office**.
3. Na caixa superior, selecione a aba **Meus Suplementos** e escolha o link **Carregar meu Suplemento / Upload My Add-in**.
4. Selecione o arquivo local `/public/manifest.xml` e clique em **Enviar**. A barra lateral do PowerPoint Live Polls se abrirá imediatamente!

---

## 📽️ Como Usar Durante Uma Apresentação

Siga este fluxo simples no dia de sua exposição:
1. **Prepare a Sessão:** Abra o suplemento no PowerPoint ou acesse o Painel do Instrutor no navegador. Crie uma aula interativa e prepare as perguntas (ou puxe do banco de modelos existente).
2. **Monte os Slides:** 
   * Na barra lateral do suplemento, clique no botão **"Texto do Slide"** para colocar automaticamente o enunciado e os botões A-B-C-D nas caixas de textos de seu slide atual.
   * Clique em **"Inserir QR Code"** para colocar a imagem de votações customizada diretamente no seu slide de PowerPoint.
3. **Projete a Apresentação:** No projetor da sala, projete seus slides em tela cheia normalmente ou abra a aba **Tela do Projetor** do aplicativo na tela estendida.
4. **Instrua a Audiência:** Peça para os participantes escanear o QR Code de seus próprios assentos ou digitar o código de 4 dígitos no navegador do celular.
5. **Comande no PowerPoint:** Conforme você muda e avança as perguntas pela barra lateral do PowerPoint, a tela do celular dos alunos muda e acompanha seu ritmo de forma automática! 
6. **Revele Resultados:** Clique no botão **"Exibir Resultados"** da barra lateral para atualizar o gráfico de barras ao vivo e divertir os alunos. Se for quiz, clique em **"Revelar Gabarito"** para dar pontuações!
7. **Baixe o Relatório:** Ao término de suas exposições pedagógicas, vá na aba **Resultados** e extraia o balanço de frequências e acertos em um maravilhoso documento **CSV (Excel)**.
