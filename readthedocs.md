# Documentação do Suplemento para Google Docs

## Introdução

Bem-vindo à documentação do nosso suplemento para Google Docs. Esta ferramenta foi projetada para otimizar e automatizar tarefas repetitivas, permitindo que você crie documentos ricos e bem formatados com muito mais agilidade.

Seja para inserir blocos de texto padronizados, aplicar formatações complexas com um único clique ou usar o poder da inteligência artificial para extrair e organizar informações, este suplemento é o seu assistente pessoal para a criação de documentos.

## Primeiros Passos: O Menu Principal

Após instalar o suplemento, um novo menu chamado **"Custom Menu"** aparecerá na barra de ferramentas do seu Google Docs. É a partir dele que você acessará todas as funcionalidades.

Aqui está um resumo do que cada opção faz:

*   **Abrir Sidebar:** Abre a principal interface do suplemento, uma barra lateral onde a maior parte da mágica acontece.
*   **Cabeçalhos e Rodapés:** Permite que você insira imagens padronizadas no cabeçalho e rodapé do seu documento.
*   **Extrair `Nome:` da agenda:** Uma ferramenta de inteligência artificial que lê um texto de agenda e extrai uma lista limpa de nomes de pacientes.
*   **Configurar Propriedades:** Abre uma janela para você ajustar configurações específicas do suplemento.
*   **Limpar Cache:** Uma função de manutenção que força a atualização dos dados que o suplemento utiliza, garantindo que você sempre tenha as informações mais recentes.

## A Ferramenta Principal: A Sidebar

A sidebar é o coração do suplemento. Ao abri-la, você tem acesso a um painel de ferramentas poderoso para buscar, organizar e inserir conteúdo de forma dinâmica no seu documento.

### Visão Geral

Pense na sidebar como sua biblioteca de conteúdos. Ela se conecta a uma planilha de dados e a uma pasta de "prompts" (comandos de IA) para trazer todos os textos que você precisa ao alcance de um clique. O objetivo é simples: economizar seu tempo e garantir consistência nos seus documentos.

### A Interface da Sidebar

A interface foi desenhada para ser intuitiva. Vamos conhecer seus componentes:

*   **Botões de Filtro:** No topo, você encontrará botões que representam diferentes categorias de conteúdo (por exemplo, "Anamnese", "Evolução"). Clicar em um deles filtra a lista de resultados para mostrar apenas os itens daquela categoria.
*   **Campo de Busca:** Logo abaixo dos filtros, há uma caixa de busca. Digite qualquer termo para encontrar rapidamente o item que você procura em todas as categorias.
*   **Listas de Arrastar e Soltar (Drag-and-Drop):** A sidebar é dividida em duas listas:
    *   **Lista Superior (Resultados):** É aqui que os itens aparecem após uma busca ou filtro. Esta lista é sua "biblioteca".
    *   **Lista Inferior (Trabalho):** Esta é sua "área de trabalho". Arraste os itens que deseja usar da lista de cima para a lista de baixo. Você pode reordená-los como quiser. É o conteúdo desta lista que será inserido no documento quando você usar um dos botões de ação.
### Os Botões de Ação

Localizados no rodapé da sidebar, estes botões são as ações que você pode executar com os itens da sua "Lista de Trabalho".

#### Botão Adicionar (+)

Este é o botão principal para adicionar conteúdo. Ele pega todos os textos dos itens que você colocou na lista de trabalho e os anexa ao final do seu documento. Além disso, ele aplica automaticamente um conjunto de formatações pré-definidas, garantindo que o documento mantenha um estilo consistente.

#### Botão Inserir no Cursor

Similar ao botão "Adicionar", este também insere o conteúdo da sua lista de trabalho. A grande diferença é que ele insere o texto exatamente onde o seu cursor está posicionado no documento, em vez de no final. Outro detalhe importante é que ele **não** aplica nenhuma formatação automática, inserindo o texto puro.

#### Botão Pincel (Brush)

Esta é uma ferramenta de automação poderosa. O Pincel não utiliza o conteúdo da sua lista de trabalho. Em vez disso, ao ser clicado, ele "varre" o documento inteiro e realiza uma série de substituições de texto com base em um conjunto de regras pré-configuradas. É ideal para, por exemplo, trocar todos os títulos de uma seção por novos títulos de forma automática.

#### Botão Mágico (Magic)

Este botão abre as portas para a inteligência artificial. Ele abre uma nova janela onde você pode selecionar "prompts" (instruções de IA) pré-definidos. Geralmente, esses prompts são usados para transformar ou gerar texto. Por exemplo, você pode selecionar um texto no seu documento e usar um prompt para resumi-lo, reescrevê-lo em outro tom ou até mesmo traduzi-lo.

## Gerenciador de Cabeçalhos e Rodapés

Esta funcionalidade simplifica a personalização dos cabeçalhos e rodapés do seu documento.

Ao selecionar a opção **"Cabeçalhos e Rodapés"** no menu, uma janela se abrirá, apresentando uma lista de imagens disponíveis para o cabeçalho e para o rodapé. Essas imagens são carregadas de uma pasta específica no Google Drive, garantindo que todos os documentos possam seguir um padrão visual.

O uso é simples:
1.  Selecione a imagem que deseja usar para o cabeçalho (ou escolha "Nenhum").
2.  Selecione a imagem para o rodapé (um padrão pode já vir pré-selecionado).
3.  Clique em "Aplicar".

O suplemento removerá qualquer conteúdo anterior e inserirá as imagens escolhidas. Além disso, a data atual será sempre adicionada automaticamente no canto direito do cabeçalho, mantendo seus documentos sempre atualizados.

## Ferramenta de Extração com IA: Extraindo Nomes da Agenda

Esta é uma das ferramentas mais avançadas do suplemento, utilizando inteligência artificial para poupar um trabalho manual significativo.

Ao selecionar **"Extrair `Nome:` da agenda"**, uma janela se abrirá com uma grande caixa de texto. O fluxo de trabalho é o seguinte:

1.  **Cole o Texto:** Copie o conteúdo de um relatório ou agenda (mesmo que esteja desorganizado) e cole na caixa de texto.
2.  **Escolha o Filtro:** Abaixo da caixa de texto, você terá botões para filtrar os resultados. Você pode optar por extrair todos os nomes da agenda, apenas os do período da manhã ("Matutino") ou apenas os do período da tarde ("Vespertino").
3.  **Execute a Extração:** Ao clicar em um dos botões, a IA analisará o texto, identificará e extrairá de forma inteligente os horários e os nomes dos pacientes, mesmo em casos onde a formatação é irregular.

Após o processamento, o suplemento limpará o seu documento atual e escreverá a lista de nomes de pacientes, já ordenada por horário e filtrada conforme sua escolha, com cada nome precedido por "Nome: ".

## Configurações e Utilitários

Esta seção cobre as ferramentas de manutenção e personalização do suplemento.

### Configurar Propriedades

Ao clicar em **"Configurar Propriedades"**, você acessa uma tela onde pode definir valores personalizados que são usados por outras partes do sistema. Atualmente, esta tela permite que você defina os nomes das secretárias associadas a diferentes consultórios. Embora pareça um detalhe pequeno, manter essas informações centralizadas aqui permite que o sistema as utilize de forma consistente em outras funcionalidades, como nos prompts de IA.

### Limpar Cache

Para carregar rapidamente, o suplemento guarda em uma memória temporária (cache) os dados que ele busca da sua planilha e da pasta de prompts. No entanto, se você fizer uma alteração nesses locais (como adicionar um novo item na planilha), o suplemente pode não ver a mudança imediatamente.

A opção **"Limpar Cache"** serve exatamente para isso: ela apaga essa memória temporária e força o suplemento a buscar a versão mais recente de todos os dados na próxima vez que você abrir a sidebar. Use esta função sempre que seus dados de origem forem atualizados.
