# Lumina Studio

Suíte de criação e produtividade para desktop (Windows, Linux e macOS), com interface em português e funcionamento offline. São quatro apps num só programa:

| App | Para quê | Abre | Salva / exporta |
|---|---|---|---|
| **Lumina Foto** | Edição de imagens | PNG, JPG, WebP, GIF, BMP, SVG, `.lumina` | `.lumina`, PNG, JPG, WebP |
| **Lumina Texto** | Documentos | `.lumdoc`, Word (`.docx`), Markdown, HTML, TXT | `.lumdoc`, Word, PDF, HTML, Markdown, TXT |
| **Lumina Planilhas** | Planilhas e análise de dados | `.lumsheet`, Excel (`.xlsx`), CSV | `.lumsheet`, Excel, CSV, PDF |
| **Lumina Apresentações** | Slides | `.lumslides` | `.lumslides`, PowerPoint (`.pptx`), PDF, PNG |

A tela inicial reúne os modelos, os arquivos recentes e um botão para abrir qualquer arquivo. Ela escolhe sozinha o app certo. Também dá para arrastar arquivos para a janela. O tema pode ser claro, escuro ou seguir o sistema.

## Lumina Texto

- Estilos de parágrafo (Título, Subtítulo, Título 1–4, Citação, Código), fontes, tamanhos, cores, realce, alinhamento, entrelinhas e recuos
- **Estilo do documento**: troca a aparência do documento inteiro de uma vez. Presets: Padrão, **ABNT**, Moderno, Clássico e Compacto
- **Comandos com `/`**: digite `/` para inserir títulos, listas, tabela, imagem, sumário, destaques, divisor, quebra de página…
- **Atalhos Markdown**: `# ` para título, `- ` para lista, `1. ` para lista numerada, `[] ` para lista de tarefas, `> ` para citação e `**negrito**`, `*itálico*`, `` `código` ``
- Barra flutuante ao selecionar texto, tabelas com ferramentas próprias e imagens redimensionáveis com alinhamento
- **Sumário automático**, painel de estrutura, localizar e substituir (com expressão regular), contagem de palavras e modo foco
- **Histórico de versões** guardado dentro do arquivo, recuperação automática após travamentos e várias abas
- Configuração de página (A4, Carta, Ofício…, margens, orientação, números de página)
- Corretor ortográfico em português (no app desktop)
- Modelos: trabalho acadêmico ABNT (capa, sumário, seções, referências), relatório, carta, currículo e ata

## Lumina Planilhas

- Grade rápida desenhada em canvas, que cresce até 1.048.576 linhas
- **Fórmulas em português ou em inglês**: `=SOMA(A1:A10)` e `=SUM(A1:A10)` funcionam igual. Aceita `;` com vírgula decimal (`=ARRED(A1*1,5;2)`) ou `,`
- **Mais de 200 funções**, entre elas:
  - PROCV, PROCX, ÍNDICE/CORRESP;
  - SOMASES, CONT.SES, MÉDIASES;
  - datas, textos, financeiras (PGTO, TIR, VPL…) e estatísticas;
  - SE, SES, SEERRO e LET.
- **Matrizes dinâmicas**: FILTRO, CLASSIFICAR, ÚNICO, SEQUÊNCIA, PEGAR, EMPILHARV… com resultados que se espalham pelas células vizinhas e referência `A1#`
- Autocompletar de funções com dica de sintaxe, referências coloridas e **clique para inserir referências**, inclusive de outra aba
- Alça de preenchimento com séries inteligentes (1, 2, 3… / jan, fev… / segunda, terça… / Item 1, Item 2…)
- Formatos brasileiros (R$, %, `dd/mm/aaaa`, `1.234,56`) e códigos personalizados (`#.##0,00 "kg"`)
- Formatação condicional (escalas de cor, barras de dados, maior/menor que, duplicados, top N…), bordas, mesclar células e pincel de formatação
- Filtros, classificação em vários níveis, remover duplicatas e listas suspensas (validação)
- **Tabela dinâmica feita de fórmulas**, que se atualiza sozinha quando os dados mudam
- **Gráficos** (colunas, barras, linhas, área, pizza, rosca, dispersão) que você pode arrastar e redimensionar
- Congelar painéis, várias planilhas por arquivo, copiar e colar com o Excel e o Google Sheets, e barra de status com soma, média, mínimo e máximo
- Modelos: orçamento pessoal, lista de tarefas e controle de vendas

## Lumina Apresentações

- 8 temas modernos (Claro, Aurora, Meia-noite, Coral, Floresta, Pitch, Papel e Oceano) e 9 layouts
- Texto, formas (11 tipos), imagens com filtros e cantos arredondados, tabelas e **gráficos nativos**
- **Guias magnéticas** ao arrastar, alinhar e distribuir, girar, camadas, duplicar arrastando com Alt e seleção por área
- Fundos em cor, degradê ou imagem, **transições** (esmaecer, empurrar, zoom, cobrir, girar) e **animações de entrada**
- **Modo apresentação** em tela cheia:
  - clique ou setas para avançar e voltar;
  - tela preta ou branca (B/W) e ponteiro laser (L);
  - anotações com cronômetro e prévia do próximo slide (S);
  - ir direto a um slide com número + Enter.
- Exporta para **PowerPoint** (com gráficos editáveis e anotações), PDF e PNG
- Modelo de pitch de startup com 8 slides

## Lumina Foto

Camadas com 16 modos de mesclagem, seleções (letreiro, laço, varinha mágica), transformação livre, pincel com pressão, carimbo, desfoque/nitidez, subexposição/superexposição, degradê, texto editável, formas, mais de 25 ajustes e filtros com visualização ao vivo, e histórico.

## Baixar o `.exe`

O GitHub Actions compila o aplicativo a cada push:

1. Abra a aba **Actions** do repositório e escolha a execução mais recente de **Build desktop**.
2. Em **Artifacts**, baixe **LuminaStudio-Windows**. Ele traz o instalador (`LuminaStudio-Setup-<versão>.exe`) e a versão portátil (`LuminaStudio-Portable-<versão>.exe`).

Para publicar uma versão oficial, crie uma tag (`git tag v2.0.0 && git push --tags`). Os instaladores vão para a página de *Releases*.

> Como o executável não tem assinatura digital, o Windows SmartScreen pode avisar na primeira execução. Clique em "Mais informações" e depois em "Executar assim mesmo".

## Desenvolvimento

Requisitos: Node.js 20 ou mais recente.

```bash
npm install          # instala dependências e copia as bibliotecas para src/vendor
npm start            # abre o app desktop
npm run web          # abre no navegador em http://localhost:8080

npm run dist:win     # gera o .exe em dist/
npm run dist:linux   # gera AppImage
npm run dist:mac     # gera .dmg (precisa de macOS)
```

### Estrutura

```
electron/            processo principal (protocolo app://, diálogos, PDF, corretor, menu de contexto) e preload
src/index.html       tela inicial da suíte (src/shell/)
src/shared/          estilo visual, componentes de interface, ícones e ponte com o sistema de arquivos
src/photo/           Lumina Foto
src/docs/            Lumina Texto (editor, importação/exportação, modelos)
src/sheets/          Lumina Planilhas (motor de fórmulas, grade, gráficos, Excel/CSV)
src/slides/          Lumina Apresentações (temas, renderização, editor, PPTX/PDF)
src/vendor/          bibliotecas copiadas de node_modules no "npm install" (não versionadas)
scripts/vendor.js    cópia das bibliotecas
```

Bibliotecas de terceiros (todas com licença livre): [ExcelJS](https://github.com/exceljs/exceljs) (xlsx), [Mammoth](https://github.com/mwilliamson/mammoth.js) (leitura de docx), [docx](https://github.com/dolanmiu/docx) (escrita de docx), [PptxGenJS](https://github.com/gitbrent/PptxGenJS) (pptx), [Chart.js](https://www.chartjs.org/) (gráficos) e [DOMPurify](https://github.com/cure53/DOMPurify) (limpeza de HTML colado).

### Arquitetura em resumo

- Cada app roda num `iframe` da tela inicial, servido pelo protocolo `app://`, então todos têm a mesma origem. Os apps continuam abertos quando você troca de um para outro, e o fechamento do programa pergunta sobre arquivos não salvos em todos eles.
- **Planilhas**: o motor de fórmulas é próprio. Ele tem analisador léxico, analisador de precedência e avaliação preguiçosa de intervalos, além de detectar referências circulares e espalhar resultados matriciais. Inserir ou excluir linhas e colunas reescreve as referências das fórmulas em todas as abas. O desfazer funciona por transações.
- **Texto**: o editor usa `contentEditable` com histórico próprio (instantâneos e restauração da posição do cursor), para que o desfazer funcione também nas operações estruturais (tabelas, listas, blocos).
- **Apresentações**: um único renderizador gera o slide em DOM para edição, miniaturas, apresentação e exportação. Por isso o PDF sai igual ao que aparece na tela.
- **Foto**: canvases imutáveis no histórico (copy-on-write) e recomposição apenas da região alterada.
