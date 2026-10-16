# Lumina Studio

Editor de imagens para desktop (Windows, Linux e macOS), inspirado no Photoshop, com interface em português.

![Ícone](build/icon.png)

## Recursos

**Documentos**
- Várias abas abertas ao mesmo tempo, arrastar e soltar arquivos, colar imagens (Ctrl+V)
- Abre PNG, JPEG, WebP, GIF, BMP, SVG, ICO e AVIF
- Projeto próprio `.lumina` que guarda camadas, opacidade, modos de mesclagem e textos editáveis
- Exporta PNG, JPEG e WebP com qualidade e escala ajustáveis
- Histórico com até 60 passos por documento e painel para voltar a qualquer passo

**Camadas**
- Criar, duplicar, excluir, reordenar (arrastar no painel), renomear, bloquear e ocultar
- 16 modos de mesclagem (Multiplicação, Divisão, Sobrepor, Luz suave, Cor, Luminosidade…)
- Opacidade por camada, mesclar para baixo, mesclar visíveis, achatar
- Camada via cópia (Ctrl+J) e via recorte (Ctrl+Shift+J)
- Camadas de texto que continuam editáveis

**Ferramentas** (atalhos iguais aos do Photoshop)

| Tecla | Ferramenta | Tecla | Ferramenta |
|---|---|---|---|
| V | Mover (com seleção automática de camada) | B | Pincel (dureza, fluxo, pressão da caneta, modos) |
| M | Letreiro retangular / elíptico | N | Lápis (pixel art) |
| L | Laço livre e poligonal | E | Borracha |
| W | Varinha mágica | S | Carimbo (clonagem) |
| C | Cortar (proporções predefinidas) | R | Desfoque / Nitidez por pincel |
| I | Conta-gotas | O | Subexposição / Superexposição / Esponja |
| G | Balde de tinta e Degradê (linear, radial, angular, refletido) | T | Texto |
| U | Formas (retângulo, arredondado, elipse, polígono, estrela, linha, seta) | H / Z | Mão / Zoom |
| Ctrl+T | Transformação livre (escalar, girar, espelhar) | | |

**Seleções**: nova, adicionar (Shift), subtrair (Alt), interseção; difusão, expandir, contrair, suavizar, inverter, selecionar pixels da camada (Ctrl+clique na miniatura), traçar e preencher.

**Ajustes** (Imagem ▸ Ajustes): Brilho/Contraste, Níveis (com histograma), Curvas, Exposição, Vibração, Matiz/Saturação (com Colorir), Equilíbrio de cores, Preto e branco, Inverter, Posterizar, Limiar, Sépia, Dessaturar, Tom automático.

**Filtros**: Desfoque gaussiano, de movimento e radial, Máscara de nitidez, Ruído, Mediana, Pixelizar, Detectar bordas, Relevo, Pintura a óleo, Vinheta, Sombra projetada, Contorno. Todos com visualização ao vivo e respeitando a seleção.

**Imagem**: tamanho da imagem, tamanho da tela (com âncora), girar 90°/180°, espelhar, cortar para seleção, aparar transparência.

## Baixar o `.exe`

O GitHub Actions compila o aplicativo automaticamente a cada push:

1. Abra a aba **Actions** do repositório e escolha a execução mais recente de **Build desktop**.
2. Em **Artifacts**, baixe **LuminaStudio-Windows**. Ele traz:
   - `LuminaStudio-Setup-<versão>.exe`: instalador com atalho na área de trabalho
   - `LuminaStudio-Portable-<versão>.exe`: roda sem instalar

Para publicar uma versão oficial, crie uma tag (`git tag v1.0.0 && git push --tags`). O workflow anexa os instaladores à página de *Releases*.

> Como o executável não é assinado digitalmente, o Windows SmartScreen pode avisar na primeira execução. Clique em "Mais informações" e depois em "Executar assim mesmo".

## Desenvolvimento

Requisitos: Node.js 20 ou mais recente.

```bash
npm install          # instala Electron e electron-builder
npm start            # abre o app desktop
npm run web          # abre no navegador em http://localhost:8080 (sem instalar nada além do Node)

npm run dist:win     # gera o .exe (instalador + portátil) em dist/
npm run dist:linux   # gera AppImage
npm run dist:mac     # gera .dmg (precisa de macOS)
```

No Linux, `npm run dist:win` gera o `.exe` portátil sem Wine; o instalador NSIS precisa do Wine ou do workflow do GitHub.

### Estrutura

```
electron/main.js      processo principal: janela, diálogos nativos, arquivos, área de transferência
electron/preload.js   ponte segura (contextIsolation + sandbox) exposta como window.lumina
src/index.html        interface
src/css/style.css     tema escuro
src/js/main.js        menus, atalhos, painéis, diálogos e comandos
src/js/document.js    documento, camadas, histórico (copy-on-write) e operações de imagem
src/js/view.js        área de trabalho: zoom, panorâmica, renderização incremental
src/js/tools.js       todas as ferramentas
src/js/filters.js     ajustes e filtros
src/js/selection.js   máscaras de seleção, "formigas marchando", preenchimento por inundação
src/js/io.js          abrir/salvar/exportar (nativo ou navegador)
```

### Arquitetura em resumo

- Cada camada é um `<canvas>` do tamanho do documento. As edições **nunca alteram** um canvas que já está no histórico: elas criam uma cópia nova (copy-on-write). Por isso desfazer e refazer são instantâneos e o histórico só ocupa memória com o que realmente mudou.
- Durante um traço, a camada mostra uma prévia (`layer.preview`), e só a região alterada (retângulo sujo) é recomposta. Assim o pincel continua fluido em imagens grandes, por exemplo 4000×3000 com várias camadas.
- Seleções são máscaras alfa. Por isso suportam difusão e combinações (adicionar, subtrair, interseção), e todas as ferramentas e filtros as respeitam.
