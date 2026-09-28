# Junta Planilhas

Une vários arquivos Excel e CSV em uma única planilha, casando as colunas **pelo nome**, mesmo quando estão em ordem diferente, com grafias diferentes ou quando alguns arquivos têm colunas a mais.

Tudo roda no navegador. Nenhum arquivo é enviado a servidor, e a ferramenta funciona offline: é um único arquivo HTML, sem instalação.

![Tela do Junta Planilhas](docs/captura.png)

## Como usar

1. Baixe o arquivo [`dist/junta-planilhas.html`](dist/junta-planilhas.html) e abra no navegador (Chrome, Edge, Firefox ou Safari atualizados).
2. Arraste os arquivos (`.xlsx`, `.xlsm`, `.xls` ou `.csv`) para a área indicada ou clique em **Escolher arquivos**.
3. Escolha o modo de junção e as opções.
4. Clique em **Juntar arquivos** e confira a prévia e os avisos.
5. Digite o nome do arquivo e clique em **Baixar planilha**.

Para experimentar, use os arquivos da pasta [`exemplos/`](exemplos/). Eles têm colunas em ordem e grafia diferentes, linha de título, cabeçalho repetido e um CSV no formato brasileiro.

## Funcionalidades

### Modos de junção

| Modo | Resultado |
| --- | --- |
| **Tudo numa aba só** | Empilha as linhas de todas as abas de todos os arquivos em uma aba "Unificado". |
| **Juntar abas de mesmo nome** | A aba "Vendas" de cada arquivo vira uma aba "Vendas" única, e assim por diante. |
| **Uma aba por origem** | Cada aba de cada arquivo vira uma aba separada no arquivo final. |

### Alinhamento pelo nome da coluna

- As colunas são casadas pelo nome do cabeçalho, não pela posição. Um arquivo com `Produto, SKU, Preço` e outro com `SKU, Preço, Produto` se alinham corretamente.
- Com a opção **Casar colunas ignorando maiúsculas, acentos e espaços**, "Preço", "preco" e "PREÇO " viram a mesma coluna.
- Colunas que só existem em alguns arquivos entram no resultado, com células em branco nas origens que não as têm.
- A ordem das colunas no resultado segue a primeira vez em que cada uma aparece.

### Limpeza automática

- **Linhas de título.** Títulos acima do cabeçalho, comuns em exportações de ERP, são detectados e ignorados.
- **Cabeçalhos repetidos.** Linhas iguais ao cabeçalho no meio dos dados (planilhas coladas umas sobre as outras) são removidas.
- **Colunas vazias.** Colunas sem cabeçalho e sem nenhum dado são descartadas.
- **Colunas de origem.** Opcionalmente, adiciona "Arquivo Origem" e "Aba Origem" no começo de cada linha, para rastrear de onde veio cada registro.

### CSV no formato brasileiro

- Detecta o separador (`;`, `,`, tabulação ou `|`).
- Lê arquivos em UTF-8 ou Windows-1252 (acentos de arquivos gerados pelo Excel no Windows).
- Converte números como `1.234,56` para valores numéricos.
- Mantém como texto códigos com zero à esquerda (SKU, CEP) e números longos (EAN, CPF, CNPJ), para não perder dígitos nem virar notação científica.

### Conferência antes de baixar

- **Relatório de colunas.** Lista quais colunas não existem em quais arquivos e quais grafias diferentes foram unificadas, para que nenhum desalinhamento passe despercebido.
- **Avisos.** Linhas de título ignoradas, cabeçalhos repetidos removidos, abas vazias e divisões por limite de linhas.
- **Prévia.** As 10 primeiras linhas de cada aba do resultado.
- **Aviso de desatualização.** Se a lista de arquivos ou as opções mudarem depois da junção, a ferramenta pede para juntar de novo antes de baixar.

### Carregamento e exportação

- Barra de progresso com o percentual real de leitura dos arquivos e da junção.
- Nome do arquivo de saída definido por você. Caracteres inválidos no Windows (`/ \ : * ? " < > |`) são trocados por `_`.
- A planilha final sai com filtro no cabeçalho e largura das colunas ajustada ao conteúdo.
- Abas com mais de 1.048.576 linhas (limite do Excel) são divididas automaticamente em partes.

## Limitações

- Sem a opção de cabeçalho marcada, as colunas são juntadas pela posição, pois não há nomes para casar.
- A detecção de linha de título considera as 15 primeiras linhas de cada aba. Cabeçalhos mais abaixo que isso não são encontrados.
- Fórmulas entram como o valor calculado, não como fórmula. Formatação (cores, bordas, mesclagens) não é preservada.
- Em CSV, datas permanecem como texto no formato em que estavam no arquivo.
- Todas as abas de cada arquivo são incluídas, inclusive as ocultas.
- O limite de tamanho depende da memória do computador. Arquivos com centenas de MB podem deixar o navegador lento.

## Estrutura do repositório

```
junta-planilhas/
├── dist/
│   └── junta-planilhas.html   # arquivo pronto para uso (gerado pelo build)
├── src/
│   ├── index.template.html    # interface: HTML, CSS e JavaScript
│   └── nucleo.js              # lógica: leitura, limpeza, alinhamento e exportação
├── exemplos/                  # arquivos para testar
├── docs/
│   └── captura.png
├── build.js                   # embute o SheetJS e gera o dist/
└── package.json
```

O `dist/junta-planilhas.html` fica versionado para que qualquer pessoa possa baixá-lo e usar direto, sem precisar compilar.

## Desenvolvimento

Requer [Node.js](https://nodejs.org) 16 ou superior.

```bash
npm install
npm run build
```

O build lê `src/index.template.html`, substitui os marcadores `/*__SHEETJS__*/` e `/*__NUCLEO__*/` pelo código do SheetJS e do núcleo, e grava `dist/junta-planilhas.html`. Depois de alterar qualquer arquivo em `src/`, rode o build de novo e abra o HTML gerado no navegador.

### Como funciona

- **Leitura.** Cada arquivo é lido com `FileReader`, que informa o progresso, e interpretado pelo SheetJS. Planilhas são lidas com datas como `Date`. CSVs são decodificados primeiro (UTF-8, com recuo para Windows-1252), têm o separador e o formato decimal detectados e são lidos como texto bruto, antes da conversão de números feita em `converterValorCSV`.
- **Preparação** (`prepararBloco`). Para cada aba, localiza a linha de cabeçalho, descarta colunas vazias, gera nomes para colunas sem cabeçalho (`Coluna N`) e remove cabeçalhos repetidos.
- **Alinhamento** (`combinar`). Cada nome de coluna é normalizado (`normalizar`) e vira uma chave. As linhas de cada origem são reposicionadas de acordo com essas chaves. Nomes duplicados dentro da mesma aba recebem sufixo, para não sobrescrever dados.
- **Exportação** (`gerarWorkbook`). Monta o arquivo com o SheetJS, aplica filtro e largura de colunas e divide abas acima do limite do Excel.

As funções do núcleo não dependem da interface e podem ser testadas no Node.js com o pacote `xlsx`, carregando `src/nucleo.js` com `XLSX` definido no escopo global.

### Testes manuais sugeridos

Antes de publicar uma alteração, junte os arquivos de `exemplos/` nos três modos e confira:

- se as colunas "Preço/preco" e "Produto/produto" foram unificadas;
- se o preço `1.234,56` do CSV virou número (alinhado à direita na prévia) e o SKU `004` continua com o zero;
- se a linha de título e o cabeçalho repetido de `vendas_jan.xlsx` foram removidos;
- se o relatório aponta a coluna "Canal" como ausente em dois arquivos.

## Segurança

O projeto usa o SheetJS 0.18.5, a última versão publicada no registro do npm. Essa versão tem vulnerabilidades conhecidas (CVE-2023-30533 e CVE-2024-22363), exploráveis com arquivos maliciosos preparados de propósito. O risco é baixo em uso local com arquivos de origem conhecida, mas o GitHub pode exibir alertas do Dependabot.

As versões corrigidas são distribuídas pelo site do SheetJS, não pelo npm. Para atualizar, troque a dependência no `package.json`:

```json
"xlsx": "https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz"
```

Depois rode `npm install`, `npm run build` e repita os testes acima.

## Tecnologias

- [SheetJS Community Edition](https://github.com/SheetJS/sheetjs): leitura e escrita de planilhas (licença Apache 2.0).
- JavaScript puro, sem framework.

O SheetJS é embutido no `dist/junta-planilhas.html`. Ao distribuir o arquivo, mantenha o aviso de licença dele.

## Licença

Defina a licença do projeto adicionando um arquivo `LICENSE` na raiz. A licença MIT é uma escolha comum e compatível com a biblioteca usada.
