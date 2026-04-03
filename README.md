# Validador de Glebas

Aplicacao web para validacao visual e rapida de coordenadas geodesicas em planilhas Excel, com foco em detectar inconsistencias que podem gerar erros de area invalida no fluxo de glebas.

## Visao Geral

O projeto foi construido como uma aplicacao estatica em `HTML + CSS + JavaScript`, sem necessidade de backend para o processamento principal. A leitura da planilha e a validacao das coordenadas acontecem no navegador do usuario.

Principais objetivos:

- validar glebas a partir de arquivos `.xls` e `.xlsx`
- identificar erros de estrutura e coordenadas
- apresentar relatorios visuais com status semanticos
- facilitar a analise por resumo geral e por gleba

## Interface

O sistema conta com:

- tema dark com foco em leitura e contraste
- area de upload por clique ou arrastar e soltar
- resumo lateral com quantidade de glebas, erros e glebas validas
- aba de relatorio com cards semanticos para sucesso, erro e alerta
- aba por gleba com detalhamento das ocorrencias encontradas

Observacao:

Nao encontrei um arquivo local da captura de tela na pasta do projeto para publicar junto no repositorio. Se voce salvar a imagem em algo como `docs/interface.png`, basta adicionar esta linha ao README:

```md
![Preview da aplicacao](docs/interface.png)
```

## Regras de Validacao

Atualmente a aplicacao verifica:

- poligono nao fechado
- pontos insuficientes
- ponto duplicado em excesso
- coordenada invalida

## Estrutura Esperada da Planilha

A aplicacao aceita colunas extras, mas precisa reconhecer corretamente:

- `Gleba`
- `Ponto`
- `Lat`
- `Long`

Exemplo de estrutura aceita:

| Cultura | Formato da Gleba | Area Nao Cultivada | Gleba | Ponto | Lat | Long | Alt |
|---|---|---|---|---|---|---|---|
| Soja | P | N | 1 | 1 | -14.43539142600 | -44.33006286500 | 0 |
| Soja | P | N | 1 | 2 | -14.43388207100 | -44.33622437000 | 0 |

## Arquivos do Projeto

- `index.html`: estrutura da interface
- `styles.css`: tema visual, layout e responsividade
- `app.js`: regras de validacao, leitura da planilha e renderizacao dos relatorios

## Arquivos de Teste

O repositorio inclui planilhas de exemplo para validacao manual:

- `TESTE_1_COM ERROS.xls`
- `TESTE_2_COM ERROS.xls`
- `TESTE_3_SEM ERROS.xls`

## Como Executar Localmente

Como a aplicacao e estatica, voce pode:

1. baixar o projeto
2. abrir o arquivo `index.html` no navegador

Se preferir, tambem pode servir localmente com qualquer servidor simples.

## Publicacao

O projeto esta pronto para publicacao em `GitHub Pages`, pois usa arquivos estaticos e nao depende de backend.

## Tecnologias

- HTML5
- CSS3
- JavaScript
- SheetJS (`xlsx`) para leitura de planilhas
