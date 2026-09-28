# Demonstração — Organização de arquivos de faturamento

> Projeto de portfólio de **Victória Pedrosa**. **Demonstração** de organização de arquivos de faturamento — versão com dados fictícios (nomes, CNPJs, e-mails e IDs internos substituídos).

## Problema de negócio
Arquivos de faturamento de cada empresa precisavam ser organizados por mês (prestado e tomado).

## Antes x depois
| | Antes | Depois |
|---|---|---|
| Como é feito | Organização manual de pastas por squad. | Robô organiza os arquivos na estrutura CÓDIGO-NOME / MMAAAA / PRESTADO + TOMADO; a pasta raiz de cada squad vem das Propriedades do script. |

## Ganho
- Pastas padronizadas para todos os squads com um só código.

## Tecnologias
Google Apps Script, Google Drive, Google Sheets

## Arquivos
- `Codigo.gs`

## Como usar
Crie um projeto no Google Apps Script, copie os arquivos `.gs`/`.html` e configure as Propriedades do script indicadas no código.

## Autora
Victória Pedrosa — Product Owner do Time de IA, automação de processos contábeis e fiscais.
