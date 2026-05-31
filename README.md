# mecflu-tank-draining

[![Python](https://img.shields.io/badge/Python-3.x-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Microsoft Word](https://img.shields.io/badge/Microsoft_Word-relatório-2B579A?logo=microsoftword&logoColor=white)](https://www.microsoft.com/microsoft-365/word)
![Status](https://img.shields.io/badge/status-Etapa_2-8A2BE2)

Projeto da disciplina **Mecânica dos Fluidos**, ministrada pelo professor **Júlio César**, para analisar o esvaziamento de um reservatório e comparar os resultados medidos com o modelo teórico.

## Experimento

O aparato utiliza uma **garrafa PET transparente de 2 L** com água. Durante o escoamento por um pequeno orifício lateral, são registrados o nível da água e o tempo para realizar a comparação **teórico-experimental** com o modelo baseado em Bernoulli e continuidade.

```mermaid
flowchart LR
    A["PET de 2 L com água"] --> B["Escoamento pelo orifício"]
    B --> C["Medição de nível e tempo"]
    C --> D["Tratamento dos dados"]
    D --> E["Análise teórico-experimental"]
```

## Documentos

- [Protocolo experimental](docs/mecflu-experimental-protocol.docx)
- [Proposta da disciplina](docs/mecflu-proposal.pdf)

## Estrutura

```text
.
├── docs/          # proposta e protocolo
├── data/          # medições brutas e tratadas
├── calculations/  # cálculos e planilhas
├── images/        # registros visuais
└── scripts/       # automações auxiliares
```

## Cronograma

| Data | Marco |
| --- | --- |
| **22 de junho de 2026** | Entrega interna |
| **29 de junho de 2026** | Entrega oficial |

## Status

**Etapa 2** — protocolo experimental.
