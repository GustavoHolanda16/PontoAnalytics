# PontoAnalytics
Dashboard web client-side para análise de folha de ponto — horas extras, sábados, domingos, feriados e regra de proteção noturna (21h). Sem back-end, sem instalação.

---

## Tecnologias

- HTML, CSS e JavaScript puro (sem frameworks)
- [SheetJS](https://sheetjs.com/) para leitura de arquivos `.xlsx`
- Google Fonts (Sora, DM Serif Display, DM Mono)

---

## Como usar

1. Clone o repositório
2. Abra o arquivo `sistema_folha_ponto.html` diretamente no navegador
3. Faça upload do relatório exportado do sistema de ponto (`.xls`, `.xlsx` ou `.html`)

Nenhum servidor ou instalação é necessário.

---

## Funcionalidades

**Dashboard**
- KPIs gerais: funcionários, horas extras, horas trabalhadas, média de dias, faltas
- Horas trabalhadas separadas por sábado, domingo e feriado
- Contador de funcionários beneficiados pela regra de proteção noturna

**Tabela de funcionários**
- Busca por nome ou cargo
- Filtros por horas extras, sábado, domingo e faltas
- Ordenação por qualquer coluna
- Exportação para `.csv`

**Modal de detalhe**
- KPIs individuais com horas separadas por tipo de dia
- Calendário mensal com registro diário
- Identificação visual de sábados, domingos, feriados e faltas

**Regra de proteção noturna (21h)**

Quando um funcionário trabalha até as 21h ou depois, qualquer atraso registrado no dia seguinte é automaticamente desconsiderado. Faltas continuam sendo contabilizadas normalmente. Os dias em que a regra foi aplicada são sinalizados no calendário do modal.

---

## Estrutura

```
pontoanalytics/
├── sistema_folha_ponto.html
├── style.css
└── Script.js
```

---

## Formato de arquivo suportado

O parser espera o formato padrão de exportação do sistema de ponto, com blocos de 4 tabelas por funcionário: dados cadastrais, registros diários, totalizadores e resumo do período.

---

