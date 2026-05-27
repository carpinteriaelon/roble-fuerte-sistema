# Sistema Operacional Roble Fuerte

**Carpintaria Premium · Madrid · Versão 1.0 · Maio 2026**

Este é o centro nervoso da empresa. Tudo o que acontece na operação — projetos, dinheiro, produção, equipe — está centralizado nas pastas e arquivos abaixo. Use este documento como mapa.

---

## Como Operar — Rotina Diária

Abra todos os dias, nesta ordem:

1. **`09_DASHBOARD/Dashboard_Executivo.html`** — visão geral em 30 segundos.
2. **`01_FINANCEIRO/Controle_Financeiro_Master.xlsx`** — semáforo de pagamentos.
3. **`02_PROJETOS/CRM_Projetos.xlsx`** — pipeline e prioridades.
4. **`05_PRODUCAO/Producao_Funcionarios.xlsx`** — o que cada equipe faz hoje.

Toda segunda-feira de manhã, atualize o **Planejamento Semanal** da produção. Todo dia 5 do mês, execute o **SOP 07 — Fechamento Financeiro Mensal**.

---

## Estrutura de Pastas

```
ROBLE FUERTE/
├── 00_Cadastros_Mestres.xlsx          ← Base central (clientes, fornecedores, funcionários, materiais, acabamentos)
├── LEIA-ME — Sistema Roble Fuerte.md  ← Este arquivo
│
├── 01_FINANCEIRO/
│   ├── Controle_Financeiro_Master.xlsx  ← Contas a pagar/receber, custos fixos, impostos, fluxo de caixa, semáforo
│   ├── Contas_a_Pagar/                  ← Comprovantes
│   ├── Contas_a_Receber/                ← Cópias de faturas emitidas
│   ├── Custos_Fixos/                    ← Recibos mensais (aluguel, luz, água)
│   ├── Faturas_Emitidas/                ← PDFs assinados
│   ├── Faturas_Recebidas/               ← Notas de fornecedores
│   ├── Fluxo_de_Caixa/                  ← Extratos bancários
│   ├── Gestoria/                        ← Documentos enviados à gestoria
│   └── Impostos/                        ← Modelos 303, 111, 130 etc.
│
├── 02_PROJETOS/
│   ├── CRM_Projetos.xlsx                ← Pipeline + ficha de projeto + dashboard
│   ├── 01_Leads/
│   ├── 02_Orcamentos_Enviados/
│   ├── 03_Aprovados/
│   ├── 04_Em_Producao/
│   ├── 05_Em_Lacado/
│   ├── 06_Em_Montagem/
│   ├── 07_Finalizados/
│   └── 08_Pos_Venda/
│   (mover a pasta do projeto entre estas conforme avança)
│
├── 03_CLIENTES/
│   ├── Cliente_Final/
│   ├── Arquitetos/
│   ├── Reformistas/
│   ├── Construtoras/
│   └── Empresas_Reforma/
│
├── 04_FORNECEDORES/
│   ├── Madeiras/
│   ├── Lacas_Vernizes/
│   ├── Ferragens/
│   ├── Transporte/
│   └── Servicos_Externos/
│
├── 05_PRODUCAO/
│   ├── Producao_Funcionarios.xlsx       ← Planejamento semanal, OFs, capacidade, diárias
│   ├── Planejamento_Semanal/
│   ├── Ordens_de_Fabricacao/            ← PDFs das OFs impressas para o chão de fábrica
│   ├── Controle_CNC/
│   ├── Lacado/
│   └── Instalacoes/
│
├── 06_FUNCIONARIOS/
│   ├── Diarias/
│   ├── Pagamentos/                      ← Recibos de salário
│   ├── Contratos/
│   ├── Equipe_Fabrica/
│   └── Equipe_Montagem/
│
├── 07_MARKETING/
│   ├── Fotos_Obras/                     ← Antes/depois de cada projeto
│   ├── Material_Comercial/              ← Catálogos, apresentações
│   ├── Redes_Sociais/                   ← Posts, vídeos
│   └── Portfolio/                       ← Curadoria dos melhores trabalhos
│
├── 08_DOCUMENTOS/
│   ├── Contratos/
│   ├── Planos_Tecnicos/
│   ├── Desenhos/
│   ├── Comprovantes/
│   ├── Documentos_Fiscais/
│   └── Modelos_Templates/               ← Template de orçamento, contrato, fatura
│
├── 09_DASHBOARD/
│   └── Dashboard_Executivo.html         ← Visão executiva em tempo real (abrir no navegador)
│
└── 10_SOPs_Procedimentos/
    └── Manual_SOPs_Roble_Fuerte.docx    ← Os 8 SOPs essenciais + checklists diário/semanal/mensal
```

---

## Mapa dos 4 Arquivos-Mestre

### 1. `00_Cadastros_Mestres.xlsx`
Base de dados central. Cadastre AQUI antes de criar qualquer orçamento ou fatura.

- **Clientes** — ID único (C-001, C-002...), CIF/NIF, tipo
- **Fornecedores** — categoria, condição de pagamento
- **Funcionários** — função, equipe, salário/diária
- **Materiais_Lacas** — custos de referência (alimenta orçamentos)
- **Acabamentos** — catálogo de cores e custos adicionais

### 2. `01_FINANCEIRO/Controle_Financeiro_Master.xlsx`
Tudo o que envolve dinheiro.

- **Resumo** — KPIs, semáforo geral
- **Contas_a_Pagar** — fornecedores, folha, impostos, custos fixos (semáforo automático 🔴🟡🟢)
- **Contas_a_Receber** — faturas emitidas (semáforo automático)
- **Custos_Fixos** — base recorrente mensal
- **Impostos** — calendário fiscal completo
- **Fluxo_de_Caixa** — projeção 12 meses

### 3. `02_PROJETOS/CRM_Projetos.xlsx`
Pipeline de obras de A a Z.

- **Dashboard** — projetos por status, KPIs operacionais
- **Pipeline** — lista de todos os projetos com status, prioridade, atraso
- **Ficha_Projeto** — template detalhado (duplicar para cada projeto novo)

### 4. `05_PRODUCAO/Producao_Funcionarios.xlsx`
A fábrica e a equipe.

- **Dashboard** — OFs em andamento, atrasos, custo da folha
- **Planejamento_Semanal** — distribuição diária por funcionário
- **Ordens_Fabricacao** — controle de cada OF com % concluído
- **Capacidade** — quantos projetos cabem em cada mês
- **Diarias_Funcionarios** — folha de ponto e custo mensal da equipe

---

## Regras de Ouro

1. **Nunca produzir sem sinal de 50% recebido.** Protege o caixa.
2. **Margem mínima de 35%** em todo orçamento. Premium: 45%.
3. **Provisionar 21% IVA + ~15% IRPF** em conta separada todo mês.
4. **Resposta ao cliente em até 2 horas** em horário comercial.
5. **Foto de cada projeto finalizado** vai para `07_MARKETING/Fotos_Obras/`.
6. **Toda exceção que vira regra** deve virar SOP novo.

---

## Semáforo Padrão

| Cor | Significado | Ação |
|-----|-------------|------|
| 🟢 | Em dia / OK | Manter ritmo |
| 🟡 | Atenção (vence em 7 dias / capacidade 90%) | Planejar |
| 🔴 | Urgente / Vencido / Sobrecarga | Ação imediata |

---

## Calendário Fiscal Espanha (Resumo)

| Quando | Modelo | O quê |
|--------|--------|-------|
| 20 jan/abr/jul/out | 303 | IVA trimestral |
| 20 jan/abr/jul/out | 111 | Retenções IRPF |
| 20 jan/abr/jul/out | 130 | Pagamento fracionado |
| 30 janeiro | 390 | Resumo anual IVA |
| 31 janeiro | 190 | Resumo IRPF retenções |
| 28 fevereiro | 347 | Operações > 3.005€ |
| 25 julho | 200 | Imposto de Sociedades |

Detalhes completos no Manual de SOPs (capítulo 8).

---

## Próximos Passos Recomendados

1. Preencher os cadastros reais (clientes, fornecedores, funcionários) em `00_Cadastros_Mestres.xlsx`.
2. Lançar contas pendentes reais no `Controle_Financeiro_Master.xlsx`.
3. Migrar projetos em andamento para o `CRM_Projetos.xlsx`.
4. Imprimir e ler o `Manual_SOPs_Roble_Fuerte.docx` com a equipe.
5. Definir uma reunião semanal fixa (sugestão: segunda 8h30) para revisar dashboard.

---

**Roble Fuerte · Sistema Operacional Cloud · Confidencial**
