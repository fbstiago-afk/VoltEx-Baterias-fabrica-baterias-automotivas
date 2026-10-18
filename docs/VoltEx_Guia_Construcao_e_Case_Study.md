# VoltEx Baterias S/A — Guia de Construção do Projeto Power BI & Case Study

> **Escopo:** montagem prática do relatório industrial VoltEx em `.pbip`, do modelo semântico ao último tooltip, e o estudo de caso para portfólio.
> **Base de referência:** `Base_VoltEx_Dashboard_Completo.xlsx` — 3 dimensões, 4 fatos, jan/2022 a dez/2025. Todos os números citados neste documento foram calculados diretamente dessa base.

---

## Sumário

**Parte 1 — Construção**
0. [Leitura da base antes de abrir o Desktop](#0-leitura-da-base-antes-de-abrir-o-desktop)
1. [Estrutura do projeto `.pbip`](#1-estrutura-do-projeto-pbip)
2. [Modelagem Star Schema e Calendário](#2-modelagem-star-schema-e-calendário)
3. [Camada de medidas DAX](#3-camada-de-medidas-dax)
4. [UX/UI — O Padrão VoltEx](#4-uxui--o-padrão-voltex)
5. [Painel de filtros retrátil](#5-painel-de-filtros-retrátil)
6. [New Card Visual com imagens dos produtos](#6-new-card-visual-com-imagens-dos-produtos)
7. [Tooltips dinâmicos de página](#7-tooltips-dinâmicos-de-página)
8. [Deneb / Vega-Lite](#8-deneb--vega-lite)
9. [Checklist de publicação](#9-checklist-de-publicação)

**Parte 2 — Case Study para Portfólio**

---

# PARTE 1 — ARQUITETURA E CONSTRUÇÃO

## 0. Leitura da base antes de abrir o Desktop

Antes de qualquer visual, a equipe precisa conhecer o grão de cada tabela e os pontos que vão quebrar medida se ninguém tratar.

| Tabela | Linhas | Grão | Data | Chaves |
|---|---:|---|---|---|
| `Dim_Produto` | 6 | 1 SKU | — | `ID_Produto` |
| `Dim_Cliente` | 5 | 1 conta B2B | — | `ID_Cliente` |
| `Dim_Fabrica` | 3 | 1 planta | — | `ID_Fabrica` |
| `Fato_Vendas` | 15.000 | Linha de pedido | `Data_Venda` | Produto, Cliente, Fábrica |
| `Fato_Producao` | 10.000 | Turno × produto × planta | `Data_Producao` | Produto, Fábrica |
| `Fato_Coleta_Cascos` | 8.000 | Coleta × cliente × produto | `Data_Coleta` | Cliente, Produto |
| `Fato_Estoque_Mensal` | 864 | **Snapshot** fim de mês × produto × planta | `Data_Referencia` | Produto, Fábrica |

### Pontos de atenção (resolver no Power Query ou documentar no modelo)

| # | Achado | Impacto | Tratamento |
|---|---|---|---|
| 1 | `URL_Imagem` aponta para `i.imgur.com/moto7ah.png` etc. — são *placeholders*, não as fotos reais | Cartões e tooltips renderizam vazios | Hospedar as 6 fotos (seção 6) e substituir a coluna |
| 2 | `Fato_Coleta_Cascos` **não tem** `ID_Fabrica` | Filtro de planta distorce a Taxa de Logística Reversa | Medida com guarda (seção 3.5) |
| 3 | `Valor_Desconto` varia de 0 a 15 sem unidade definida (R$/unidade? % ?) | Receita líquida ambígua | Faturamento bruto = `Quantidade × Preco_Venda`; desconto fica fora até a área comercial confirmar a unidade |
| 4 | As 3 plantas produzem e vendem os 6 SKUs, contrariando a especialização do Plano de Negócios (Manaus = motos, Joinville = pesada) | Análise por planta perde aderência ao processo real | Expor como achado de governança; não "corrigir" o dado |
| 5 | Não existe capacidade teórica (peças/hora) por linha | Fator **Performance** do OEE não é calculável | OEE = Disponibilidade × Qualidade, com Performance = 1 declarada explicitamente |
| 6 | Volume produzido (5,47 mi) é ~3,6× o volume vendido (1,53 mi) no período | Reconciliação produção × venda não fecha | Não cruzar as duas fatos em uma mesma razão; analisar cada uma no seu domínio |

---

## 1. Estrutura do projeto `.pbip`

Salve via **Arquivo › Salvar como › Projeto do Power BI (.pbip)** com as opções de pré-visualização **TMDL** e **PBIR** ativadas.

```text
voltex-industrial/
├── VoltEx.pbip
├── VoltEx.SemanticModel/
│   ├── definition/
│   │   ├── model.tmdl
│   │   ├── relationships.tmdl
│   │   ├── tables/
│   │   │   ├── Dim_Calendario.tmdl
│   │   │   ├── Dim_Produto.tmdl
│   │   │   ├── Fato_Producao.tmdl
│   │   │   ├── ...
│   │   │   └── _Medidas.tmdl          ← tabela vazia que hospeda TODAS as medidas
│   └── .pbi/cache.abf                ← no .gitignore
├── VoltEx.Report/
│   ├── definition/pages/<id>/visuals/<id>/visual.json
│   ├── StaticResources/RegisteredResources/
│   │   ├── VoltEx_Theme.json
│   │   └── bg_*.png                   ← planos de fundo
│   └── .pbi/localSettings.json       ← no .gitignore
├── assets/produtos/*.webp             ← fonte das imagens hospedadas
└── .gitignore
```

```gitignore
**/.pbi/localSettings.json
**/.pbi/cache.abf
```

**Fluxo de branches:** `main` (produção) ← `develop` ← `feature/modelo-*` e `feature/report-*`. Modelo e relatório nunca no mesmo PR: quem mexe em `.SemanticModel` não toca em `.Report`, e vice-versa. Isso torna o diff de TMDL legível e o code review de DAX real.

---

## 2. Modelagem Star Schema e Calendário

### 2.1 Diagrama

```text
                    ┌───────────────┐
                    │ Dim_Calendario│
                    └──────┬────────┘
        ┌──────────────┬───┴──────────┬────────────────┐
        ▼              ▼              ▼                ▼
 ┌────────────┐ ┌─────────────┐ ┌───────────────┐ ┌──────────────────┐
 │Fato_Vendas │ │Fato_Producao│ │Fato_Coleta_   │ │Fato_Estoque_     │
 │            │ │             │ │Cascos         │ │Mensal (snapshot) │
 └─┬───┬───┬──┘ └──┬───────┬──┘ └──┬────────┬───┘ └──┬────────┬──────┘
   │   │   │       │       │       │        │        │        │
   │   │   └──┐    │       │       │        │        │        │
   ▼   ▼      ▼    ▼       ▼       ▼        ▼        ▼        ▼
 Dim_Produto ◄─────┴───────┼───────┘    Dim_Cliente  │   Dim_Fabrica
 Dim_Cliente ◄─────────────┼────────────────┘        │
 Dim_Fabrica ◄─────────────┴─────────────────────────┘
```

Todos os relacionamentos: **1:\* , filtro em direção única (dimensão → fato), ativos.**

### 2.2 Por que um único calendário com quatro relacionamentos ativos funciona

As quatro fatos não se relacionam entre si. Logo, não existe caminho ambíguo: `Dim_Calendario[Data]` pode ter relacionamento **ativo** com `Data_Venda`, `Data_Producao`, `Data_Coleta` e `Data_Referencia` ao mesmo tempo. Um único slicer de período filtra as quatro fatos de forma coerente, e medidas que cruzam fatos (ex.: cascos ÷ vendas) recebem o mesmo recorte temporal sem `USERELATIONSHIP`.

`USERELATIONSHIP` só entra quando uma *mesma* fato tem duas datas (ex.: futura `Fato_Paradas` com `Inicio` e `Fim`).

### 2.3 `Dim_Calendario` (Power Query M)

Gerada no Power Query para que o range acompanhe os dados e o modelo continue dobrável (sem tabela calculada pesada).

```powerquery
let
    DataMin   = #date(2022, 1, 1),
    DataMax   = #date(2025, 12, 31),
    Dias      = Duration.Days(DataMax - DataMin) + 1,
    Lista     = List.Dates(DataMin, Dias, #duration(1, 0, 0, 0)),
    Tabela    = Table.FromList(Lista, Splitter.SplitByNothing(), {"Data"}, null, ExtraValues.Error),
    Tipada    = Table.TransformColumnTypes(Tabela, {{"Data", type date}}),
    Ano       = Table.AddColumn(Tipada, "Ano", each Date.Year([Data]), Int64.Type),
    MesNum    = Table.AddColumn(Ano, "Mes_Num", each Date.Month([Data]), Int64.Type),
    MesNome   = Table.AddColumn(MesNum, "Mes", each Text.Proper(Date.ToText([Data], "MMM", "pt-BR")), type text),
    AnoMes    = Table.AddColumn(MesNome, "AnoMes", each [Ano] * 100 + [Mes_Num], Int64.Type),
    AnoMesTxt = Table.AddColumn(AnoMes, "Ano_Mes", each Date.ToText([Data], "MMM/yy", "pt-BR"), type text),
    Trimestre = Table.AddColumn(AnoMesTxt, "Trimestre", each "T" & Text.From(Date.QuarterOfYear([Data])), type text),
    FimMes    = Table.AddColumn(Trimestre, "Fim_Mes", each Date.EndOfMonth([Data]), type date),
    Inverno   = Table.AddColumn(FimMes, "Janela_Inverno", each if List.Contains({5,6,7,8}, [Mes_Num]) then "Mai–Ago" else "Demais", type text)
in
    Inverno
```

Após carregar: **Marcar como tabela de datas** em `Data`; `Mes` e `Ano_Mes` classificados por `Mes_Num` e `AnoMes`. A coluna `Janela_Inverno` existe porque o Plano de Negócios concentra estoque entre maio e agosto — ela permite validar essa premissa (ver Insights).

### 2.4 Higiene do modelo

- [ ] Chaves (`ID_*`) e colunas de data das fatos **ocultas** na visão de relatório.
- [ ] `Dim_Produto[URL_Imagem]` com **Categoria de dados = URL da imagem**.
- [ ] Todas as agregações implícitas desligadas (`Resumir por: Nenhum` nas colunas numéricas das dimensões).
- [ ] Tabela `_Medidas` com pastas de exibição: `01 Produção`, `02 Manutenção`, `03 Qualidade`, `04 Vendas`, `05 Estoque`, `06 EcoCiclo`, `90 UX`.

---

## 3. Camada de medidas DAX

### 3.1 Bases (aditivas)

```dax
Qtd Produzida      = SUM ( Fato_Producao[Qtd_Produzida] )
Qtd Refugo         = SUM ( Fato_Producao[Qtd_Refugo] )
Horas Programadas  = SUM ( Fato_Producao[Horas_Turno_Programado] )
Horas Parada       = SUM ( Fato_Producao[Horas_Parada_Maquina] )
Qtd Falhas         = SUM ( Fato_Producao[Qtd_Falhas_Equipamento] )

Qtd Vendida        = SUM ( Fato_Vendas[Quantidade] )
Cascos Recolhidos  = SUM ( Fato_Coleta_Cascos[Qtd_Cascos_Recolhidos] )

Faturamento Bruto  =
SUMX ( Fato_Vendas, Fato_Vendas[Quantidade] * RELATED ( Dim_Produto[Preco_Venda] ) )

CPV =
SUMX ( Fato_Vendas, Fato_Vendas[Quantidade] * RELATED ( Dim_Produto[Custo_Fabricacao] ) )

Margem Bruta R$    = [Faturamento Bruto] - [CPV]
Margem Bruta %     = DIVIDE ( [Margem Bruta R$], [Faturamento Bruto] )
```

### 3.2 OEE

**Lógica.** OEE = Disponibilidade × Performance × Qualidade.
- *Disponibilidade* = tempo que a máquina efetivamente rodou ÷ tempo programado = `(Horas Programadas − Horas Parada) ÷ Horas Programadas`.
- *Qualidade* = peças boas ÷ peças produzidas.
- *Performance* exige capacidade teórica (peças/hora). A base não tem esse dado, então o fator fica em 1 **e isso aparece no tooltip do KPI** — ninguém deve comparar este OEE com benchmark de OEE completo sem essa nota.

A razão é calculada sobre somas, nunca média de percentuais por linha: um turno de 8 h não pode pesar o mesmo que um de 24 h.

```dax
Disponibilidade =
DIVIDE ( [Horas Programadas] - [Horas Parada], [Horas Programadas] )

Qualidade =
DIVIDE ( [Qtd Produzida] - [Qtd Refugo], [Qtd Produzida] )

Performance =
1    -- substituir por DIVIDE([Qtd Produzida], [Capacidade Teorica]) quando a capacidade por linha existir

OEE =
[Disponibilidade] * [Performance] * [Qualidade]
```

### 3.3 MTBF e MTTR

**Lógica.** A Matriz de KPIs propõe contar linhas com parada > 0 (`COUNTROWS(FILTER(...))`). A base já traz `Qtd_Falhas_Equipamento`, que é o denominador correto (um turno pode ter mais de uma falha) e evita um iterador. Na base atual os dois métodos dão o mesmo resultado (4.288 falhas = 4.288 turnos com parada), mas o `SUM` continua correto quando isso deixar de ser verdade.

- MTBF = tempo operando ÷ número de falhas.
- MTTR = tempo parado ÷ número de falhas.

```dax
MTBF (h) =
DIVIDE ( [Horas Programadas] - [Horas Parada], [Qtd Falhas] )

MTTR (h) =
DIVIDE ( [Horas Parada], [Qtd Falhas] )
```

### 3.4 Scrap Rate e custo do refugo

```dax
Scrap Rate =
DIVIDE ( [Qtd Refugo], [Qtd Produzida] )

Custo do Refugo R$ =
SUMX (
    Fato_Producao,
    Fato_Producao[Qtd_Refugo] * RELATED ( Dim_Produto[Custo_Fabricacao] )
)

Saving Potencial Refugo R$ =
-- Quanto custaria o refugo se a operação estivesse na meta de 1,5%
VAR _Meta = 0.015
VAR _Atual = [Scrap Rate]
RETURN
    IF ( _Atual > _Meta, [Custo do Refugo R$] * ( 1 - DIVIDE ( _Meta, _Atual ) ), 0 )
```

### 3.5 Estoque (semi-aditivo) e Giro

**Lógica.** Estoque é foto, não fluxo: somar 12 fechamentos mensais dá um número sem sentido. A medida precisa pegar **o último snapshot disponível no contexto** e somar apenas entre produtos e plantas.

Giro anualizado = vendas do período ÷ estoque médio do período, normalizado para 12 meses. Sem a normalização, filtrar um trimestre derrubaria o giro para ¼ e o semáforo ficaria vermelho por erro de cálculo.

```dax
Estoque Final =
VAR _UltimaFoto =
    CALCULATE ( MAX ( Fato_Estoque_Mensal[Data_Referencia] ) )
RETURN
    CALCULATE (
        SUM ( Fato_Estoque_Mensal[Qtd_Estoque_Pecas] ),
        Dim_Calendario[Data] = _UltimaFoto
    )

Estoque Medio =
AVERAGEX ( VALUES ( Dim_Calendario[AnoMes] ), [Estoque Final] )

Meses com Foto =
CALCULATE ( DISTINCTCOUNT ( Fato_Estoque_Mensal[Data_Referencia] ) )

Giro Anualizado =
DIVIDE ( [Qtd Vendida], [Estoque Medio] ) * DIVIDE ( 12, [Meses com Foto] )
```

### 3.6 Taxa de Logística Reversa (EcoCiclo)

**Lógica.** Cascos ÷ baterias vendidas no mesmo recorte. O problema: `Fato_Coleta_Cascos` não tem planta. Se o usuário filtrar "Planta Sul", o numerador continua com cascos de todas as plantas e o denominador só com vendas da Sul — a taxa infla ~3×. A medida se recusa a responder nesse contexto.

```dax
Taxa Logistica Reversa =
IF (
    ISFILTERED ( Dim_Fabrica ),
    BLANK (),                                  -- coleta não tem grão de planta
    DIVIDE ( [Cascos Recolhidos], [Qtd Vendida] )
)

Gap Cascos p/ Meta =
VAR _Meta = 0.75
RETURN
    MAX ( 0, _Meta * [Qtd Vendida] - [Cascos Recolhidos] )
```

No card, o subtítulo exibe *"Indicador não segmentável por planta"* quando a medida retorna vazio (medida de texto em 3.7).

### 3.7 Medidas de UX (pasta `90 UX`)

```dax
Cor OEE =
SWITCH ( TRUE (), [OEE] >= 0.85, "#2FBF71", [OEE] >= 0.70, "#FF8C00", "#E5484D" )

Cor Scrap =
SWITCH ( TRUE (), [Scrap Rate] < 0.015, "#2FBF71", [Scrap Rate] <= 0.03, "#FF8C00", "#E5484D" )

Cor MTBF =
SWITCH ( TRUE (), [MTBF (h)] > 300, "#2FBF71", [MTBF (h)] > 150, "#FF8C00", "#E5484D" )

Cor MTTR =
IF ( [MTTR (h)] < 2.5, "#2FBF71", "#E5484D" )

Cor Giro =
IF ( [Giro Anualizado] > 6, "#2FBF71", "#FF8C00" )

Cor Reversa =
SWITCH ( TRUE (), [Taxa Logistica Reversa] > 0.75, "#2FBF71", [Taxa Logistica Reversa] >= 0.50, "#FF8C00", "#E5484D" )

Subtitulo Reversa =
IF (
    ISFILTERED ( Dim_Fabrica ),
    "Indicador não segmentável por planta",
    "Meta > 75% · faltam " & FORMAT ( [Gap Cascos p/ Meta], "#,0" ) & " cascos"
)

URL Produto =
SELECTEDVALUE ( Dim_Produto[URL_Imagem] )
```

> `URL Produto`: definir **Categoria de dados = URL da imagem** na própria medida (Ferramentas de medida › Categoria de dados).

Faixas de MTBF intermediárias (150 h) são proposta desta equipe — a Matriz só define a meta de 300 h. Validar com a manutenção antes de publicar.

---

## 4. UX/UI — O Padrão VoltEx

### 4.1 Princípios

| Princípio | Regra prática |
|---|---|
| **O fundo é o layout** | Toda a estrutura (faixas, divisórias, áreas) vive no PNG de fundo. Visuais flutuam sem caixa, sem borda, sem sombra. |
| **Laranja é sinal, não decoração** | `#FF8C00` só em KPI crítico, alerta e elemento selecionado. Se a tela tem laranja em 5 lugares, ele parou de comunicar. |
| **Uma pergunta por zona** | Cada página tem no máximo 3 zonas de leitura: *status* (topo), *explicação* (centro), *ação/detalhe* (rodapé ou tooltip). |
| **Densidade por camadas** | Poucos visuais na tela; o detalhe sobe sob demanda em tooltip de página e drill-through. |

### 4.2 Tokens de design

| Token | Valor | Uso |
|---|---|---|
| `bg/base` | `#232B32` Grafite Industrial | Plano de fundo / wallpaper |
| `bg/raised` | `#2C363F` | Zonas do fundo, drawer |
| `line/subtle` | `#3A4651` | Divisórias desenhadas no PNG |
| `text/primary` | `#FFFFFF` | Valores, títulos |
| `text/secondary` | `#7A8B99` Cinza Técnico | Rótulos, eixos, legendas |
| `accent/energy` | `#FF8C00` | KPI crítico, seleção, alerta |
| `status/ok` | `#2FBF71` | Meta atingida |
| `status/bad` | `#E5484D` | Fora do padrão |
| Tipografia título | Montserrat / Segoe UI Black, 20–28 pt | Header HTML |
| Tipografia dados | Segoe UI Semibold, 9–12 pt | Visuais nativos |
| Grade | 8 px; margens de 24 px | Alinhamento |

### 4.3 Tema JSON (`VoltEx_Theme.json`)

```json
{
  "name": "VoltEx Industrial Dark",
  "dataColors": ["#FF8C00", "#7A8B99", "#C9D3DB", "#FFB547", "#4E5D6A", "#2FBF71", "#E5484D", "#A3B1BD"],
  "foreground": "#FFFFFF",
  "foregroundNeutralSecondary": "#7A8B99",
  "foregroundNeutralTertiary": "#4E5D6A",
  "background": "#232B32",
  "backgroundLight": "#2C363F",
  "backgroundNeutral": "#3A4651",
  "tableAccent": "#FF8C00",
  "good": "#2FBF71",
  "neutral": "#FF8C00",
  "bad": "#E5484D",
  "maximum": "#FF8C00",
  "center": "#7A8B99",
  "minimum": "#2C363F",
  "textClasses": {
    "title":   { "fontFace": "Segoe UI Semibold", "fontSize": 11, "color": "#FFFFFF" },
    "label":   { "fontFace": "Segoe UI",          "fontSize": 9,  "color": "#7A8B99" },
    "callout": { "fontFace": "Segoe UI Semibold", "fontSize": 24, "color": "#FFFFFF" },
    "header":  { "fontFace": "Segoe UI Semibold", "fontSize": 12, "color": "#FFFFFF" }
  },
  "visualStyles": {
    "*": {
      "*": {
        "background":     [{ "show": false }],
        "border":         [{ "show": false }],
        "dropShadow":     [{ "show": false }],
        "visualHeader":   [{ "show": false }],
        "visualTooltip":  [{ "background": { "solid": { "color": "#2C363F" } }, "titleFontColor": { "solid": { "color": "#7A8B99" } }, "valueFontColor": { "solid": { "color": "#FFFFFF" } } }]
      }
    },
    "page": {
      "*": {
        "background": [{ "transparency": 0 }],
        "outspace":   [{ "color": { "solid": { "color": "#232B32" } }, "transparency": 0 }]
      }
    }
  }
}
```

Importar em **Exibição › Temas › Procurar temas**. O `"*": { "background": false }` é o que mata as "caixinhas" de uma vez para todos os visuais.

### 4.4 Aplicando os planos de fundo

1. **Exporte cada fundo em 2560 × 1440 px** (2× a página de 1280 × 720). O Power BI reduz a imagem e o texto/linhas desenhados ficam nítidos em telas de alta densidade. Formato PNG, < 1 MB.
2. **Formatar página › Tela › Tipo: 16:9 (1280 × 720).**
3. **Formatar página › Plano de fundo da tela › Adicionar imagem** → `bg_executivo.png`; **Ajuste da imagem: Ajustar**; **Transparência: 0 %**.
4. **Formatar página › Papel de parede**: `#232B32`, transparência 0 % (evita a faixa cinza no modo "Ajustar à largura" do serviço).
5. Ligue **Exibição › Linhas de grade + Ajustar à grade** e posicione os visuais sobre as zonas desenhadas no fundo, nunca "aproximadamente".
6. Um fundo por *tipo* de página, não por página: `bg_executivo`, `bg_analitico` (header + 2 zonas), `bg_detalhe` (header + 1 zona larga), `bg_tooltip`. Quatro arquivos sustentam as 7 abas.

### 4.5 Header em HTML Content

Um visual **HTML Content** (AppSource) de 1280 × 64 px no topo de cada página, alimentado por medida:

```dax
Header HTML Executivo =
VAR _Periodo =
    FORMAT ( MIN ( Dim_Calendario[Data] ), "mmm/yy" ) & " — " &
    FORMAT ( MAX ( Dim_Calendario[Data] ), "mmm/yy" )
VAR _Chip =
    IF (
        [OEE] >= 0.85,
        "<span style='background:#2FBF7122;color:#2FBF71;border:1px solid #2FBF71;padding:3px 10px;border-radius:999px;font:600 11px Segoe UI'>OEE NA META</span>",
        "<span style='background:#FF8C0022;color:#FF8C00;border:1px solid #FF8C00;padding:3px 10px;border-radius:999px;font:600 11px Segoe UI'>OEE EM ALERTA</span>"
    )
RETURN
"<div style='display:flex;align-items:center;justify-content:space-between;height:64px;padding:0 24px;font-family:Segoe UI'>
   <div>
     <div style='font:900 22px Montserrat, Segoe UI Black;letter-spacing:.5px;color:#FFFFFF'>SUMÁRIO EXECUTIVO</div>
     <div style='font:400 11px Segoe UI;color:#7A8B99;margin-top:2px'>Operação industrial VoltEx · " & _Periodo & "</div>
   </div>
   <div>" & _Chip & "</div>
 </div>"
```

Fonte Montserrat só renderiza se estiver instalada na máquina do usuário; o fallback `Segoe UI Black` garante consistência no serviço.

---

## 5. Painel de filtros retrátil

### 5.1 Anatomia

```text
┌──────────────── Página 1280×720 ────────────────────────────┐
│ [≡ Filtros]  HEADER HTML                                    │
│                                                             │
│  ┌──────────┐┌──────────────────────────────────────────┐   │
│  │ DRAWER   ││ SCRIM (retângulo #000, 55% transparente) │   │
│  │ 320×720  ││ — clicável: fecha o drawer               │   │
│  │ ───────  ││                                          │   │
│  │ Trilha   ││                                          │   │
│  │ (HTML)   ││                                          │   │
│  │ ───────  ││                                          │   │
│  │ Período  ││                                          │   │
│  │ Planta   ││                                          │   │
│  │ Linha    ││                                          │   │
│  │ Cliente  ││                                          │   │
│  │ ───────  ││                                          │   │
│  │[Limpar][✕]                                          │   │
│  └──────────┘└──────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────┘
```

### 5.2 Passo a passo

1. **Monte o drawer**: retângulo `#2C363F` (320 × 720, x = 0), cabeçalho em HTML Content, e os **New Slicer** (Período em modo *Entre*; Planta, Linha e Cliente em modo *Botões em lista vertical*, tamanho de botão 288 × 36, preenchimento selecionado `#FF8C00` com 15 % de opacidade e borda esquerda laranja de 3 px).
2. **Monte o scrim**: retângulo preto, 55 % de transparência, x = 320, cobrindo o resto da página. Ação: **Indicador › BM_Filtros_Fechado**.
3. **Agrupe** no Painel de Seleção: `GRP_Drawer` = drawer + slicers + trilha + botões internos + scrim.
4. **Crie dois indicadores** (Exibição › Indicadores):

   | Indicador | Estado de `GRP_Drawer` | **Dados** | **Exibição** | Página atual |
   |---|---|:---:|:---:|:---:|
   | `BM_Filtros_Aberto` | visível | ☐ | ☑ | ☑ |
   | `BM_Filtros_Fechado` | oculto | ☐ | ☑ | ☑ |

   Desmarcar **Dados** é o detalhe que importa: sem isso, abrir o painel reseta os filtros que o usuário acabou de aplicar.

5. Em cada indicador, marque **Visuais selecionados** e selecione apenas `GRP_Drawer` — assim ele não congela o estado dos outros visuais da página.
6. **Botões**: ícone `≡` no header → `BM_Filtros_Aberto`; botão `✕` no drawer e o scrim → `BM_Filtros_Fechado`; botão **Limpar** → ação *Limpar todas as segmentações*.
7. **Sincronize os slicers** (Exibição › Sincronizar segmentações) entre as 7 abas. Copie o grupo de uma página para outra com Ctrl+C/Ctrl+V e aceite "Sincronizar".
8. Painel de filtros nativo: mantido **oculto para o consumidor** e usado só para filtros de página/relatório fixos (ex.: excluir dias sem produção). Ele não compete com o drawer.

### 5.3 Trilha de filtros (breadcrumbs)

Visual HTML Content dentro do drawer *e* uma versão compacta (só contagem) ao lado do ícone `≡` no header, para o usuário saber que há filtro ativo com o painel fechado.

```dax
Trilha Filtros HTML =
VAR _ChipIni = "<span style='display:inline-block;margin:0 6px 6px 0;padding:4px 10px;border:1px solid #FF8C00;border-radius:999px;color:#FFFFFF;font:600 11px Segoe UI'><span style='color:#7A8B99;font-weight:400'>"
VAR _ChipMeio = "</span> "
VAR _ChipFim = "</span>"
VAR _Periodo =
    IF (
        ISFILTERED ( Dim_Calendario ),
        _ChipIni & "Período" & _ChipMeio
            & FORMAT ( MIN ( Dim_Calendario[Data] ), "dd/mm/yy" ) & " – "
            & FORMAT ( MAX ( Dim_Calendario[Data] ), "dd/mm/yy" ) & _ChipFim
    )
VAR _Planta =
    IF (
        ISFILTERED ( Dim_Fabrica[Nome_Fabrica] ),
        _ChipIni & "Planta" & _ChipMeio
            & CONCATENATEX ( VALUES ( Dim_Fabrica[Nome_Fabrica] ), Dim_Fabrica[Nome_Fabrica], ", " ) & _ChipFim
    )
VAR _Linha =
    IF (
        ISFILTERED ( Dim_Produto[Categoria] ),
        _ChipIni & "Linha" & _ChipMeio
            & CONCATENATEX ( VALUES ( Dim_Produto[Categoria] ), Dim_Produto[Categoria], ", " ) & _ChipFim
    )
VAR _Produto =
    IF (
        ISFILTERED ( Dim_Produto[Nome_Produto] ),
        _ChipIni & "Produto" & _ChipMeio
            & CONCATENATEX ( VALUES ( Dim_Produto[Nome_Produto] ), Dim_Produto[Nome_Produto], ", " ) & _ChipFim
    )
VAR _Cliente =
    IF (
        ISFILTERED ( Dim_Cliente[Nome_Cliente] ),
        _ChipIni & "Cliente" & _ChipMeio
            & CONCATENATEX ( VALUES ( Dim_Cliente[Nome_Cliente] ), Dim_Cliente[Nome_Cliente], ", " ) & _ChipFim
    )
VAR _Todos = _Periodo & _Planta & _Linha & _Produto & _Cliente
RETURN
    "<div style='padding:4px 0'>"
        & IF (
            _Todos = "",
            "<span style='color:#7A8B99;font:400 11px Segoe UI'>Sem filtros ativos · visão consolidada da operação</span>",
            _Todos
        )
        & "</div>"

Qtd Filtros Ativos =
INT ( ISFILTERED ( Dim_Calendario ) )
    + INT ( ISFILTERED ( Dim_Fabrica[Nome_Fabrica] ) )
    + INT ( ISFILTERED ( Dim_Produto[Categoria] ) )
    + INT ( ISFILTERED ( Dim_Produto[Nome_Produto] ) )
    + INT ( ISFILTERED ( Dim_Cliente[Nome_Cliente] ) )
```

---

## 6. New Card Visual com imagens dos produtos

### 6.1 Hospedagem das imagens (resolver primeiro)

As URLs atuais da base são placeholders. Opções, em ordem de preferência:

| Opção | Como | Prós | Contras |
|---|---|---|---|
| **A. Repositório público / CDN** | Subir `assets/produtos/*.webp` num repositório público e usar a URL *raw* | Versionado junto do projeto, zero custo | Repositório precisa ser público; imagens corporativas expostas |
| **B. Blob Storage com acesso anônimo de leitura** | Container `produtos` no Azure Storage | Controle corporativo, rápido | Requer conta Azure |
| **C. Base64 embutido** | Coluna com `data:image/webp;base64,...` | Funciona offline e em ambientes sem internet | Limite de ~32 mil caracteres por valor de coluna → imagem precisa ter < ~20 KB |

SharePoint/OneDrive **não** funciona para URL de imagem: exige autenticação que o visual não repassa.

**Preparação das 6 fotos** (as fornecidas no projeto — 7Ah, 60Ah, 72Ah, 150Ah, 220Ah; a 9Ah AGM ainda não tem foto):
- Recorte com fundo transparente, 600 × 600 px, WebP qualidade 80. Para Base64, 240 × 240 px.
- Nome padronizado: `voltex_<id>_<modelo>.webp` (ex.: `voltex_3_60ah.webp`).

Script para gerar a coluna Base64 (rodar uma vez e colar em `Dim_Produto` ou ler via Power Query de CSV):

```python
# gerar_base64.py — python -m pip install pillow
import base64, io, csv, pathlib
from PIL import Image

pasta = pathlib.Path("assets/produtos")
linhas = []
for arq in sorted(pasta.glob("voltex_*.*")):
    id_produto = int(arq.stem.split("_")[1])
    img = Image.open(arq).convert("RGBA")
    img.thumbnail((240, 240))
    buf = io.BytesIO()
    img.save(buf, format="WEBP", quality=75)
    b64 = base64.b64encode(buf.getvalue()).decode()
    uri = f"data:image/webp;base64,{b64}"
    assert len(uri) < 32000, f"{arq.name}: {len(uri)} caracteres — reduza a imagem"
    linhas.append({"ID_Produto": id_produto, "Imagem_Base64": uri})

with open("produto_imagens.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=["ID_Produto", "Imagem_Base64"])
    w.writeheader(); w.writerows(linhas)
print(f"{len(linhas)} imagens convertidas")
```

### 6.2 Catálogo vivo com o New Card Visual

Objetivo: uma fileira de 6 cartões, um por SKU, com foto, nome e KPIs — sem caixas: o cartão não tem fundo, a foto "pousa" sobre o grafite.

1. Insira o **Cartão (novo)** e arraste `Faturamento Bruto`, `Margem Bruta %`, `Scrap Rate` e `Giro Anualizado` para **Dados**.
2. Arraste `Dim_Produto[Nome_Produto]` para **Categorias** (o cartão novo passa a repetir uma instância por produto). Layout: **1 linha × 6 colunas**, espaçamento 16 px.
3. **Formatar › Cartões › Imagens**: ligar; **Tipo de imagem: URL da imagem**; **fx → medida `URL Produto`**. Posição: **acima**; tamanho 96 px; ajuste *Ajustar*.
4. **Cartões › Plano de fundo**: desligado. **Borda**: só a inferior, 2 px, cor `fx → Cor Scrap` — a "régua de status" é o único elemento colorido do cartão.
5. **Valor de referência** em cada KPI: `Scrap Rate` com rótulo "Meta 1,5%"; cor do valor `fx → Cor Scrap`. Repetir a lógica para Giro (`Cor Giro`).
6. **Cartão de destaque**: o Faturamento Bruto em 18 pt branco; demais KPIs 10 pt, rótulos `#7A8B99`.
7. **Tooltip**: apontar para a página `TT_Produto` (seção 7).

Mesmo padrão em **matriz**: `Dim_Produto[URL_Imagem]` (ou `Imagem_Base64`) como primeira coluna; **Formatar › Imagem de grade › Altura 56 px**.

---

## 7. Tooltips dinâmicos de página

### 7.1 `TT_Produto` (abas Executivo, Qualidade, Estoque)

1. Nova página → **Informações da página › Permitir uso como dica de ferramenta: Ligado**; **Tela › Tipo: Dica de ferramenta**, tamanho personalizado **360 × 260**.
2. Plano de fundo `bg_tooltip.png` (grafite com divisória de 1 px a 120 px da esquerda). Página oculta.
3. **Campos de dica de ferramenta**: `Dim_Produto[Nome_Produto]` — a página é filtrada pelo produto sob o cursor.
4. Conteúdo:

```text
┌──────────────────────────────────────────────┐
│ ┌────────┐  BATERIA CARRO 60AH               │
│ │ FOTO   │  Linha Leve · Curva A             │
│ │ 104px  │                                   │
│ └────────┘  R$ 276,5 mi   ▲ margem 58,7%     │
│ ──────────────────────────────────────────── │
│  Vendas 24 meses  ▁▂▃▃▅▄▆▅▇▆▇█  (sparkline)   │
│ ──────────────────────────────────────────── │
│  Scrap 2,6% ●     Giro 16,1×      Reversa 20%│
└──────────────────────────────────────────────┘
```

   - **Foto**: New Card Visual só com imagem (`URL Produto`), sem valor.
   - **Título**: HTML Content com nome + categoria.
   - **Sparkline**: no New Card, o próprio recurso **Minigráfico** (Dados › Sparkline: `Qtd Vendida` por `Dim_Calendario[Ano_Mes]`), linha `#FF8C00` 2 px, sem eixo. Alternativa: gráfico de linhas nativo com eixos, grade e rótulos desligados e marcador só no último ponto.
   - **Rodapé**: cartão novo com 3 KPIs e cor condicional.

5. No visual de origem: **Formatar › Propriedades gerais › Dicas de ferramenta › Tipo: Página do relatório › Página: TT_Produto**.

### 7.2 `TT_Planta` (abas OEE, Manutenção, Custos)

Mesmo esqueleto; campo de dica `Dim_Fabrica[Nome_Fabrica]`. Conteúdo: decomposição do OEE em barras horizontais finas (Disponibilidade, Performance com nota "não medido", Qualidade), MTBF vs. 300 h, MTTR vs. 2,5 h, e foto da planta/linha (as imagens de chão de fábrica do Media Kit).

---

## 8. Deneb / Vega-Lite

Deneb (AppSource) recebe os campos arrastados como `dataset`. Cada campo vira uma coluna com o nome exibido no painel — renomeie os campos no visual para nomes curtos sem espaço (`Data`, `Planta`, `Disp`).

### 8.1 Mapa de calor de Disponibilidade — planta × semana

**Por que Deneb:** a matriz nativa com formatação condicional vira uma planilha colorida ilegível com 208 semanas; aqui cada célula é um retângulo de 4 px e o olho lê padrão, não número.

Campos: `Dim_Calendario[Data]` (renomear `Data`), `Dim_Fabrica[Nome_Fabrica]` (`Planta`), `[Disponibilidade]` (`Disp`).

```json
{
  "$schema": "https://vega.github.io/schema/vega-lite/v5.json",
  "data": { "name": "dataset" },
  "background": "transparent",
  "transform": [
    { "timeUnit": "yearweek", "field": "Data", "as": "Semana" },
    { "aggregate": [{ "op": "mean", "field": "Disp", "as": "DispSemana" }], "groupby": ["Semana", "Planta"] }
  ],
  "mark": { "type": "rect", "cornerRadius": 1 },
  "encoding": {
    "x": {
      "field": "Semana", "type": "temporal", "timeUnit": "yearweek",
      "axis": { "format": "%b/%y", "labelColor": "#7A8B99", "domain": false, "ticks": false, "grid": false, "title": null }
    },
    "y": {
      "field": "Planta", "type": "nominal",
      "axis": { "labelColor": "#FFFFFF", "labelFontSize": 11, "domain": false, "ticks": false, "title": null, "labelLimit": 220 }
    },
    "color": {
      "field": "DispSemana", "type": "quantitative",
      "scale": { "domain": [0.85, 0.92, 1], "range": ["#E5484D", "#FF8C00", "#2C363F"] },
      "legend": { "title": "Disponibilidade", "titleColor": "#7A8B99", "labelColor": "#7A8B99", "format": ".0%", "orient": "top" }
    },
    "tooltip": [
      { "field": "Planta" },
      { "field": "Semana", "type": "temporal", "timeUnit": "yearweek", "format": "%d/%m/%Y", "title": "Semana de" },
      { "field": "DispSemana", "format": ".1%", "title": "Disponibilidade" }
    ]
  },
  "config": { "view": { "stroke": null } }
}
```

> A média de `Disp` por dia dentro da semana é aproximação aceitável para *padrão visual*; o número exato vive no tooltip de página. Escala invertida de propósito: semana boa some no grafite, semana ruim acende.

### 8.2 Bullet chart de Scrap por SKU vs. meta

**Por que Deneb:** a meta tem faixas (verde < 1,5 %, alerta até 3 %, crítico acima) e o visual nativo de medidor só mostra um KPI por vez.

Campos: `Dim_Produto[Nome_Produto]` (`Produto`), `[Scrap Rate]` (`Scrap`), `[Custo do Refugo R$]` (`Custo`).

```json
{
  "$schema": "https://vega.github.io/schema/vega-lite/v5.json",
  "data": { "name": "dataset" },
  "background": "transparent",
  "height": { "step": 34 },
  "encoding": {
    "y": { "field": "Produto", "type": "nominal", "sort": { "field": "Custo", "order": "descending" },
           "axis": { "labelColor": "#FFFFFF", "labelLimit": 260, "domain": false, "ticks": false, "title": null } }
  },
  "layer": [
    { "mark": { "type": "rect", "color": "#2FBF71", "opacity": 0.12, "height": 22 },
      "encoding": { "x": { "datum": 0 }, "x2": { "datum": 0.015 } } },
    { "mark": { "type": "rect", "color": "#FF8C00", "opacity": 0.12, "height": 22 },
      "encoding": { "x": { "datum": 0.015 }, "x2": { "datum": 0.03 } } },
    { "mark": { "type": "bar", "height": 6, "cornerRadiusEnd": 3 },
      "encoding": {
        "x": { "field": "Scrap", "type": "quantitative", "scale": { "domain": [0, 0.035] },
               "axis": { "format": ".1%", "labelColor": "#7A8B99", "grid": false, "domain": false, "title": null } },
        "color": { "condition": { "test": "datum.Scrap < 0.015", "value": "#2FBF71" }, "value": "#FF8C00" },
        "tooltip": [ { "field": "Produto" }, { "field": "Scrap", "format": ".2%" }, { "field": "Custo", "format": ",.0f", "title": "Custo do refugo (R$)" } ]
      } },
    { "mark": { "type": "tick", "color": "#FFFFFF", "thickness": 2, "size": 22 },
      "encoding": { "x": { "datum": 0.015 } } },
    { "mark": { "type": "text", "align": "left", "dx": 6, "color": "#7A8B99", "fontSize": 10 },
      "encoding": { "x": { "field": "Scrap" }, "text": { "field": "Custo", "format": ",.0f" } } }
  ],
  "config": { "view": { "stroke": null } }
}
```

### 8.3 Gantt de paradas de máquina

**Por que Deneb:** não existe Gantt nativo, e os visuais de AppSource de Gantt não aceitam o tema VoltEx nem cor por severidade.

**Pré-requisito honesto:** a `Fato_Producao` atual registra *horas paradas por turno*, não *início e fim* de cada evento. Gantt exige um evento por linha. A recomendação é criar `Fato_Paradas` a partir do MES/apontamento de manutenção:

| Coluna | Tipo |
|---|---|
| `ID_Parada` | int |
| `ID_Fabrica` | int → Dim_Fabrica |
| `Equipamento` | texto (Injetora 03, Forno de chumbo 01…) |
| `Inicio` / `Fim` | datetime |
| `Tipo` | Corretiva / Preventiva / Preditiva / Setup |

Relacionamento: `Dim_Calendario[Data]` → `Fato_Paradas[Data_Inicio]` (coluna de data derivada de `Inicio`). Campos no Deneb: `Equipamento`, `Inicio`, `Fim`, `Tipo`, `Duracao_h`.

```json
{
  "$schema": "https://vega.github.io/schema/vega-lite/v5.json",
  "data": { "name": "dataset" },
  "background": "transparent",
  "height": { "step": 26 },
  "mark": { "type": "bar", "cornerRadius": 2, "height": 14 },
  "encoding": {
    "y": { "field": "Equipamento", "type": "nominal",
           "axis": { "labelColor": "#FFFFFF", "domain": false, "ticks": false, "title": null } },
    "x":  { "field": "Inicio", "type": "temporal",
            "axis": { "format": "%d/%m %Hh", "labelColor": "#7A8B99", "gridColor": "#3A4651", "gridDash": [2, 3], "domain": false, "title": null } },
    "x2": { "field": "Fim" },
    "color": {
      "field": "Tipo", "type": "nominal",
      "scale": { "domain": ["Corretiva", "Preventiva", "Preditiva", "Setup"],
                 "range":  ["#E5484D", "#7A8B99", "#2FBF71", "#4E5D6A"] },
      "legend": { "orient": "top", "labelColor": "#7A8B99", "title": null }
    },
    "tooltip": [
      { "field": "Equipamento" }, { "field": "Tipo" },
      { "field": "Inicio", "type": "temporal", "format": "%d/%m/%Y %H:%M" },
      { "field": "Duracao_h", "format": ".1f", "title": "Duração (h)" }
    ]
  },
  "config": { "view": { "stroke": null } }
}
```

Laranja fica de fora da legenda de propósito: ele é reservado para destacar a parada selecionada via interação cruzada (`pbiColor` / seleção do Deneb), não para um tipo.

---

## 9. Checklist de publicação

**Modelo**
- [ ] 7 relacionamentos 1:\* direção única; nenhum bidirecional.
- [ ] `Dim_Calendario` marcada como tabela de datas.
- [ ] Medidas validadas contra os números de referência abaixo.
- [ ] `Taxa Logistica Reversa` retorna vazio com filtro de planta.

**Números de referência para teste (período completo 2022–2025, sem filtros)**

| Medida | Valor esperado |
|---|---:|
| Qtd Vendida | 1.525.264 |
| Faturamento Bruto | R$ 609,6 mi |
| OEE (D × Q) | 91,8 % |
| Disponibilidade | 94,2 % |
| Scrap Rate | 2,63 % |
| MTTR | 2,15 h |
| MTBF | 35,2 h |
| Taxa Logística Reversa | 41,8 % |
| Giro Anualizado (2025) | 7,6× |
| Custo do Refugo | R$ 33,3 mi |

**Relatório**
- [ ] Tema importado; nenhum visual com fundo/borda/sombra.
- [ ] Fundos 2560 × 1440 aplicados com "Ajustar".
- [ ] Indicadores do drawer com **Dados** desmarcado.
- [ ] Slicers sincronizados nas 7 abas.
- [ ] URLs de imagem respondendo (abrir cada uma no navegador anônimo).
- [ ] Tooltips de página atribuídos aos visuais de origem.
- [ ] Teste no serviço + app móvel (layout móvel das 2 primeiras abas).

---
---

# PARTE 2 — CASE STUDY PARA PORTFÓLIO

> **Nota de transparência para a publicação:** a VoltEx é uma empresa-modelo e a base de dados é simulada para fins de portfólio. Os números abaixo são reais *para essa base* — foram calculados dela, não inventados para o texto. Os "resultados" são apresentados como diagnóstico e oportunidade quantificada, que é exatamente o que um projeto desses entrega no primeiro ciclo. Recomendo manter esta nota (ou uma versão curta dela) no post: recrutadores e gestores industriais valorizam quem separa simulação de resultado de cliente.

---

## ⚡ R$ 14 milhões escondidos no refugo: como um Power BI industrial expôs onde a fábrica perde dinheiro

### Descrição curta

Projeto de BI industrial ponta a ponta para uma fabricante de baterias automotivas: modelo Star Schema com quatro fatos, KPIs de chão de fábrica (OEE, MTBF, MTTR, Scrap) cruzados com vendas e logística reversa, e uma interface dark mode sem "caixinhas". O resultado é um painel que não só mede a operação — ele aponta em reais onde está a perda e qual linha atacar primeiro.

---

### Contexto e Objetivos

A VoltEx Baterias opera três plantas — Campinas (linha leve), Joinville (linha pesada e reciclagem) e Manaus (motos) — com seis SKUs que vão de uma bateria de moto de R$ 90 a uma de trator/ônibus de R$ 1.200. A estratégia da empresa se apoia em duas apostas: rentabilidade na linha pesada e o **EcoCiclo**, programa de logística reversa que devolve chumbo reciclado à fundição e promete reduzir o custo do produto vendido.

O problema era o de toda indústria que cresce sem camada analítica: os dados de produção, vendas, coleta de cascos e estoque viviam em silos. Ninguém conseguia responder, na mesma tela, perguntas como:

- A fábrica está perto da classe mundial de OEE ou só parece estar?
- Quanto custa, em reais, cada ponto percentual de refugo — e em qual linha?
- O EcoCiclo está fechando o ciclo ou só existe no catálogo?
- O estoque está girando onde importa?

**Objetivos do projeto**
1. Unificar quatro processos (produzir, vender, recolher, estocar) em um único modelo semântico confiável.
2. Traduzir as metas da Matriz de KPIs em semáforos objetivos, sem interpretação.
3. Converter indicadores técnicos em impacto financeiro para a diretoria.
4. Entregar uma interface à altura da marca — industrial, escura, precisa.

---

### Metodologia / Abordagem

**Dados.** Quatro anos de operação (2022–2025): 15 mil linhas de venda, 10 mil apontamentos de produção por turno, 8 mil coletas de cascos e 864 fotografias mensais de estoque, apoiadas por dimensões de Produto, Cliente e Fábrica.

**Arquitetura.** Star Schema rigoroso com um único calendário relacionado de forma ativa às quatro fatos — possível porque as fatos não conversam entre si, o que elimina ambiguidade e dispensa `USERELATIONSHIP`. Projeto em formato `.pbip` com TMDL e PBIR, versionado em Git, com modelo e relatório evoluindo em branches separadas.

**KPIs.** Seis indicadores com meta explícita:

| KPI | Meta | Como foi calculado |
|---|---|---|
| OEE | > 85 % | Disponibilidade × Qualidade (Performance declarada como não medida) |
| MTBF | > 300 h | Horas operando ÷ nº de falhas |
| MTTR | < 2,5 h | Horas paradas ÷ nº de falhas |
| Scrap Rate | < 1,5 % | Refugo ÷ produção |
| Giro de Estoque | > 6×/ano | Vendas ÷ estoque médio, anualizado |
| Logística Reversa | > 75 % | Cascos recolhidos ÷ baterias vendidas |

Todos calculados como razão de somas — nunca média de percentuais —, e o estoque tratado como medida semi-aditiva (último snapshot do período).

**UX/UI.** Fundos desenhados fora do Power BI carregam toda a estrutura visual; os visuais flutuam sem moldura. Laranja Energético só aparece onde há sinal. Painel de filtros retrátil em estilo *drawer* com trilha de filtros ativos em HTML, catálogo de produtos com fotos reais no New Card Visual, tooltips de página com sparkline e foto, e Deneb/Vega-Lite para mapa de calor de disponibilidade e bullet chart de refugo contra meta.

---

### Desafios e Aprendizados

**1. Cruzar logística reversa com vendas sem mentir.** A tabela de coleta de cascos não tem o grão de planta. Um filtro por fábrica inflaria a taxa de retorno em cerca de três vezes — numerador inteiro sobre denominador parcial. A solução foi uma medida que se recusa a responder nesse contexto e avisa o usuário no próprio card. Aprendizado: às vezes o indicador mais confiável é o que sabe quando ficar em branco.

**2. OEE sem Performance.** A base não traz capacidade teórica por linha. Em vez de inventar uma, o OEE foi calculado com os dois fatores mensuráveis e a limitação foi exposta no tooltip. Um OEE de 91,8 % sem Performance não é comparável a benchmark — e o painel diz isso.

**3. Estoque é foto, não filme.** Somar fechamentos mensais gera números absurdos. O giro precisou de estoque semi-aditivo e de anualização, senão filtrar um trimestre derrubaria o indicador para um quarto do valor real.

**4. Dados que contradizem o plano.** O Plano de Negócios define plantas especializadas, mas os dados mostram as três plantas produzindo os seis SKUs. Em vez de "limpar" a contradição, ela virou achado de governança — o tipo de coisa que só aparece quando alguém olha o dado de verdade.

---

### Insights e Descobertas

**A linha pesada carrega a empresa — com 1 em cada 5 baterias.**
Ela responde por 19,8 % das unidades vendidas, mas 46,6 % do faturamento bruto. A margem percentual é a menor do portfólio (54,9 % contra 58,7 % da leve e 60,9 % das motos), porém cada bateria pesada deixa R$ 516 de lucro bruto, contra R$ 212 da leve e R$ 65 das motos. Margem percentual alta não paga a conta; margem em reais por unidade, sim.

**O OEE está bom. A confiabilidade das máquinas, não.**
OEE de 91,8 % e MTTR de 2,15 h — dentro da meta. Mas o MTBF é de 35 horas, contra uma meta de 300. A equipe de manutenção conserta rápido; as máquinas é que quebram demais. É o retrato de uma operação que vive de manutenção corretiva e precisa migrar para preventiva e preditiva.

**O refugo é o maior vazamento financeiro da fábrica.**
Scrap de 2,63 %, quase o dobro da meta de 1,5 % e estável nos quatro anos. São 144 mil baterias refugadas, ou R$ 33,3 milhões a custo de fabricação. A linha pesada concentra 68 % desse custo — porque cada unidade perdida ali vale muito mais.

**O EcoCiclo não está fechando o ciclo — e o vazamento está no produto mais vendido.**
A taxa de retorno de cascos é de 41,8 %, longe dos 75 %, e caiu de 45,0 % em 2022 para 40,8 % em 2025. O pior desempenho é justamente da bateria 60Ah, campeã de vendas: só 20 % dos cascos voltam. Já a 220Ah retorna mais cascos do que vende (138 %), sinal de que o canal de frotas pesadas funciona e o varejo leve não.

**O giro geral é saudável; o capital parado está no item mais caro.**
Giro de 7,6× ao ano, acima da meta. Mas a 220Ah gira só 2,3× e a 9Ah AGM 4,5×, enquanto a 60Ah gira 16×. O estoque mais caro do portfólio é o que mais fica parado.

**A sazonalidade de inverno não aparece nas vendas.**
O plano concentra estoque entre maio e agosto, mas o maior mês de vendas é janeiro e a variação mensal é discreta. Ou a premissa sazonal está errada, ou a campanha de inverno não está convertendo — as duas hipóteses merecem investigação.

---

### Resultados e Impacto

O primeiro ciclo de um projeto de BI industrial entrega diagnóstico com valor em reais. Foi o que este painel fez:

| Frente | Linha de base medida | Meta | Oportunidade quantificada |
|---|---|---|---|
| **Refugo** | 2,63 % · R$ 33,3 mi em 4 anos | < 1,5 % | **≈ R$ 14,3 mi** em 4 anos (≈ R$ 3,6 mi/ano) se a meta fosse atingida |
| **Confiabilidade** | MTBF 35 h · 4.288 falhas | > 300 h | Plano preventivo priorizado pelas plantas e semanas mais críticas do mapa de calor |
| **EcoCiclo** | 41,8 % de retorno · 637 mil cascos | > 75 % | **≈ 507 mil cascos** adicionais no período (≈ 127 mil/ano) — com foco no varejo da linha leve 60Ah |
| **Estoque** | Giro 7,6× · 220Ah a 2,3× | > 6× | Revisão de política de estoque dos SKUs pesados de alto ticket |
| **Governança** | Plantas sem especialização nos dados | — | Achado levado à operação para validação do apontamento |

**Impacto para o negócio**
- **Para a diretoria:** perdas operacionais deixaram de ser percentuais abstratos e viraram reais por linha de produto.
- **Para a manutenção:** a conversa mudou de "consertar mais rápido" para "quebrar menos".
- **Para o ESG:** o EcoCiclo ganhou um indicador honesto, com o gargalo identificado por produto e canal — matéria-prima para o Certificado de Destinação de Resíduos dos frotistas.
- **Para a equipe de dados:** um projeto `.pbip` versionado, com medidas documentadas e testáveis, pronto para crescer com novas fontes como apontamento de paradas e capacidade por linha.

---

> **Stack:** Power BI Desktop (.pbip · TMDL · PBIR) · Power Query M · DAX · Deneb / Vega-Lite · HTML Content · Git
> **Competências:** Modelagem dimensional · KPIs industriais (OEE, MTBF, MTTR) · Economia circular / ESG · Data storytelling · UI/UX para BI

`#PowerBI` `#BusinessIntelligence` `#Industria40` `#OEE` `#DataViz` `#ESG` `#EconomiaCircular` `#DAX`
