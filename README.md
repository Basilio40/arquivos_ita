# arquivos_ita — repositório de dados do dashboard ITALINEA

Fonte única de dados consumida em runtime pelo app **dash_moema**
(github.com/Basilio40/dash_moema, deploy na Railway).

## Como os dados chegam aqui

**Via robô local (automático):** diariamente às 08:00 (Agendador de Tarefas →
`ITALINEA_AtualizacaoDados`) o `atualizar_dados_italinea.bat` roda: robôs
Focco baixam os .xlsx → **etapa [5/5]**
`scripts/sincronizar_dados_github.py` (na raiz ITALINEA da máquina local)
copia as **planilhas** para cá e faz commit+push (só quando algo mudou).

**Via GitHub Action (a partir daí):** todo push que mexe em `xlsx/**`
dispara o workflow `gerar-app-json`, que roda `scripts/gerar_app_json.py`
e regenera os JSONs de `app/` a partir dos xlsx do próprio repo —
commitando só se o conteúdo mudou.

**Via upload manual:** subir/atualizar um .xlsx direto pelo site do GitHub
(ex.: `xlsx/carteira_mensal/Nov_2026.xlsx`) cai no mesmo caminho — a Action
regenera o `app/` e o dashboard atualiza sozinho (cache de 5 min).
O `app/` é **propriedade do CI**: a sincronização local não o escreve.

## Layout

```
arquivos_ita/
├── indice.json          # gerado: caminho, origem, bytes, mtime, sha256
│                        #   (xlsx/* → sync local; app/* → GitHub Action)
├── .github/workflows/gerar-app-json.yml   # CI que regenera app/ dos xlsx
├── scripts/             # geradores rodados pelo CI (cópias vendorizadas dos
│   │                    # scripts do pro-dash-guide-main/pro-dash-mp, com
│   │                    # caminhos apontando para xlsx/ → app/)
│   ├── gerar_app_json.py          # orquestrador (ano/período auto-detectados)
│   ├── transform_dre.py           # dashboard.json + dre-details.json
│   ├── gerar_projecao_contratos.py + gerar_comissoes.py   (xlsx/contratos)
│   └── gerar_{previsao,realizado}_fluxo.py               (xlsx/fluxo_caixa)
├── xlsx/                # planilhas brutas (origem: pastas loja_18314_*)
│   ├── carteira_mensal/     Mmm_AAAA.xlsx            (série completa)
│   ├── dre/                 DRE_MM_AAAA.xlsx          (série completa)
│   ├── contratos/           latest.xlsx + nomeados    (hash → mais recente)
│   ├── projetos/            latest.xlsx               (wbpvenprojetos)
│   ├── movimentacao/        latest.xlsx               (wbpfinmovimentacaorpt)
│   └── fluxo_caixa/
│       ├── realizado/<ANO>/    série nomeada
│       ├── previsao/           série nomeada
│       └── {previsao,realizado}_latest.xlsx  (resíduos Hifn… com hash)
└── app/                 # JSONs prontos — é isto que o dashboard lê
    ├── dashboard.json, dre-details.json, comissoes.json,
    └── fluxo-caixa-{previsao,realizado}.json, projecao-contratos.json
```

**Regra de seleção:** arquivo do Focco com hash de 32 hex no nome é o mesmo
relatório baixado N vezes — só o mais recente (mtime) de cada tipo vira
`latest.xlsx`. Séries nomeadas entram completas. Pastas-mortas
(`realizado/HISTORICO`, `HISTORICO_PREV`) são ignoradas. No CI, o
`xlsx/contratos/latest.xlsx` é preferido deterministicamente (o checkout
iguala os mtimes).

## Consumo (dash_moema)

`src/lib/dados-remotos.ts` busca `app/*.json` via
`raw.githubusercontent.com/Basilio40/arquivos_ita/main/app/<nome>` com cache
de 5 min + ETag. Repo **privado** → exige a variável `DADOS_TOKEN` na Railway
(PAT fine-grained, *Contents: Read-only*, somente este repo). Qualquer falha
(gitHub fora, token inválido, shape inesperado) degrada para o JSON estático
embutido no build — o dashboard nunca fica em branco.

## Alternativas consideradas (registro de decisão — 2026-10-11)

| Alternativa | Veredito |
|---|---|
| A — cópia manual para o repo | Descartada: morre com a motivação |
| **B — repo de dados dedicado + leitura em runtime** | **Escolhida** |
| C — dados no próprio repo do app | Descartada: dado novo exigiria redeploy |
| D — object storage (S3/R2) | Descartada: conta/credencial a mais p/ ~200 KB/dia |
| E — banco de dados (Postgres) | Descartada por ora; volta à mesa se o dashboard precisar de query/agregação sobre histórico longo |

Deixa de valer (B) se: volume crescer para dezenas de MB/dia (→ D) ou o
dashboard passar a precisar de agregação pesada (→ E).

Evolução de 2026-10-12 (ainda dentro de B): a geração dos JSON mudou da
máquina local para o **CI deste repo** — uploads manuais de xlsx passam a
alimentar o dashboard sem a máquina ligada; o pipeline local ficou
responsável só por baixar (Focco) e versionar as planilhas.

## Limites conhecidos

- `latest` é escolhido por mtime: um download corrompido/parcial mais recente
  que o bom substitui o bom (hoje o robô já valida conteúdo antes de salvar).
- O índice registra sha256 para auditoria, mas nada valida o conteúdo das
  planilhas no sync.
- Os scripts em `scripts/` são cópias vendorizadas (fonte da verdade do
  cálculo continua em pro-dash-guide-main/scripts + pro-dash-mp/scripts);
  mudança de regra de negócio precisa ser replicada aqui.
- O DRE do mês em andamento entra no dashboard assim que commitado (mês
  parcial: receita/receita-margem ainda incompletas — ex.: break-even de
  Out/26 veio negativo no relatório do Focco). Para mostrar só mês fechado,
  não suba o DRE do mês corrente.
- O `gerar_realizado_fluxo` lê apenas o **topo** de `xlsx/fluxo_caixa/realizado/`
  (subpastas por ano são ignoradas) — replica o comportamento do pipeline
  local histórico.
