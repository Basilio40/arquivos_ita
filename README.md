# arquivos_ita — repositório de dados do dashboard ITALINEA

Fonte única de dados consumida em runtime pelo app **dash_moema**
(github.com/Basilio40/dash_moema, deploy na Railway).

## Como os dados chegam aqui

Diariamente às 08:00 (Agendador de Tarefas → `ITALINEA_AtualizacaoDados`)
o `atualizar_dados_italinea.bat` roda: robôs Focco baixam os .xlsx →
scripts Python geram os JSONs → **etapa [5/5]**
`scripts/sincronizar_dados_github.py` (na raiz ITALINEA da máquina local)
copia tudo para cá e faz commit+push (só quando algo mudou).

## Layout

```
arquivos_ita/
├── indice.json          # gerado: caminho, origem, bytes, mtime, sha256
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
(`realizado/HISTORICO`, `HISTORICO_PREV`) são ignoradas.

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

## Limites conhecidos

- `latest` é escolhido por mtime: um download corrompido/parcial mais recente
  que o bom substitui o bom (hoje o robô já valida conteúdo antes de salvar).
- O índice registra sha256 para auditoria, mas nada valida o conteúdo das
  planilhas no sync.
- Variante futura possível: GitHub Action neste repo que commita os JSONs no
  repo do app e auto-redeploya a Railway (elimina o carregador runtime, mas
  amarra dado a deploy).
