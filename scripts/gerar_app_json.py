#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Regenera os JSONs de app/ a partir dos .xlsx versionados neste repo.

É o que a GitHub Action (.github/workflows/gerar-app-json.yml) executa a cada
push em xlsx/** — e pode ser rodado manualmente para teste:

    python scripts/gerar_app_json.py

Etapas (espelham o passo [4/5] do atualizar_dados_italinea.bat local):
  1. dashboard.json + dre-details.json  ← transform_dre (xlsx/dre + xlsx/carteira_mensal)
  2. filtro do dre-details              ← remove o item 4.03.01.01 (como o wrapper local)
  3. projecao-contratos.json            ← gerar_projecao_contratos (xlsx/contratos)
  4. comissoes.json                     ← gerar_comissoes         (xlsx/contratos)
  5. fluxo-caixa-previsao.json          ← gerar_previsao_fluxo    (xlsx/fluxo_caixa/previsao)
  6. fluxo-caixa-realizado.json         ← gerar_realizado_fluxo   (xlsx/fluxo_caixa/realizado)
  7. validação de shape mínimo + atualização das entradas app/* do indice.json

O ano é derivado do DRE mais recente presente em xlsx/dre (nada de hardcoded
de 2026 aqui); o rótulo do período ("Janeiro a {mês} / {ano}") também.

Qualquer falha → exit 1 (a Action falha e NADA é commitado: app/ permanece
no último estado bom).
"""
import hashlib
import json
import os
import re
import sys
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DRE_DIR = os.path.join(BASE, 'xlsx', 'dre')
APP_DIR = os.path.join(BASE, 'app')

EMPRESA = 'ITALINEA - DECORA INTERIORES LTDA'

# Item removido em definitivo das tabelas de detalhe (valores continuam
# subtraídos dos totais pela UI — ver src/lib/despesas.ts do dash_moema).
ITENS_REMOVIDOS_DETALHE = {'4.03.01.01'}

MESES_EXTENSO = {
    'Jan': 'Janeiro', 'Fev': 'Fevereiro', 'Mar': 'Março', 'Abr': 'Abril',
    'Mai': 'Maio', 'Jun': 'Junho', 'Jul': 'Julho', 'Ago': 'Agosto',
    'Set': 'Setembro', 'Out': 'Outubro', 'Nov': 'Novembro', 'Dez': 'Dezembro',
}

ARQUIVOS_APP = [
    'dashboard.json',
    'dre-details.json',
    'projecao-contratos.json',
    'comissoes.json',
    'fluxo-caixa-previsao.json',
    'fluxo-caixa-realizado.json',
]


def detectar_ano():
    """Ano do DRE mais recente em xlsx/dre (fallback: ano corrente)."""
    anos = []
    if os.path.isdir(DRE_DIR):
        for nome in os.listdir(DRE_DIR):
            m = re.match(r'DRE_\d{2}_(\d{4})\.xlsx$', nome)
            if m:
                anos.append(int(m.group(1)))
    return max(anos) if anos else datetime.now().year


def validar(nome, valor):
    """Shape mínimo — o suficiente para barrar JSON vazio/corrompido."""
    if nome == 'dashboard.json':
        ok = (isinstance(valor, dict)
              and isinstance(valor.get('dreMonth'), list) and valor['dreMonth']
              and isinstance(valor.get('dreGroups'), list))
    elif nome in ('dre-details.json', 'fluxo-caixa-previsao.json',
                  'fluxo-caixa-realizado.json'):
        ok = isinstance(valor, list)  # previsão pode razoavelmente vir vazia
    else:  # projecao-contratos / comissoes
        ok = isinstance(valor, dict) and isinstance(valor.get('series'), list)
    if not ok:
        raise ValueError(f'{nome}: forma inesperada após a geração')
    print(f'  ok: {nome}')


def atualizar_indice():
    """Atualiza só as entradas app/* do indice.json (xlsx/* são do sync diário).

    Entrada cujo sha256 não mudou é preservada verbatim — o índice só muda
    quando o CONTEÚDO muda (mesma política do sincronizador, evita commit vazio).
    """
    caminho_indice = os.path.join(BASE, 'indice.json')
    existentes = {}
    if os.path.exists(caminho_indice):
        with open(caminho_indice, encoding='utf-8') as f:
            existentes = {e['caminho']: e
                          for e in json.load(f).get('arquivos', [])}

    agora = datetime.now().strftime('%Y-%m-%d %H:%M')
    for nome in ARQUIVOS_APP:
        caminho = os.path.join(APP_DIR, nome)
        rel = f'app/{nome}'
        sha = hashlib.sha256(open(caminho, 'rb').read()).hexdigest()
        antiga = existentes.get(rel)
        if antiga and antiga.get('sha256') == sha:
            continue  # conteúdo igual → preserva entrada (e o mtime antigo)
        existentes[rel] = {
            'caminho': rel,
            'origem': 'scripts/gerar_app_json.py (GitHub Actions)',
            'bytes': os.path.getsize(caminho),
            'modificado': agora,
            'sha256': sha,
        }

    indice = sorted(existentes.values(), key=lambda e: e['caminho'])
    with open(caminho_indice, 'w', encoding='utf-8') as f:
        json.dump({'arquivos': indice}, f, ensure_ascii=False, indent=2)
        f.write('\n')
    print(f'  indice.json: {len(indice)} entradas')


def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

    import transform_dre
    import gerar_projecao_contratos
    import gerar_comissoes
    import gerar_previsao_fluxo
    import gerar_realizado_fluxo

    ano = detectar_ano()
    transform_dre.ANO = ano
    meses = transform_dre.detect_months(DRE_DIR)
    if not meses:
        raise SystemExit(f'nenhum DRE_{{mm}}_{ano}.xlsx encontrado em xlsx/dre')
    ultimo = MESES_EXTENSO[meses[-1]]
    periodo = f'Janeiro a {ultimo} / {ano}'
    print(f'=== Gerando app/*.json (ano {ano}: {periodo}) ===')

    # 1) dashboard + dre-details a partir do DRE com override da carteira
    transform_dre.main([
        '--source-dir', DRE_DIR,
        '--carteira-dir', os.path.join(BASE, 'xlsx', 'carteira_mensal'),
        '--output-dir', APP_DIR,
        '--ano', str(ano),
        '--empresa', EMPRESA,
        '--periodo', periodo,
    ])

    # 2) filtro do dre-details (mesma regra do wrapper local do guide-main)
    detalhe_path = os.path.join(APP_DIR, 'dre-details.json')
    with open(detalhe_path, encoding='utf-8') as f:
        itens = json.load(f)
    filtrados = [it for it in itens
                 if it.get('itemCod') not in ITENS_REMOVIDOS_DETALHE]
    if len(filtrados) != len(itens):
        with open(detalhe_path, 'w', encoding='utf-8') as f:
            json.dump(filtrados, f, ensure_ascii=False, indent=2)
            f.write('\n')
        print(f'  dre-details: removido(s) '
              f'{len(itens) - len(filtrados)} item(s) de '
              f'{sorted(ITENS_REMOVIDOS_DETALHE)}')

    # 3-6) contratos, comissões e fluxo de caixa
    gerar_projecao_contratos.main()
    gerar_comissoes.main()
    gerar_previsao_fluxo.main()
    gerar_realizado_fluxo.main()

    # 7) validação + índice
    print('--- validação ---')
    for nome in ARQUIVOS_APP:
        with open(os.path.join(APP_DIR, nome), encoding='utf-8') as f:
            validar(nome, json.load(f))
    atualizar_indice()
    print('=== concluído ===')
    return 0


if __name__ == '__main__':
    sys.exit(main())
