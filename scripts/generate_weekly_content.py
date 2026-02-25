#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import random
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MANIFEST = ROOT / 'assets' / 'gallery-manifest.json'
OUT_DIR = ROOT / 'content' / 'weekly'

CHANNELS = [
    'Facebook',
    'Instagram',
    'X',
    'YouTube',
    'TikTok',
    'LinkedIn',
    'Pinterest',
    'Google Business Profile'
]

CATEGORY_COPY = {
    'Trim': 'acabamentos de alto padrao',
    'Wainscoting': 'painelamento e boiserie sob medida',
    'Stairs': 'escadas e corrimaos com acabamento refinado',
    'Ceiling': 'tetos e detalhes arquitetonicos',
    'Decks': 'decks externos duraveis',
    'Kitchen & Vanities': 'cozinhas e vanities personalizadas',
    'Fireplaces & Bars': 'lareiras e bares planejados',
    'Outside Doors & Windows': 'portas e janelas externas com acabamento premium',
    'Pergola': 'pergolados e estruturas externas',
    'Port & Portal': 'portais e entradas com identidade'
}

HASHTAGS = [
    '#CarpintariaFina',
    '#FinishCarpentry',
    '#NoLimitCarpentry',
    '#ReformaComQualidade',
    '#DetalhesQueValorizam'
]


def load_manifest() -> dict[str, list[str]]:
    if not MANIFEST.exists():
        raise SystemExit('Manifest not found: assets/gallery-manifest.json')
    return json.loads(MANIFEST.read_text(encoding='utf-8'))


def pick_assets(manifest: dict[str, list[str]], n: int, seed: int) -> list[tuple[str, str]]:
    random.seed(seed)
    pool = []
    for cat, items in manifest.items():
        for item in items:
            pool.append((cat, item))
    if not pool:
        return []
    random.shuffle(pool)
    return pool[: min(n, len(pool))]


def caption(channel: str, cat: str, path: str, i: int) -> str:
    work = CATEGORY_COPY.get(cat, 'projeto sob medida')
    file_name = Path(path).name
    if channel == 'X':
        return (
            f'Projeto #{i}: {work}. '
            f'Acabamento real em {file_name}. '
            '#NoLimitCarpentry #FinishCarpentry'
        )
    if channel == 'YouTube':
        return (
            f'Video #{i}: Bastidores de {work}. '\
            f'Mostramos execucao, acabamento e resultado final do projeto {file_name}.'
        )
    if channel == 'Google Business Profile':
        return (
            f'Projeto #{i}: {work}. '
            f'Trabalho concluido com foco em qualidade, prazo e acabamento. '
            f'Foto de referencia: {file_name}.'
        )
    if channel == 'LinkedIn':
        return (
            f'Case #{i}: {work}. '\
            f'Projeto com foco em prazo, precisao e acabamento. Arquivo de referencia: {file_name}.'
        )
    if channel in {'TikTok', 'Instagram'}:
        return (
            f'Projeto #{i}: {work}. '\
            f'Detalhes do antes e depois no trabalho {file_name}. '
            f"{' '.join(HASHTAGS[:4])}"
        )
    return (
        f'Projeto #{i} da semana: {work}. '\
        f'Resultado real do trabalho {file_name}. '
        f"{' '.join(HASHTAGS)}"
    )


def build_plan(manifest: dict[str, list[str]], start: date) -> str:
    days = [start + timedelta(days=i) for i in range(7)]
    assets = pick_assets(manifest, len(days), seed=int(start.strftime('%Y%m%d')))
    if not assets:
        raise SystemExit('No assets found in manifest.')

    lines = []
    lines.append(f'# Plano semanal de conteudo ({start} a {start + timedelta(days=6)})')
    lines.append('')
    lines.append('## Agenda base')
    for idx, day in enumerate(days, 1):
        cat, path = assets[(idx - 1) % len(assets)]
        lines.append(f'- {day} | Tema: {cat} | Midia: {path}')

    lines.append('')
    lines.append('## Copys por canal')
    for channel in CHANNELS:
        lines.append(f'### {channel}')
        for idx, day in enumerate(days, 1):
            cat, path = assets[(idx - 1) % len(assets)]
            lines.append(f'- {day}: {caption(channel, cat, path, idx)}')
        lines.append('')

    lines.append('## Checklist de publicacao')
    lines.append('- Confirmar 1 foto/video principal por dia')
    lines.append('- Incluir CTA para orcamento no final da legenda')
    lines.append('- Responder comentarios em ate 24h')
    lines.append('- Reaproveitar melhor post da semana em Shorts/Reels')
    return '\n'.join(lines).strip() + '\n'


def main():
    parser = argparse.ArgumentParser(description='Generate weekly social content plan')
    parser.add_argument('--start', default=str(date.today()), help='Week start date (YYYY-MM-DD)')
    args = parser.parse_args()

    start = date.fromisoformat(args.start)
    manifest = load_manifest()
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    out_file = OUT_DIR / f'weekly-social-{start}.md'
    out_file.write_text(build_plan(manifest, start), encoding='utf-8')
    print(out_file)


if __name__ == '__main__':
    main()
