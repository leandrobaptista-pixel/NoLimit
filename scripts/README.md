# Scripts de operacao

## 1) Validar links e referencias locais

```bash
python3 scripts/check_links.py
```

## 1.5) Sincronizar novas fotos da galeria automaticamente

Se voce colocar novas fotos diretamente dentro das pastas reais da galeria, este comando:

- escaneia cada categoria
- atualiza `assets/gallery-manifest.json`
- atualiza `assets/gallery-manifest.js`
- regenera as imagens otimizadas da homepage
- valida links e referencias locais
- mostra quais arquivos novos entraram e qual sera o link hospedado

```bash
python3 scripts/gallery_sync_pipeline.py
```

Publicar automaticamente depois da sincronizacao:

```bash
python3 scripts/gallery_sync_pipeline.py --publish
```

Opcoes uteis:

```bash
# so ver o que mudou, sem escrever arquivos
python3 scripts/gallery_sync_pipeline.py --dry-run

# sincronizar sem regenerar thumbs/features
python3 scripts/gallery_sync_pipeline.py --skip-optimize

# sincronizar sem rodar verificacao final
python3 scripts/gallery_sync_pipeline.py --skip-check

# sincronizar e tambem rodar a checagem ampla do site inteiro
python3 scripts/gallery_sync_pipeline.py --full-site-check
```

## 2) Gerar plano semanal de conteudo (Facebook, YouTube e outros canais)

```bash
python3 scripts/generate_weekly_content.py --start 2026-02-24
```

- Troque a data para o inicio da semana desejada (`YYYY-MM-DD`).
- Saida: `content/weekly/weekly-social-<data>.md`

## 3) Atualizar manifest da galeria (quando adicionar fotos)

```bash
python3 - <<'PY'
import json
from pathlib import Path

root = Path('.')
category_dirs = {
  'Trim': 'assets/00-Trim',
  'Wainscoting': 'assets/00-Wainscoting',
  'Stairs': 'assets/00-Stairs',
  'Ceiling': 'assets/00-Ceiling',
  'Decks': 'assets/00-Decks',
  'Kitchen & Vanities': 'assets/00-kitchen & Vanities',
  'Fireplaces & Bars': 'assets/00-Fireplaces & Bars',
  'Outside Doors & Windows': 'assets/00-Outside Doors & Windows',
  'Pergola': 'assets/00-Pergola',
  'Port & Portal': 'assets/00-Port & Portal',
}
allowed = {'.jpg', '.jpeg', '.png', '.webp', '.gif', '.JPG', '.JPEG', '.PNG', '.WEBP', '.GIF'}

manifest = {}
for cat, rel in category_dirs.items():
    d = root / rel
    items = []
    if d.exists():
        for p in sorted(d.iterdir()):
            if p.is_file() and p.suffix in allowed:
                name = p.name
                if 'copy' in name.lower() or 'originalfullresolutionimage' in name.lower():
                    continue
                items.append(f"{rel}/{name}")
    manifest[cat] = items

(root / 'assets/gallery-manifest.json').write_text(
    json.dumps(manifest, indent=2, ensure_ascii=True) + '\n',
    encoding='utf-8'
)
print('assets/gallery-manifest.json atualizado')
PY
```
