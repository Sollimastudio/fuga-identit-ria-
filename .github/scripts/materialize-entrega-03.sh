#!/usr/bin/env bash
set -euo pipefail

sudo apt-get update
sudo apt-get install -y libreoffice-writer poppler-utils fonts-urw-base35 fonts-noto-core python3-pip libcairo2
python -m pip install --upgrade pip
python -m pip install python-docx cairosvg pillow

rm -rf /tmp/entrega03-source
mkdir -p /tmp/entrega03-source
cat .entrega03-source.part-* | base64 -d > /tmp/entrega03-source.tar.gz
tar -xzf /tmp/entrega03-source.tar.gz -C /tmp/entrega03-source

DELIVERY='04-ENTREGAS-PARA-REVISAO/ENTREGA-03'
mkdir -p "$DELIVERY"
cp -a /tmp/entrega03-source/ENTREGA-03/. "$DELIVERY/"
cp /tmp/entrega03-source/FERRAMENTAS/build_entrega_03.py "$DELIVERY/build_entrega_03.py"

(
  cd "$DELIVERY"
  python build_entrega_03.py
  test -s Fuga_Identitaria_Entrega_03_Parte_III.docx
  test -s Fuga_Identitaria_Entrega_03_Parte_III.pdf
  test "$(pdfinfo Fuga_Identitaria_Entrega_03_Parte_III.pdf | awk '/^Pages:/ {print $2}')" -ge 58
  pdftotext Fuga_Identitaria_Entrega_03_Parte_III.pdf /tmp/entrega03-text.txt
  grep -q 'Capítulo 12' /tmp/entrega03-text.txt
  grep -q 'Notas de pesquisa da Entrega 3' /tmp/entrega03-text.txt
)

python - <<'PY'
from pathlib import Path
import hashlib, re, shutil
root = Path('.')
delivery = root / '04-ENTREGAS-PARA-REVISAO/ENTREGA-03'
registry = delivery / 'REGISTRO_EDITORIAL_ENTREGA_03.md'
paths = {
 'FUGA_IDENTITARIA_ENTREGA_03_TEXTO.md': delivery/'FUGA_IDENTITARIA_ENTREGA_03_TEXTO.md',
 'Fuga_Identitaria_Entrega_03_Parte_III.docx': delivery/'Fuga_Identitaria_Entrega_03_Parte_III.docx',
 'Fuga_Identitaria_Entrega_03_Parte_III.pdf': delivery/'Fuga_Identitaria_Entrega_03_Parte_III.pdf',
 'figuras/FIG-07_DE_ONDE_CHEGAM_OS_MAPAS_DO_EU.svg': delivery/'figuras/FIG-07_DE_ONDE_CHEGAM_OS_MAPAS_DO_EU.svg',
 'figuras/FIG-07_DE_ONDE_CHEGAM_OS_MAPAS_DO_EU.png': delivery/'figuras/FIG-07_DE_ONDE_CHEGAM_OS_MAPAS_DO_EU.png',
 'figuras/FIG-08_O_CICLO_DA_RECOMENDACAO.svg': delivery/'figuras/FIG-08_O_CICLO_DA_RECOMENDACAO.svg',
 'figuras/FIG-08_O_CICLO_DA_RECOMENDACAO.png': delivery/'figuras/FIG-08_O_CICLO_DA_RECOMENDACAO.png',
 'figuras/FIG-09_EU_VIVIDO_E_EU_EXIBIDO.svg': delivery/'figuras/FIG-09_EU_VIVIDO_E_EU_EXIBIDO.svg',
 'figuras/FIG-09_EU_VIVIDO_E_EU_EXIBIDO.png': delivery/'figuras/FIG-09_EU_VIVIDO_E_EU_EXIBIDO.png',
 'figuras/FIG-10_QUANDO_A_CAUSA_COMECA_A_USAR_A_PESSOA.svg': delivery/'figuras/FIG-10_QUANDO_A_CAUSA_COMECA_A_USAR_A_PESSOA.svg',
 'figuras/FIG-10_QUANDO_A_CAUSA_COMECA_A_USAR_A_PESSOA.png': delivery/'figuras/FIG-10_QUANDO_A_CAUSA_COMECA_A_USAR_A_PESSOA.png',
 'build_entrega_03.py': delivery/'build_entrega_03.py',
}
text = registry.read_text(encoding='utf-8')
for name, path in paths.items():
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    pat = re.compile(rf'(\| `{re.escape(name)}` \| `)[0-9a-f]{{64}}(` \|)')
    text, n = pat.subn(rf'\g<1>{digest}\g<2>', text)
    if n != 1:
        raise SystemExit(f'Hash ausente: {name}')
registry.write_text(text, encoding='utf-8')

for dest in (root/'manuscrito/entregas/entrega-03', root/'repo_package/04-ENTREGAS-PARA-REVISAO/ENTREGA-03'):
    if dest.exists(): shutil.rmtree(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(delivery, dest)

research = root/'02-PESQUISA-E-CONTROLE/REGISTRO_PESQUISA_ENTREGA_03.md'
research.parent.mkdir(parents=True, exist_ok=True)
shutil.copy2('/tmp/entrega03-source/PESQUISA/REGISTRO_PESQUISA_ENTREGA_03.md', research)
for dest in (root/'05-FERRAMENTAS/build_entrega_03.py', root/'repo_package/05-FERRAMENTAS/build_entrega_03.py'):
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(delivery/'build_entrega_03.py', dest)
PY

rm -f .entrega03-source.part-* .entrega03-materialize
rm -f .github/workflows/materialize-entrega-03.yml
rm -f .github/scripts/materialize-entrega-03.sh

git config user.name 'github-actions[bot]'
git config user.email '41898282+github-actions[bot]@users.noreply.github.com'
git add -A -f
git commit -m 'build: materializa Entrega 03 completa'
git push origin HEAD:agent/entrega-03
