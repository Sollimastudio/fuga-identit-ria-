#!/usr/bin/env bash
set -euo pipefail

base_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
parts_dir="$base_dir/binarios-em-partes"
output_dir="$base_dir/output-reconstruido"

mkdir -p "$output_dir"

cat "$parts_dir"/Fuga_Identitaria_Camada_11_Kindle_Reflow.docx.part-* > "$output_dir/Fuga_Identitaria_Camada_11_Kindle_Reflow.docx"
cat "$parts_dir"/FUGA_IDENTITARIA_CAMADA_11_FECHAMENTO_KINDLE.epub.part-* > "$output_dir/FUGA_IDENTITARIA_CAMADA_11_FECHAMENTO_KINDLE.epub"
cat "$parts_dir"/Fuga_Identitaria_Camada_11_Prova_Tecnica.pdf.part-* > "$output_dir/Fuga_Identitaria_Camada_11_Prova_Tecnica.pdf"

cd "$output_dir"

printf '%s  %s\n' \
  '8ce8bc560d5e9de71f84d87cfe45d2a0661ae3883308931b049d8aa545259ee5' \
  'Fuga_Identitaria_Camada_11_Kindle_Reflow.docx' \
  'ecdf7d703505cdd07a6fbfb0b8da2dd7f89709e848c385fa9133523151361963' \
  'FUGA_IDENTITARIA_CAMADA_11_FECHAMENTO_KINDLE.epub' \
  'd01e2e3ccf9950660397854bae98740a1b6905847c110a685ef6018c35a64893' \
  'Fuga_Identitaria_Camada_11_Prova_Tecnica.pdf' | sha256sum -c -

printf 'Camada 11 reconstruída e validada em: %s\n' "$output_dir"
