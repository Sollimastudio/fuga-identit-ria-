# CAMADA 11 — FECHAMENTO KINDLE NO GITHUB

## Estado correto

A Camada 11 foi executada integralmente e auditada localmente. Esta pasta registra o fechamento técnico da obra completa para revisão de Sol.

Ela não substitui silenciosamente o manuscrito mestre histórico e não está marcada como aprovada, canônica ou publicada.

## Conteúdo preservado

- manuscrito integral em Markdown;
- DOCX refluível para Kindle Create;
- EPUB 3.3 validado;
- PDF de prova técnica;
- capa Kindle 1600 × 2560;
- 21 diagramas editoriais em PNG;
- metadados KDP;
- checklist Kindle Create;
- decisões pendentes de Sol;
- relatórios editorial e de QA;
- script e recursos de geração;
- hashes SHA-256.

## Arquivos binários divididos

DOCX, EPUB e PDF foram preservados em partes numeradas dentro de `binarios-em-partes/` para permitir a sincronização segura pelo conector do GitHub. Execute:

```bash
bash RECONSTRUIR_BINARIOS_CAMADA_11.sh
```

O script recompõe os três arquivos em `output-reconstruido/` e confere os hashes dos resultados.

## Ilustrações

As 21 figuras funcionais estão presentes em `figuras/`, com 2000 × 1200 px, legendas no manuscrito e textos alternativos no EPUB.

Pendência visual declarada: os PNGs foram auditados e estão funcionais, mas a última lapidação mobile-first e os mestres SVG editáveis das figuras que ainda não os possuem devem ser concluídos antes de considerar o fechamento visual definitivo.

## Próximo marco

1. Sol revisa manuscrito, capa, biografia, metadados e decisões pendentes.
2. A camada visual recebe a última lapidação para telefone e E Ink.
3. O DOCX aprovado é importado no Kindle Create.
4. O KPF é inspecionado no Previewer.
5. Somente depois disso a edição pode ser liberada para publicação.
