# RELATÓRIO DE QA — CAMADA 11

## Artefatos verificados

- `FUGA_IDENTITARIA_CAMADA_11_FECHAMENTO_KINDLE.md`
- `Fuga_Identitaria_Camada_11_Kindle_Reflow.docx`
- `FUGA_IDENTITARIA_CAMADA_11_FECHAMENTO_KINDLE.epub`
- `Fuga_Identitaria_Camada_11_Prova_Tecnica.pdf`
- `CAPA_FUGA_IDENTITARIA_KINDLE_1600x2560.jpg`

## Estrutura editorial

| Verificação | Resultado |
|---|---:|
| Partes | 6 |
| Capítulos | 24 |
| Figuras | 21 |
| Dicionário autoral | 35 verbetes |
| Entradas principais no sumário EPUB | 38 |
| Imagens em linha no DOCX | 21 |
| Textos alternativos no EPUB | 21/21 |
| Maior texto alternativo | 134 caracteres |

O auditor de títulos sinaliza listas numeradas como possíveis títulos manuais. A inspeção confirma falsos positivos: o Mapa da travessia, perguntas e listas internas são conteúdo, não entradas omitidas do sumário.

## DOCX refluível

- pacote ZIP íntegro;
- uma seção em orientação retrato;
- estilos semânticos de título preservados;
- 21 imagens em linha e nenhuma imagem flutuante;
- nenhum cabeçalho;
- nenhum rodapé;
- nenhum campo `PAGE` ou `NUMPAGES`;
- nenhuma margem espelhada;
- links internos do mapa e do dicionário presentes;
- metadados de título, autora e assunto presentes;
- auditoria automática de acessibilidade: **0 achados altos, 0 médios e 0 baixos**.

## EPUB

- padrão: EPUB 3.3;
- layout: refluível;
- idioma: `pt-BR`;
- pacote ZIP íntegro;
- EPUBCheck: **0 fatais, 0 erros, 0 avisos e 0 informações**;
- 68 itens no manifesto;
- 43 itens na ordem de leitura;
- 38 destinos no sumário lógico;
- links do Mapa da travessia resolvem para as seis partes;
- link da abertura para o Dicionário autoral preservado;
- 21 imagens significativas, todas com texto alternativo;
- imagem de capa tratada como apresentação no XHTML de capa;
- 11 links externos e 47 links internos identificados.

## PDF de prova

- páginas: **357**;
- tamanho: **396,85 × 595,304 pt**, equivalente a 14 × 21 cm;
- PDF marcado: sim;
- metadados presentes: sim;
- fontes incorporadas e com mapeamento Unicode: sim;
- criptografia: não;
- JavaScript: não;
- primeira página, direitos, abertura, epílogo, mapa, segurança, biografia e referências inspecionados;
- todas as 357 páginas renderizadas;
- todos os 357 PNGs de inspeção carregados integralmente;
- páginas 121–150 e 331–357 reinspecionadas em folhas de contato;
- nenhuma sobreposição, corte de figura ou transbordamento detectado;
- órfão do epílogo eliminado;
- páginas de respiro restantes correspondem a términos de seção, não a páginas vazias acidentais.

## Capa

- dimensões: **1600 × 2560 px**;
- proporção: 1:1,6;
- modo: RGB;
- metadado de resolução: 300 dpi;
- formato: JPEG;
- tamanho: 197.549 bytes;
- título, subtítulo e autoria legíveis em miniatura e em tamanho integral;
- nenhuma informação encosta na borda de corte digital.

## Conteúdo e amostra

- palavras antes das referências: 65.694;
- palavras totais: 67.971;
- referência editorial de 10%: 6.569 palavras;
- início do Capítulo 1: aproximadamente 7,07%;
- início do Capítulo 2: aproximadamente 10,01%;
- a amostra estimada contém promessa, orientação, mapa, limites conceituais, Caso do Leitor e o Capítulo 1 completo.

O limite real do “Leia uma amostra” será calculado pela Amazon depois do processamento e não pode ser garantido por paginação local.

## Resultado

**QA local aprovado.**

Pendência externa obrigatória: importar o DOCX aprovado no Kindle Create, revisar o projeto em diferentes aparelhos, fontes, tamanhos e fundos, exportar o KPF e repetir a prova no Previewer on-line do KDP.
