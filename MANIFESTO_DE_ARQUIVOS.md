# Manifesto de Arquivos e Integridade

**Data do inventário:** 26 de julho de 2026  
**Objetivo:** comprovar que todo material editorial existente nesta etapa foi preservado e recebeu destino explícito.

## Arquivos materiais

| Destino no repositório | Origem | Tamanho | SHA-256 | Função |
|---|---|---:|---|---|
| `00-FONTES-BRUTAS/Texto_colado.txt` | `upload/Texto colado.txt` | 26.017 B | `3611562e9e1ef5e0b9842fa9dae6355e777f0c49411cffaa7c5c8982a563e972` | Texto-fonte bruto preservado. |
| `01-MANUSCRITO/FUGA_IDENTITARIA_MANUSCRITO_MESTRE.md` | arquivo canônico local | 36.268 B | `43385485a13c4351e124bf8e1d3989216164a6feb68b21be96035e7d599dda69` | Manuscrito mestre. |
| `02-PESQUISA-E-CONTROLE/FUGA_IDENTITARIA_CENTRAL_DE_PESQUISA.md` | arquivo canônico local | 109.986 B | `3daa8b4e066f35ec189fcad425f6caa0dd182d6b1fe476d3a87d2c11d9543344` | Governança, pesquisa, lotes e decisões. |
| `02-PESQUISA-E-CONTROLE/FUGA_IDENTITARIA_TRIAGEM_EDITORIAL_V1.md` | arquivo local | 35.681 B | `8c0451e92a7c2f01ee9cb8cb026a61a06df68d6a27936df145048a85678c920a` | Auditoria do texto-fonte e do sistema editorial. |
| `03-ARQUITETURA/FUGA_IDENTITARIA_MAPA_MESTRE_PRODUCAO_2_0.md` | arquivo canônico local | 33.252 B | `77342ebfa744da8022e6fc72c31bacdf1bd32081530e910ec6405f6b730e3472` | Estrutura e produção modular. |
| `04-ENTREGAS-PARA-REVISAO/Fuga_Identitaria_Mapa_Mestre_Producao_2_0.docx` | entrega local | 33.202 B | `2dfc47374ae3133f20c6874a914d669fae8bb89cc07c94aadff0f924da741451` | Versão editável para revisão. |
| `04-ENTREGAS-PARA-REVISAO/Fuga_Identitaria_Mapa_Mestre_Producao_2_0.pdf` | entrega local | 480.834 B | `c373544cb2c3f6ab07756ea247e58647e5756a64d03852c8afa79d2f03fb665a` | Cópia estável de revisão. |
| `05-FERRAMENTAS/build_editorial_map.py` | script local | 9.710 B | `6e4449188e01df6398ab2f1cebd44fbe5d6aa57e89d34e542638fbec858b56ae` | Reprodução do DOCX editorial. |
| `99-ARQUIVO-SENSIVEL-PRIVADO/B88F1503-F1E2-4A69-90EA-FB48F287C9D2.jpeg` | anexo original | 290.280 B | `1b4929c892603a7e0a805d2e73aead8e4d81af0d1f8e218c812e20e1888a21d9` | Fotografia familiar preservada, uso proibido. |
| `99-ARQUIVO-SENSIVEL-PRIVADO/37F4BF3B-BAC8-4ABB-82F4-4B1855CBC6C3.jpeg` | anexo original | 26.797 B | `158bd814223234cd7cdd77e894d377a880a47a6b4025d1c66fdb3135a156df91` | Fotografia familiar preservada, uso proibido. |

## Arquivos de governança criados para o repositório

- `README.md`
- `GOVERNANCA_E_PRIVACIDADE.md`
- `DIREITOS_AUTORAIS.md`
- `STATUS_EDITORIAL.md`
- `CHANGELOG.md`
- `MANIFESTO_DE_ARQUIVOS.md`
- `.gitignore`
- `99-ARQUIVO-SENSIVEL-PRIVADO/README.md`

## Exclusões intencionais

Não foram promovidos ao repositório:

- páginas PNG geradas apenas para inspeção visual;
- folhas de contato;
- relatórios temporários de acessibilidade;
- caches;
- arquivos de sistema;
- sessões de upload abandonadas.

Esses itens são reproduzíveis ou transitórios e não contêm conteúdo autoral adicional.

## Regra de conferência

Se houver dúvida sobre perda ou alteração:

1. conferir este manifesto;
2. recalcular o SHA-256;
3. comparar o resultado com a tabela;
4. não substituir o arquivo silenciosamente;
5. registrar a nova versão no `CHANGELOG.md`.
