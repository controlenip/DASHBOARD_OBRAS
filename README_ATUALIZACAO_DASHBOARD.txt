NIP | Dashboard Demanda - Atualizacao de interface

Substituir nip_dashboard_template.html no repositorio. Nao e necessario trocar os arquivos Python se ja estao na ultima versao.

Alteracoes: QUALIDADE DA BASE removida da navegacao e conteudo do FAQ; busca avancada mantida como busca unica visivel; legenda com Prioritarias vermelho #dc2626.

Nota tecnica: elementos legados de qualidade e busca simples permanecem inacessiveis na interface por compatibilidade com scripts antigos, evitando quebrar os modulos e contagens existentes.

Rodar: python test_remocao_qualidade_busca.py
