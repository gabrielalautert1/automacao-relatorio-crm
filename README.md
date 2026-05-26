# Automação de Relatório CRM

Script Python que automatiza a reorganização e formatação de relatórios exportados do CRM, gerando um Excel estruturado por cadência e status de leads.

# O que o script faz

- Lê um arquivo Excel exportado do CRM
- Agrupa os leads por cadência
- Separa e conta leads por status: **Ativo**, **Ganho** e **Perdido**
- Lista motivos de perda por cadência
- Exibe resumo dos leads ativos com campo personalizado
- Gera um novo Excel formatado com cores por status e data no nome do arquivo

##Tecnologias Usadas

- Python
- Pandas
- OpenPyXL

# Como usar

1. Instale as dependências:
   pip install pandas openpyxl
2. Execute o script:
   python relatorio_crm.py
3. Informe o nome do arquivo Excel quando solicitado
4. O relatório gerado será salvo como `relatorio_detalhado_DD-MM-YYYY.xlsx`

# Contexto

Projeto desenvolvido em ambiente profissional real para otimizar a análise de prospecções exportadas do CRM, eliminando retrabalho manual na geração de relatórios semanais.
