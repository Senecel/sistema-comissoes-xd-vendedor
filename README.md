# Sistema de Comissões XD - Vendedor

**Autor:** António Vanga II  
**Versão:** 1.0.0  

---

## Descrição
Este projeto nasceu a partir de uma necessidade de uma entidade que utiliza o XD e pretendia automatizar o cálculo das comissões dos seus vendedores. O objetivo foi criar uma ferramenta simples, rápida e eficiente para consolidar dados de vendas e calcular comissões automaticamente, evitando erros manuais e poupando tempo.

Desenvolvido em **Python** com **Tkinter** e **ttkbootstrap**, o sistema permite importar dados de vendas, processar e gerar resumos detalhados com valores e comissões, oferecendo uma interface moderna e responsiva.

---

## Funcionalidades Principais
- **Importação de ficheiros:** Suporta `.xlsx`, `.xls`, `.csv` e `.ods`.  
- **Processamento inteligente:** Limpeza de dados, conversão de valores e filtragem vetorizada usando **pandas**.  
- **Cálculo de comissões:** Regras escalonadas com base no total de vendas:
  - Até 300.000 → 5%
  - 300.001 a 700.000 → 10%
  - 700.001 a 1.000.000 → 12%
  - Acima de 1.000.000 → 15%
- **Resumo por vendedor:** Agrupa e soma valores, exibindo total e comissão.  
- **Visualização interativa:** Filtragem por vendedor ou exibição completa dos dados.  
- **Eliminação de linhas:** Remoção de registos selecionados diretamente na interface.  
- **Exportação de resumo:** Gera ficheiros Excel formatados com totais, percentagens e comissões.  
- **Interface moderna:** Botões, labels, comboboxes e indicadores de progresso com **ttkbootstrap**.  
- **Feedback visual:** Loading animado durante o processamento de dados.

---

## Tecnologias Utilizadas
- Python 3.x  
- pandas  
- Tkinter (via ttkbootstrap)  
- xlsxwriter / openpyxl / xlrd / odf para leitura/escrita de ficheiros

---

## Estrutura do Código
- `calcular_percentagem(total)` → Calcula a comissão com base no total de vendas.  
- `ler_arquivo(caminho)` → Leitura otimizada de ficheiros em diferentes formatos.  
- `processar(caminho)` → Limpeza de dados e conversão para valores numéricos.  
- `gerar_resumo(df)` → Criação de resumo com totais e comissões.  
- **Interface:** Organizada em frames para header, ações, tabela e rodapé.  
- **Threads:** Mantêm a interface responsiva durante o processamento.  
- **Filtragem eficiente:** Atualização em blocos para datasets grandes, evitando travamentos.

---

## Motivação
O sistema foi desenvolvido para automatizar a geração de comissões numa entidade que utiliza XD, substituindo processos manuais, garantindo precisão e oferecendo uma solução fácil de usar para gestores e vendedores.

---

## Como Utilizar
1. Abrir a aplicação:  
```bash
python calcular_comissao.py
