# PDF Ledger Extractor

Este projeto transforma relatórios financeiros em PDF (texto selecionável) em bases estruturadas e prontas para análise.

O foco é resolver problemas reais encontrados em PDFs financeiros:

- Linhas quebradas que fragmentam um mesmo lançamento
- Mistura de campos (Parceiro invadindo o Histórico)
- Rodapés (“Saldo Final”, “Totais”) misturados aos registros
- Palavras coladas decorrentes da extração de PDF
- Trechos duplicados
- Formatação inconsistente de valores monetários

A solução combina:

- Extração de texto via pdfplumber
- Regex para identificação de datas e valores
- Parsing heurístico para separar Parceiro e Histórico
- Normalização e limpeza de texto
- Validação automática por totais e saldo final
- Opção de anonimização para compartilhamento seguro

## Como rodar

```bash
python src/pdf_ledger_extractor.py


## Observações

Funciona melhor em PDFs com texto selecionável. PDFs escaneados exigem OCR (não incluso).
Por segurança, este repositório não inclui PDFs ou extratos reais.
Use o modo de anonimização do script para gerar exemplos sem dados sensíveis.
