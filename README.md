# Ticket Processor

Processador de NFS-e e DPS a partir de PDFs de faturas Ticket e Sem Parar.

## 🤘 Tecnologias

- **Flask**: Framework web
- **pdfplumber**: Extração de texto de PDFs
- **openpyxl**: Geração de planilhas Excel
- **Gunicorn**: Servidor WSGI para produção

## 🤖 Estrutura

```
ticket-processor/
├── web_app.py           # Aplicação Flask principal
├── app.py               # Versão Streamlit (local)
├── requirements.txt     # Dependências Python
├── render.yaml          # Configuração do Render
├── templates/           # Templates HTML
│   └── index.html
└── static/              # Arquivos estáticos (CSS)
    └── css/
        └── style.css
```
