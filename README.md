# Ticket Processor

Processador de NFS-e e DPS a partir de PDFs de faturas Ticket.

## 🚀 Deploy no Render

### Pré-requisitos
- Conta no [Render](https://render.com)
- Repositório GitHub com o código

### Passo a Passo

1. **Faça commit e push do código para o GitHub:**
   ```bash
   git add .
   git commit -m "Preparar para deploy no Render"
   git push origin main
   ```

2. **Acesse o Render:**
   - Vá para [render.com](https://render.com)
   - Faça login com sua conta GitHub

3. **Crie um novo Web Service:**
   - Clique em **"New +"** → **"Web Service"**
   - Selecione o repositório `ticket-processor`
   - Clique em **"Connect"**

4. **Configure o serviço:**
   - **Name**: `ticket-processor` (ou o nome que preferir)
   - **Environment**: `Python 3`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn web_app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`
   - **Instance Type**: `Free` (ou escolha um plano pago para melhor performance)

5. **Variáveis de Ambiente (opcional):**
   - `FLASK_DEBUG`: `false`
   - `MAX_CONTENT_LENGTH`: `52428800`

6. **Deploy:**
   - Clique em **"Create Web Service"**
   - Aguarde o deploy (pode levar alguns minutos)
   - Seu app estará disponível em `https://ticket-processor.onrender.com`

## 📝 Uso Local

### Instalação
```bash
# Criar ambiente virtual
python -m venv .venv

# Ativar ambiente virtual (Windows)
.venv\Scripts\Activate.ps1

# Instalar dependências
pip install -r requirements.txt
```

### Executar
```bash
# Modo desenvolvimento
python web_app.py

# Com Gunicorn (produção local)
gunicorn web_app:app --workers 2 --timeout 120
```

Acesse: `http://localhost:5000`

## 🔧 Tecnologias

- **Flask**: Framework web
- **pdfplumber**: Extração de texto de PDFs
- **openpyxl**: Geração de planilhas Excel
- **Gunicorn**: Servidor WSGI para produção

## 📁 Estrutura

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

## 🐛 Troubleshooting

### Erro 503 Service Unavailable
- Verifique os logs no Render Dashboard
- Confirme que o `startCommand` está correto
- Aumente o timeout se necessário

### Uploads não funcionam
- O Render usa sistema de arquivos efêmero
- Arquivos temporários são deletados após o deploy
- Para armazenamento persistente, use serviços como AWS S3

### Application timeout
- Aumente o valor de `--timeout` no `startCommand`
- Considere usar um plano pago com mais recursos

## 📌 Notas

- A versão free do Render pode hibernar após 15 minutos de inatividade
- O primeiro acesso após hibernação pode demorar ~30 segundos
- Para produção séria, considere planos pagos do Render
