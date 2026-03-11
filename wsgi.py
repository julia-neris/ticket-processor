"""
WSGI entry point para produção
"""
import logging
import sys
import os

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

logger.info("=" * 50)
logger.info("Iniciando aplicação...")
logger.info(f"Python version: {sys.version}")
logger.info(f"Working directory: {os.getcwd()}")
logger.info(f"Python path: {sys.path}")
logger.info("=" * 50)

try:
    from web_app import app
    logger.info("✓ web_app.app importado com sucesso!")
except Exception as e:
    logger.error(f"✗ Erro ao importar web_app.app: {e}")
    raise

# Exportar a aplicação para o Gunicorn
application = app

if __name__ == "__main__":
    app.run()
