"""
Cartón Médico - Aplicación Flask para autocompletar cartones médicos desde PDFs.
"""
from flask import Flask


def create_app():
    """Factory de la aplicación Flask."""
    application = Flask(__name__)

    from app.routes import bp
    application.register_blueprint(bp)

    return application
