"""
Punto de entrada de la aplicación Cartón Médico.
"""
from app import create_app
from config import HOST, PORT

application = create_app()

if __name__ == "__main__":
    print(f">> Servidor Cartón médico en http://127.0.0.1:{PORT}")
    application.run(host=HOST, port=PORT, debug=False)
