import os
import sys
from urllib.parse import parse_qs

# Ensure project root is in sys.path
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

from app import app


class VercelPathMiddleware:
    """
    Vercel rewrites forward requests to /api/index with the original path
    passed in the __path__ query parameter.
    This WSGI middleware restores the original PATH_INFO so Flask routes work transparently.
    """
    def __init__(self, wsgi_app):
        self.wsgi_app = wsgi_app

    def __call__(self, environ, start_response):
        qs = parse_qs(environ.get("QUERY_STRING", ""))
        if "__path__" in qs and qs["__path__"]:
            target_path = qs["__path__"][0]
            if not target_path.startswith("/"):
                target_path = "/" + target_path
            environ["PATH_INFO"] = target_path
        elif environ.get("PATH_INFO") == "/api/index":
            environ["PATH_INFO"] = "/"
        return self.wsgi_app(environ, start_response)


app.wsgi_app = VercelPathMiddleware(app.wsgi_app)
