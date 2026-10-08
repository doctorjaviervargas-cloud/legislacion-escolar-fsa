import os
bind = f"0.0.0.0:{os.getenv('PORT','8000')}"
workers = 1
threads = int(os.getenv('GUNICORN_THREADS','4'))
timeout = 120
preload_app = True
accesslog = '-'
errorlog = '-'
