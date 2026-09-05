# Gunicorn configuration to prevent timeouts and optimize memory on Render
import multiprocessing

# Increase timeout to 120 seconds to allow for large Excel processing
timeout = 120

# Use a single worker with multiple threads to save memory (Render free tier has 512MB RAM)
workers = 1
threads = 4

# Max requests before worker restart (helps with memory leaks)
max_requests = 50
max_requests_jitter = 5

import os

# Bound address using Render PORT env variable
bind = f"0.0.0.0:{os.environ.get('PORT', '10000')}"

# Preload app for memory efficiency
preload_app = True
