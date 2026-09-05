# Gunicorn configuration to prevent timeouts and optimize memory on Render
import multiprocessing

# Increase timeout to 300 seconds to allow for large Excel processing
timeout = 300

# Use a single worker with multiple threads to save memory (Render free tier has 512MB RAM)
workers = 1
threads = 8

# High max_requests to prevent worker restarts from killing background processing threads
# Previously 50 - that was killing background threads mid-reconciliation
max_requests = 1000
max_requests_jitter = 50

import os

# Bound address using Render PORT env variable
bind = f"0.0.0.0:{os.environ.get('PORT', '10000')}"

# Preload app for memory efficiency
preload_app = True

# Keep-alive connections
keepalive = 65
