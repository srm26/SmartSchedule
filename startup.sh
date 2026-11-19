#!/bin/bash
# Install dependencies
pip install -r /home/site/wwwroot/requirements.txt --quiet

# Auto-sync: SharePoint -> Outlook every 15 minutes, in its own process, so it runs right after a deploy or
# restart without anyone opening the app. Restarted if it ever stops.
cd /home/site/wwwroot
(while true; do python /home/site/wwwroot/sync_engine.py; sleep 60; done) &

# Start Streamlit
exec python -m streamlit run /home/site/wwwroot/app.py \
    --server.port 8000 \
    --server.address 0.0.0.0 \
    --server.headless true
