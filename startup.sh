#!/bin/bash
# Write secrets from App Service Application Settings
mkdir -p /home/site/wwwroot/.streamlit
cat > /home/site/wwwroot/.streamlit/secrets.toml <<EOF
[graph]
app_id         = "${GRAPH_APP_ID}"
tenant_id      = "${GRAPH_TENANT_ID}"
client_secret  = "${GRAPH_CLIENT_SECRET}"
shared_mailbox = "${GRAPH_SHARED_MAILBOX}"
EOF

# Install dependencies
pip install -r /home/site/wwwroot/requirements.txt --quiet

# Start Streamlit
exec python -m streamlit run /home/site/wwwroot/app.py \
    --server.port 8000 \
    --server.address 0.0.0.0 \
    --server.headless true
