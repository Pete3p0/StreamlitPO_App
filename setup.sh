#!/usr/bin/env bash
mkdir -p ~/.streamlit
cat > ~/.streamlit/config.toml << 'EOF'
[server]
headless = true
enableCORS = false
enableXsrfProtection = false
port = ${PORT}
EOF
