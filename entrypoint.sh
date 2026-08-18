#!/bin/bash
set -e

if [ -n "$EXO_CERT_BASE64" ]; then
    echo "$EXO_CERT_BASE64" | base64 -d > /app/cert.pfx
    export EXO_CERT_PATH="/app/cert.pfx"
    echo "Certificate written to /app/cert.pfx"
else
    echo "WARNING: EXO_CERT_BASE64 not set. Exchange auth will fail."
fi

exec pwsh -NoLogo -NonInteractive -File /app/listener.ps1

