#!/usr/bin/env bash
# Generate a localhost TLS certificate the bridge (and the task pane it serves)
# can use. Prefers mkcert when available because it is trusted by Windows.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
CERT_DIR="$ROOT/certs"
CERT="$CERT_DIR/localhost.crt"
KEY="$CERT_DIR/localhost.key"

mkdir -p "$CERT_DIR"

if [[ -f "$CERT" && -f "$KEY" ]]; then
  echo "certificate already present at $CERT"
  exit 0
fi

if command -v mkcert >/dev/null 2>&1; then
  echo "generating trusted certificate with mkcert"
  mkcert -install >/dev/null 2>&1 || true
  mkcert -cert-file "$CERT" -key-file "$KEY" localhost 127.0.0.1 ::1
else
  echo "mkcert not found; generating a self-signed certificate with openssl"
  cat > "$CERT_DIR/openssl.cnf" <<'EOF'
[req]
distinguished_name = dn
x509_extensions = v3
prompt = no
[dn]
CN = localhost
[v3]
subjectAltName = @alt
keyUsage = digitalSignature, keyEncipherment
extendedKeyUsage = serverAuth
[alt]
DNS.1 = localhost
IP.1 = 127.0.0.1
IP.2 = ::1
EOF
  openssl req -x509 -nodes -newkey rsa:2048 -days 825 \
    -keyout "$KEY" -out "$CERT" -config "$CERT_DIR/openssl.cnf"
fi

echo "certificate: $CERT"
echo "key:         $KEY"
echo
echo "On Windows, import $CERT into 'Trusted Root Certification Authorities'"
echo "so Excel trusts the task pane origin (see docs/install.md)."
