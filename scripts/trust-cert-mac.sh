#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
CERT_DIR="$ROOT/certs"
PEM="$CERT_DIR/localhost.pem"
KEY="$CERT_DIR/localhost-key.pem"
CFG="$CERT_DIR/openssl.cnf"

mkdir -p "$CERT_DIR"

echo "==> 生成带 SAN 的 localhost 证书…"
openssl req -x509 -newkey rsa:2048 \
  -keyout "$KEY" \
  -out "$PEM" \
  -days 825 \
  -nodes \
  -config "$CFG" \
  -extensions v3_req

echo "==> 证书信息："
openssl x509 -in "$PEM" -noout -subject -dates -ext subjectAltName

echo
echo "==> 接下来需要管理员密码，把证书设为「始终信任」(SSL)…"
echo "    （会先清理旧的同名证书，再写入系统钥匙串）"
echo

# 清理旧条目（忽略失败）
sudo security delete-certificate -c localhost /Library/Keychains/System.keychain 2>/dev/null || true

# 以信任根方式安装，并显式信任 SSL
sudo security add-trusted-cert -d -r trustRoot \
  -k /Library/Keychains/System.keychain \
  "$PEM"

echo
echo "==> 当前系统信任列表中的 localhost："
sudo security dump-trust-settings -d 2>/dev/null | awk '/localhost/{p=1} p&&/Cert /{if(!/localhost/)exit} p{print}'

echo
echo "完成。请："
echo "  1) 保持 npm run dev 运行"
echo "  2) 完全退出 Word (⌘Q) 再打开"
echo "  3) 再点「格式助手」"
