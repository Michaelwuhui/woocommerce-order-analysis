#!/bin/sh
set -eu

if [ "$#" -ne 1 ]; then
  echo "usage: $0 /absolute/backup/directory" >&2
  exit 2
fi

backup_dir=$1
case "$backup_dir" in
  /www/backup/woo-analysis-pre-migration-*) ;;
  *) echo "refusing unexpected backup path: $backup_dir" >&2; exit 2 ;;
esac

mkdir -p "$backup_dir"
chmod 700 "$backup_dir"
cd /www/wwwroot/woo-analysis
sqlite3 woocommerce_orders.db ".backup '$backup_dir/woocommerce_orders.db'"
sqlite3 "$backup_dir/woocommerce_orders.db" \
  "PRAGMA integrity_check; SELECT COUNT(*) FROM orders;" > "$backup_dir/verification.txt"
sha256sum "$backup_dir/woocommerce_orders.db" > "$backup_dir/SHA256SUMS"
ls -lh "$backup_dir"
