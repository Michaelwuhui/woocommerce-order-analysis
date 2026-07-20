#!/bin/sh
set -eu
set -a
. /etc/woo-analysis/fulfillment.env
set +a
cd /www/wwwroot/woo-analysis
exec ./venv/bin/python tests/validate_wms_readonly.py --audit
