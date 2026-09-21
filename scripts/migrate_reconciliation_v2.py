"""Explicit additive migration; requires the operator to name the intended database."""
import argparse
import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from inv_common import get_conn
from reconciliation_schema import migrate

if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--expected-database',required=True)
    args=parser.parse_args()
    conn=get_conn()
    if not hasattr(conn,'_raw'):
        raise SystemExit('This deployment entry point requires PostgreSQL.')
    actual=conn.execute('SELECT current_database()').fetchone()[0]
    if actual!=args.expected_database:
        conn.close();raise SystemExit('Database target differs from --expected-database; no migration performed.')
    migrate(conn)
    print('Additive reconciliation schema installed in '+actual)
    conn.close()
