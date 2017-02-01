# xls2pgsql.py

## How to use

Script to export xls or xlsx to postgres database
```
$ xls2pgsql.py -x <xlsfile> -e <xls_encode> -t <table to export> -h <host> -p <port> -u <user> -w <password> -d <database>

$ xls2pgsql.py -x Proyecciones_hogares_indigenas.xls -t conapo.proy_hog_indigena_10_20 -h localhost -u postgres -d db_proys
```
For security propouses use .pgpass to store your postgres passwords to don't use the -w option


## Requiered libs
 * [xlrd](https://pypi.python.org/pypi/xlrd)
 * [psycopg2](https://pypi.python.org/pypi/psycopg2)
