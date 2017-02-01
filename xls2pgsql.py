#!/usr/bin/python
###############################
# Script to export xls or xlsx to PostgreSQL
# Angel Joyce Torres Ramirez
# joys.tower@gmail.com
# require xlrd psycopg2
# free to use
# usage:
#       xls2pgsql.py -x <xlsfile> -e <xls_encode> -t <table to export> -h <host> -p <port> -u <user> -w <password> -d <database>
#
# the new table created all columns created as varchar
###############################
import os, sys, getopt
import unicodedata


PG_CONN_STRING = ""
PG_CONN=None
XLS_FILE=None
XLS_ENCODE=None
TABLE=None

#Comand Usage
def printCmdUsage():
   print 'xls2pgsql.py -x <xlsfile> -e <xls_encode> -t <table to export> -h <host> -p <port> -u <user> -w <password> -d <database>'
   sys.exit(2)


#Conection definitions##########
def conect2PG():
   
   global PG_CONN_STRING
   global PG_CONN
   try:
      PG_CONN=psycopg2.connect(PG_CONN_STRING)
   except:
      print "Can't connect to the database"
      sys.exit(2)

def closeConn():
   global PG_CONN
   if PG_CONN is not None:
      PG_CONN.close()
################################



#Read XLS##########

def remove_accents(input_str):
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return u"".join([c for c in nfkd_form if not unicodedata.combining(c)])

def xls2pg():
   from xlrd.sheet import ctype_text
   global XLS_FILE
   global XLS_ENCODE
   global PG_CONN
   bk = None
   print "Reading XLS"
   if( XLS_ENCODE == None):
      bk = xlrd.open_workbook(XLS_FILE)
   else:
      bk = xlrd.open_workbook(XLS_FILE, encoding_override=XLS_ENCODE)
   print "Creating "+TABLE+" table"
   sheet = bk.sheet_by_index(0)
   row = sheet.row(0)
   create_table = "CREATE TABLE "+TABLE+"("
   for idx, cell_obj in enumerate(row):
      #cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
      #print cell_type_str
      col = remove_accents(cell_obj.value)
      col = col.replace(" ","_").lower()
      create_table = create_table + col + " varchar,"
   create_table = create_table[0:-1] + ");"
   
   num_cols = sheet.ncols
   print "Inserting data"
   try:
      cur = PG_CONN.cursor()
      cur.execute( create_table )
      for row_idx in range(1, sheet.nrows):
         insert = "INSERT INTO "+TABLE+" VALUES("
         row = sheet.row(row_idx)
         for idx, cell_obj in enumerate(row):
            cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
            if cell_type_str == "text":
               cell = unicode(cell_obj.value)
               cell = cell.encode('UTF-8')
               cell = cell.replace("'","''")
            else:
               cell = str(cell_obj.value)
            insert = insert + "'"+ cell +"',"
         insert = insert[0:-1] + ");"
         cur.execute( insert )
         if( (row_idx % 100) == 0 ):
            print ("."),
      print "Commiting data"
      cur.close()
      PG_CONN.commit()
   except (Exception, psycopg2.DatabaseError) as error:
        print(error)
   finally:
      closeConn()
       
   
   

################################   
   
   
   
#Main
def readOptions():
   try:
      opts, args = getopt.getopt(sys.argv[1:],"x:e::t:h:p::u:w::d:")
   except getopt.GetoptError:
      printCmdUsage("1")
      
   global XLS_ENCODE
   global XLS_FILE
   global TABLE
   
   conn_options = ('-h', '-d', '-p', '-u', '-w')
   connString = ""
   for opt, arg in opts:
      if opt == '-x':
         XLS_FILE = arg
      elif opt == '-e':
         XLS_ENCODE = arg
      elif opt == '-t':
         TABLE = arg
      elif opt == '-h':
         connString = connString + " host='" + arg +"'"
      elif opt == '-p':
         connString = connString + " port='" + arg +"'"
      elif opt == '-u':
         connString = connString + " user='" + arg +"'"
      elif opt == '-w':
         connString = connString + " password='" + arg +"'"
      elif opt == '-d':
         connString = connString + " dbname='" + arg +"'"
   if(XLS_FILE == ""):
      printCmdUsage()
   if(TABLE == ""):
      printCmdUsage()
   if(connString == ""):
      printCmdUsage()

   global PG_CONN_STRING
   PG_CONN_STRING = connString


if __name__ == "__main__":
   readOptions()
   import psycopg2
   conect2PG()
   import xlrd
   xls2pg()
   closeConn()
   
