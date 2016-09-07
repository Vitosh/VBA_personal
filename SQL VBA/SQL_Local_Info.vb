Servertyp:
Datenbankmodul

Servername:
(localdb)\MSSQLLocalDB

Authentifizierung:
Windows-Authentifizierung

str_connection_string = "Provider=SQLNCLI11;Server=(localdb)\MSSQLLocalDB;Initial Catalog=Tempt;Trusted_Connection=yes;timeout=30;"

str_connection_string = "Provider=" & arr_info(0) & _
  "; Data Source=" & arr_info(1) & _
  "; Database=" & str_generator(arr_info(2), True) & _
  ";User ID=" & str_generator(arr_info(3), True) & _
  "; Password=" & str_generator(arr_info(4), True) & ";"
