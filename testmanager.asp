<%option explicit%>
<!--#include file = "class.aspdbbox.asp"-->
<h1>esempio di utilizzo di ASPdBManager</h1>
<%
private function report(value)
	if value then
		report = "<span style=""color:green"">riuscito</span><br/>"
	else
		report =  "<span style=""color:red"">fallito</span><br/>"
	end if
end function

private function reportbool(value)
	if value then
		reportbool = "<span style=""color:green"">vero</span><br/>"
	else
		reportbool =  "<span style=""color:red"">falso</span><br/>"
	end if
end function

dim dbManager
set dbManager = new Class_ASPDbManager
dbManager.debugging = false

'### CONNESSIONE MYSQL
dbManager.database = "mysql"
response.write "connessione al database mysql: " & report(dbManager.connect("test","root","",""))

'### CONNESSION ACCESS
'dbManager.database = "mdb"
'response.write "connessione al database mdb: " & report(dbManager.Connect("db/test.mdb","","",""))

response.write "<hr/>"

'### ESEGUIRE UNA QUERY DI AGGIORNAMENTO/INSERIMENTO/ELIMINAZIONE
'o.conn.execute("UPDATE tabella1 SET campoa = 5")

'### ELIMINO TABELLA
response.write "elimino ""tabella1"": " & report(dbManager.DropTable("tabella1"))

'### VERIFICARE L'ESISTENZA
response.write "esiste la tabella ""tabella1""? " & reportbool(dbManager.TableExists("tabella1"))
response.write "esiste ""campoa"" nella tabella ""tabella1""? " & reportbool(dbManager.ColumnExists("tabella1","campoa"))

'### CREA TABELLA E CAMPI (anche se esistono già restituisce TRUE)
response.write "crea ""tabella1"": " & report(dbManager.CreateTable("tabella1","id",null,true))
response.write "crea ""campoa"" in ""tabella1"" (testuale): " & report(dbManager.AddColumn("tabella1","campoa",":type=char,:size=20",null,true))
response.write "crea ""campob"" in ""tabella1"" (numerico): " & report(dbManager.AddColumn("tabella1","campob",":type=long",null,true))

'### VERIFICARE L'ESISTENZA
response.write "esiste la tabella ""tabella1""? " & reportbool(dbManager.TableExists("tabella1"))
response.write "esiste ""campoa"" nella tabella ""tabella1""? " & reportbool(dbManager.ColumnExists("tabella1","campoa"))

response.write "<hr/>"

'e così via con tutti gli altri metodi...
'vedi la documentazione per maggiori informazioni
'su http://www.imente.org
%>