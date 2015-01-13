<%option explicit%>
<!--#include file = "class.aspdbbox.asp"-->
<h1>esempio di utilizzo di ASPdBPagination</h1>
<%

dim dbManager
set dbManager = new Class_ASPDbManager
dbManager.debugging = true

'######################################################
'######################################################
'### IMPOSTA LE VARIABILI PRIMA DI CONTINUARE

dim db,user,password
dbManager.database = "mdb"	'oppure MYSQL
db = "db.mdb"				'nome del database 'es: mysql "nomedb", es: mdb "db/nomedb.mdb"
user = ""					'solo per MYSQL
password = ""				'se necessaria

'######################################################
'######################################################


if request.QueryString("setup") = "ok" then

%><h1>setup db</h1>
<p>se si presentano errori verifica le variabili inserite nel codice<br />
se &egrave; un database MDB verifica di avere permessi di scrittura nella cartella indicata<br />
se &egrave; un database MYSQL verifica di avere i permessi e che il database esista già (NOTA CHE IL SUO CONTENUTO VERRA' COMPLETAMENTE CANCELLATO)</p><%


	'### ELIMINO TABELLA
	call dbManager.CreateNew(db,user,password,"",true)	
	call dbManager.Connect(db,user,password,"")	
	'### CREA TABELLA E CAMPI
	call dbManager.CreateTable("tabella1","id",null,true)
	call dbManager.AddColumn("tabella1","campoa",":type=char,:size=20",null,true)
	call dbManager.AddColumn("tabella1","campob",":type=long",null,true)
	'### CI METTE UN PO' DI DATI
	randomize
	dim ii
	for ii = 0 to 49
		call dbManager.conn.execute("INSERT INTO tabella1 (campoa,campob) VALUES ('testo" & int(rnd*1000) & "'," & int(rnd*1000) & ")")
	next
	
	response.write "<strong>ho fatto tutto. vai alla <a href=""testpagination.asp"">paginazione</a></strong>"
	response.end
	

else

%><a href="?setup=ok">crea e prepara database</a> (prima di preparare il db imposta le variabile nel codice (linea 14-18)<%

	call dbManager.connect(db,user,password,"")

	dim obj, database, looper
	'Conn è un oggetto Connection valido e in stato OPEN
	set obj = new Class_ASPDbPagination
	obj.debugging = true 'come al solito il debuggin non è necessario, ma utile durante i test
	if obj.Paginate(dbManager.Conn, dbManager.database, null, null, _
						 "SELECT id, campoa, campob FROM tabella1 ORDER BY id") then
		
		
		looper = 0
		if obj.recordset.eof then
			response.write "recordset vuoto"
		end if
		while not obj.recordset.eof and (looper < obj.RecordsPerPage or database = dbManager.database)
			response.write "<strong>id</strong>:" & obj.recordset("id") & " | " & _
						   "<strong>campoa</strong>:" & obj.recordset("campoa") & " | " & _
						   "<strong>campob</strong>:" & obj.recordset("campob") & "<br/>"
			obj.recordset.movenext
			looper = looper + 1
		wend
		obj.recordset.close
		
		'* opzionale: stampa il box di navigazione (per maggiori personalizzazioni è possibile crearlo autonomamente)
		response.write obj.printNavigator(null,null,null,false)
	else
		'errore nella query o nella classe (sob)
		'attivare il debugging per verificare quale sia
		response.write "errore nella query o niente connessione"
	end if

end if

response.write "<hr/>"

'vedi la documentazione per maggiori informazioni
'su http://www.imente.org
%>