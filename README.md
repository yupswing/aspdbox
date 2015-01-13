<h2>ASPdBox</h2>
<p>ASPdBox &egrave; un insieme di classi per ASP3 che si occupano di interfacciarsi fra la vostra applicazione e il database (sono supportati MDB e MYSQL) senza farvi preoccupare troppo di quale database stiate effettivamente utilizzando; inoltre vi offre una serie di utilità atte a semplificare la creazione di query sql sia per l'input che per l'output.</p>
<p>Attualmente esistono due classi:<br/>
<a href="#aspdbmanager"><b>ASPdBManager</b></a> si occupa di gestire la connessione con il database, di fornirvi tutti i metodi utili alla creazione di query compatibili, metodi per il controllo dei dati e soprattutto una serie di potenti metodi per la modifica della struttura del database.<br/>
<a href="#aspdbpagination"><b>ASPdBPagination</b></a> si occupa di rendere "ridicola" l'esecuzione di una paginazione, proprio come se steste facendo una normalissima query; oltre a ciò vi chiede quanti record per pagina e che pagina visualizzare, il resto lo fa lei e vi restituisce un recordset statico (sola lettura) pronto per essere "stampato"</p>

<hr/>
<a id="aspdbmanager"></a>
<h2>Classe ASPdBManager</h2>
<p>Tutto ciò che bisogna sapere su questa classe è che prima di fare qualsiasi cosa (dopo ovviamente averla istanziata) è impostare un database (se non si imposta obj.database verrà considerato "mdb") e successivamente aprire una connessione, poi automaticamente tutti i metodi diventano disponibili e funzionanti, senza dimenticarsi di disconnettersi alla fine del lavoro/pagina<br />
(Per vedere esempi pratici potete andare sulla <a href="demo.asp">DEMO</a>)</p>
<p>Instanza minimale:<br />
<pre>
dim obj
set obj = Class_ASPdbManager
obj.database = "mdb" 'oppure "mysql", mi raccomando MINUSCOLO
obj.debugging = true 'false è di default, se fate prove lasciatelo a TRUE, conviene
'pronti per i metodi
</pre>
</p>
<p>Vediamo ora di analizzare i metodi che mette a disposizione la classe</p>

    <p><b>Connect(db,user,password,options)</b><br/>
    Niente di complesso da spiegare. Senza questa operazione tutti gli altri metodi genereranno errori gestiti (nessuna connessione attiva) e quindi risulteranno inutilizzabili.<br/>
    Nel caso <b>mdb</b> sar&agrave; necessario indicare un percorso DB (il metodo si occupa di fare il Server.MapPath) esistente, una password (se necessario). La stringa base di connessione è la seguente <code>"Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " & server.MapPath(db) & "; Persist Security Info = False;" & options</code><br />
    Nel caso <b>mysql</b> sar&agrave; necessario indicare un nome DB, un utente e una password (se necessario). La stringa base di connessione è la seguente <code>"db=" & db & "; driver=MySQL ODBC 3.51 Driver; uid=" & user & "; pwd=" & password & ";" & options</code> (nelle options può venire utile indicare "Server=[host]" se mysql lo richiede)<br/>
    <strong>restituisce FALSE se la connessione non va a buon fine</strong></p>
    
    <p><b>Disconnect()</b><br/>
    Chiude la connessione e distrugge gli oggetti utilizzati.</p>
    
    <hr/>
    
    <h3>metodi di modifica</h3>
    
    <p><b>CreateNew(db,user,password,options,overwrite)</b><br/>
    Crea un nuovo database (da effettuarsi PRIMA di un Connect).<br />
    Nel caso <b>mdb</b> è necessario indicare un percorso DB (il metodo si occupa di fare il Server.MapPath) [user, password e options non vengono utilizzati]<br />
     - Il metodo si occupa di creare un nuovo MDB. Se esiste già restituisce FALSE, se overwrite = true elimina il file e ne crea uno nuovo vuoto e restituisce TRUE.<br />
    Nel caso <b>mysql</b> è necessario indicare un DB esistente (!), e i valori user, password e options necessari per connettervisi (come Connect).<br />
     - Il metodo si occupa di verificare che il database sia vuoto ed esista. Se non lo è restituisce FALSE, se overwrite = true distrugge tutte le tabelle presenti e restituisce TRUE. [attualmente il metodo NON può creare un nuovo database mysql]<br />
    Quindi in entrambi i casi PRUDENZA con overwrite = true.
    </p>
    
    <p><b>CreateTable(tablename,columnid,tableoptions,skipifexist)</b><br/>
    Crea una nuova tabella.<br/>
    tablename è il nome della nuova tabella, columnid è il nome del campo che vi verrà generato all'interno (sarà un campo di tipo CONTATORE).<br />
    tableoptions è una stringa che verrà aggiunta al termine della SQL (per mysql se lasciato vuoto o null il valore assume "TYPE=INNODB CHARSET=UTF8")<br />
    skipifexist: se TRUE e la tabella già esiste restituisce TRUE senza fare nulla, se FALSE e la tabella esiste la elimina e ne crea una nuova e restituisce TRUE<br />
     <strong>restituisce FALSE se la creazione della tabella non va a buon fine</strong>
    </p>
    
    <p><b>DropTable(tablename)</b><br/>
    Elimina una tabella (viene effettivamente eseguito SOLO se la tabella esiste)<br/>
    <strong>restituisce FALSE se la eliminazione della tabella non va a buon fine</strong>
    </p>
    
    <p><b>AddColumn(tablename,columnname,columnoptions,fill,skipifexist)</b><br/>
    Crea un campo in una tabella.<br />
    tablename è il nome della tabella, così come columnname è il nome del nuovo campo.<br />
    le columnoptions sono OPTIONS (<a href="#options">vedi appendice</a>).<br />
    fill è il valore che deve essere impostato sul campo per i record già esistenti (<strong>il valore non viene controllato, assicurarsi di indicarlo correttamente, si consiglia l'utilizzo delle <a href="#formatinput">funzioni di formattazione</a></strong>)<br/>
    skipifexist: se TRUE e il campo già esiste restituisce TRUE senza fare nulla, se FALSE e il campo esiste la elimina e ne crea uno nuovo e restituisce TRUE<br />
    <strong>restituisce FALSE se la creazione del campo non va a buon fine</strong>
    </p>
    
    <p><b>ModifyColumn(tablename,columnname,columnoptions,adapt)</b><br/>
    Modifica la struttura di un campo in una tabella.<br />
    tablename è il nome della tabella, così come columnname è il nome del nuovo campo.<br />
    le columnoptions sono OPTIONS (<a href="#options">vedi appendice</a>).<br />
    se adapt = TRUE prima di eseguire la modifica fa un Left (se il campo è testuale) per evitare errori di troncamento<br />
    <strong>restituisce FALSE se la modifica del campo non va a buon fine</strong>
    </p>
    
    <p><b>DropColumn(tablename,columnname)</b><br/>
    Elimina un campo (viene effettivamente eseguito SOLO se il campo esiste)<br/>
    <strong>restituisce FALSE se la eliminazione del campo non va a buon fine</strong>
    </p>
    
    <p><b>AddIndex(tablename,columnname,indexoptions,skipifexist)</b><br/>
    Crea un indice su un campo<br />
    tablename è il nome della tabella, così come columnname è il nome del campo.<br />
    le indexoptions sono OPTIONS (<a href="#options">vedi appendice</a>).<br />
    skipifexist: se TRUE e l'indice già esiste restituisce TRUE senza fare nulla, se FALSE e l'indice esiste la elimina e ne crea uno nuovo e restituisce TRUE<br />
    <strong>restituisce FALSE se la creazione dell'indice non va a buon fine</strong>
    </p>
    
    <p><b>DropIndex(tablename,indexname)</b><br/>
    Elimina un indice (viene effettivamente eseguito SOLO se l'indice esiste)<br/>
    <strong>restituisce FALSE se la eliminazione dell'indice non va a buon fine</strong>
    </p>
    
    <p><b>AddForeignKey(tablename,columnname,byval foreignkeyname,references,foreignkeyoptions,skipifexist)</b><br/>
    Crea una relazione fra due (o più) campi<br />
    tablename è il nome della tabella, così come columnname è il nome del campo.<br />
    la foreignkeyname è il nome della relazione (se non impostato la chiama "FK_tablename_columnname")<br />
    references è il riferimento (es di valore accettato: "nometabella(nomecampo)")<br />
    le foreignkeyoptions sono OPTIONS (<a href="#options">vedi appendice</a>).<br />
    skipifexist: se TRUE e la foreign key già esiste restituisce TRUE senza fare nulla, se FALSE e la foreign key esiste la elimina e ne crea una nuova e restituisce TRUE<br />
    <strong>restituisce FALSE se la creazione della foreign key non va a buon fine</strong>
    </p>
    
    <p><b>DropForeignKey(tablename,foreignkeyname)</b><br/>
    Elimina una relazione (viene effettivamente eseguito SOLO se la relazione esiste)<br/>
    <strong>restituisce FALSE se la eliminazione della relazione non va a buon fine</strong>
    </p>
    
    <hr/>
    
    <h3>metodi di interrogazione</h3>
    
    
    <p><b>TableExists(tablename)</b><br/>
    Resituisce TRUE/FALSE a seconda se la tabella indicata esista o meno
    </p>
    
    <p><b> TablesOnDatabase()</b><br/>
    Restituisce un array con i nomi di tutte le tabelle nel database<br />
    (se non vi sono tabelle restituisce un array con un elemento vuoto)<br />
    (in MDB non inserisce le tabelle di sistema "MSys*"
    </p>
    
    <p><b>ColumnExists(tablename,columnname)</b><br/>
    Resituisce TRUE/FALSE a seconda se la colonna indicata esista o meno
    </p>
    
    <p><b>ColumnsOnTable(tablename)</b><br/>
    Restituisce un array con i nomi di tutti i campi nella tabella<br />
    (se non vi sono tabelle restituisce un array con un elemento vuoto)
    </p>
    
    <p><b>IndexExists(tablename,indexname)</b><br/>
    Resituisce TRUE/FALSE a seconda se l'indice indicato esista o meno
    </p>
    
    <p><b>IndexesOnTable(tablename)</b><br/>
    Restituisce un array con i nomi di tutti gli indici nella tabella<br />
    (se non vi sono tabelle restituisce un array con un elemento vuoto)
    </p>
    
    <p><b>ForeignKeyExists(tablename,foreignkeyname)</b><br/>
    Resituisce TRUE/FALSE a seconda se la relazione indicata esista o meno
    </p>
    
    <p><b>ForeignKeysOnTable(tablename)</b><br/>
    Restituisce un array con i nomi di tutte le relazioni nella tabella<br />
    (se non vi sono tabelle restituisce un array con un elemento vuoto)
    </p>
    
    <hr/>
    
    <h3>metodi di compatibilità/input/output</h3>
    
    <p><b>compatibleSQL(SQL)</b><br/>
    Data una SQL ne restituisce una versione compatibile con il database in uso.<br />
    L'SQL se già corretta non verrà modificata<br />
    (attualmente applica modifiche SOLO alle istruzioni DELETE e LIMIT/TOP<br />
    verificare la compatibilità delle proprie query se esulano da questi due casi per il/i database utilizzati<br />
    se ci sono consigli da dare segnalate!)
    </p>
    <a id="formatinput"></a>
    <p><b>fDate(date)</b><br/>
    Data una data restituisce la stringa da inserire in SQL per il database in uso<br />
    es:</p>
        <pre>SQL = "UPDATE miatabella SET campodata = " & obj.fDate(now) & ";"
        obj.conn.execute(SQL)</pre>
    <p>nel caso indicato aggiorna tutta la tabella impostando la data odierna.<br />
    notare che non sono stati inserite apici (') o altri simboli.<br />
    la funzione provvede a restituirli se necessario</p>
    
    <p><b>fString(value)</b><br/>
    Restituisce la stringa formattata per evitare errori/sql injection<br />
    es:</p>
        <pre>SQL = "UPDATE miatabella SET campotesto = '" & obj.fString("questo e' un test") & "';"
        obj.conn.execute(SQL)</pre>
    <p>nel caso indicato aggiorna tutta la tabella impostando "questo e' un test".<br />
    notare che debbono essere indicati apici (') prima e dopo.<br />
    inoltre grazie a fString la presenza di un apici nel testo non presenterà errori</p>
    
    <p><b>fNumber(value)</b><br/>
    Restituisce la stringa formattata per evitare errori<br />
    es:</p>
        <pre>SQL = "UPDATE miatabella SET camponumerico = " & obj.fNumber("25,2") & ";"
        obj.conn.execute(SQL)</pre>
        <pre>SQL = "UPDATE miatabella SET camponumerico = " & obj.fNumber(clng(12)) & ";"
        obj.conn.execute(SQL)</pre>
    <p>nel caso indicato aggiorna tutta la tabella impostando 25.2 o 12<br />
    notare che la funzione accetta sia valori effettivamente numerici che stringhe<br />
    si occuper&agrave; lei di modificarli secondo necessità.<br />
    inoltre, indicare stringhe non numeriche non presenterà errori, ma una restituzione di 0</p>
    
    
    

<hr/>
    <a id="options"></a>
    <h3>Le opzioni, che sono?</h3>
    <p>Per semplificarci la vita, invece che mettere decine di argomenti molto spesso inutilizzati ho pensato di fare un'unica stringa in cui vengono scritti tutti i valori necessari.<br />
    La sintassi è: <code>":nome=valore,:altronome=altrovalore"</code><br/>
    Le funzioni che utilizzano le OPZIONI sono AddColumn, ModifyColumn, AddIndex e AddForeignKey</p>
    <p><strong>AddColumn,ModifyColumn</strong><br />
    le opzioni valide (a seconda dei casi) sono</p>
    <ul>
        <li><strong>:type</strong> tipologia di campo(<a href="#types">vedi sotto per la tabella dei valori</a>)</li>
        <li><strong>:size</strong> numero di caratteri (utilizzato per type = char o varchar)</li>
        <li><strong>:signed</strong> true/false [il numero è positivo o negativo] (utilizzato SOLO in mysql per i campi numerici)</li>
        <li><strong>:null</strong> true/false [il campo può essere null] (tutti i campi)</li>
        <li><strong>:default</strong> stringa/numero valore automatico (tutti i campi, se non dato e NOT NULL per i numerici è 0)</li>
        <li><strong>:binary</strong> true/false [campo binario] (utilizzato per type = char, varchar e memo)</li>
    </ul></p>
    <p><strong>AddIndex</strong><br />
    le opzioni valide (a seconda dei casi) sono</p>
    <ul>
        <li><strong>:name</strong> nome dell'indice (se non dato utilizza il nome del campo)</li>
        <li><strong>:index</strong><ul><li><em>UNIQUE</em> (mysql) (mdb)</li><li><em>FULLTEXT</em> (mysql [myisam {char,varchar,text}])</li><li><em>SPATIAL</em> (mysql [myisam {not null}])</li></ul></li>
        <li><strong>:type</strong><ul><li><em>BTREE</em> (mysql [myisam,innodb,memory])</li><li><em>HASH</em> mysql [memory,ndb])</li><li><em>RTREE</em> (mysql [myisam]) [solo per SPATIAL index]</li></ul></li>
    </ul></p>
    <p><strong>AddForeignKey</strong><br />
    le opzioni valide (a seconda dei casi) sono</p>
    <ul>
        <li><strong>:onupdate, :ondelete</strong> (solo mysql)
            <ul>
                <li><em>CASCADE</em> delete or update the row from the parent table and automatically delete or update the matching rows in the child table</li>
            <li><em>SET NULL</em>   delete or update the row from the parent table and set the foreign key column or columns in the child table to NULL.</li>
            <li><em>NO ACTION</em>      reject delete or update actions on parent table</li>
            <li><em>RESTRICT</em>   same as NO ACTION</li>
            </ul>
</li>
    </ul></p>
<hr/>
    <a id="types"></a>
    <h3>Tabella dei tipi compatibili</h3>
    <p>Rappresenta la conversione dal valore indicato in <strong>:type</strong> a quello effettivamente utilizzato nelle query di creazione/modifica dei campi<br />
    Se vi sono errori o imperfezioni ogni consiglio è bene accetto, ho cercato di renderli più sinonimi possibile</p>
    <table>
        <tr style="background-color:#FFCC00;font-weight:bold;">
            <td>:type</td>
            <td>mdb</td>
            <td>mysql</td>
        </tr>
        <tr>
            <td><strong>counter</strong></td>
            <td><small>COUNTER UNIQUE NOT NULL PRIMARY KEY</small></td>
            <td><small>INT(10) AUTO_INCREMENT UNIQUE NOT NULL PRIMARY KEY</small></td>
        </tr>
        <tr>
            <td><strong>boolean</strong></td>
            <td>BIT</td>
            <td>BOOLEAN</td>
        </tr>
        <tr>
            <td><strong>date</strong></td>
            <td>DATETIME</td>
            <td>DATETIME</td>
        </tr>
        <tr>
            <td><strong>short</strong></td>
            <td>SHORT</td>
            <td>SMALLINT</td>
        </tr>
        <tr>
            <td><strong>long</strong></td>
            <td>LONG</td>
            <td>INT</td>
        </tr>
        <tr>
            <td><strong>single</strong></td>
            <td>SINGLE</td>
            <td>FLOAT</td>
        </tr>
        <tr>
            <td><strong>double</strong></td>
            <td>DOUBLE</td>
            <td>DOUBLE</td>
        </tr>
        <tr>
            <td><strong>currency</strong></td>
            <td>CURRENCY</td>
            <td>DECIMAL(19,4)</td>
        </tr>
        <tr>
            <td><strong>char</strong></td>
            <td>TEXT</td>
            <td>CHAR</td>
        </tr>
        <tr>
            <td><strong>varchar</strong></td>
            <td>TEXT</td>
            <td>VARCHAR</td>
        </tr>
        <tr>
            <td><strong>memo</strong></td>
            <td>LONGTEXT</td>
            <td>LONGTEXT</td>
        </tr>
        <tr>
            <td><strong>blob</strong></td>
            <td>LONGBINARY</td>
            <td>LONGBLOB</td>
        </tr>
    </table>
    
<hr/>
<a id="aspdbpagination"></a>
<h2>Classe ASPdBPagination</h2>
<p>E ora veniamo a questa piccola perla (si lo so sono megalomane): ASPdBPagination<br />
E' una piccola classe che si occupa di fare la paginazione e risparmiarvi la noiosa implementazione "volta per volta".
L'utilizzo è di una semplicità che definirei goduriosa (si sono molto soddisfatto)<br /></p>
<p>Prima di tutto istanziamo la classe</p>
<pre>
&lt;!--#include file = "includes/class.aspdbbox.asp"--&gt;
&lt;%
dim obj, database
'Conn è un oggetto Connection valido e in stato OPEN
database = "mysql" 'oppure "mdb"
set obj = new Class_ASPDbPagination
onj.RecordsPerPageDefault = 25 'si può impostare il valore base di record per pagina
obj.debugging = true 'come al solito il debuggin non è necessario, ma utile durante i test
%&gt;</pre>
<p>E poi?</p>
<pre>
&lt;%
dim looper
if obj.Paginate(Conn, database, 5, null, _
        "SELECT id, campoa FROM tabella ORDER BY id") then
    
    
    looper = 0
    while not obj.recordset.eof and (looper < obj.RecordsPerPage or database = "mysql")
        response.write "&lt;strong>id&lt/strong&gt;:" &amp; obj.recordset("id") &amp; "&lt;br/&gt;" & _
                   "&lt;strong>campoa&lt/strong&gt;:" &amp; obj.recordset("campoa") &amp; "&lt;br/&gt;"
        obj.recordset.movenext
        looper = looper + 1
    wend
    obj.recordset.close
    
    '* opzionale: stampa il box di navigazione (per maggiori personalizzazioni è possibile crearlo autonomamente)
    response.write obj.printNavigator(null,null,null,false)
else
    'errore nella query o nella classe (sob)
    'attivare il debugging per verificare quale sia
end if
%&gt;</pre>
<p><strong>In questo caso la query è semplice, ma non ci sono quasi limiti nella sua creazione, vediamo quali:</strong><br/>
- la query deve essere funzionante [potete testarla dando come parametro RECORDSPERPAGE = 0]<br />
- per l'utilizzo su entrambi i database mantenere la sintassi comune<br />
- vietato l'utilizzo di LIMIT x,y (la classe serve proprio a questo)
<br /><br />

<strong>Per il resto sono supportati l'utilizzo di:</strong><br/>
- tutta la sintassi base SQL (compatibile con MDB e/o MYSQL) [JOIN,GROUP BY,HAVING,SELECT DISTINCT,WHERE...etc.etc.]<br />
- UNION e UNION ALL<br />
- TOP x e LIMIT x (sia in query UNION che per singole select si può utilizzare TOP x e LIMIT x)<br />
&nbsp;&nbsp;nella stessa query sono permessi anche mescolamenti (una SELECT con TOP e una con LIMIT)<br />
&nbsp;&nbsp;l'applicazione si occupa di renderli compatibili per il database in uso<br />
- se utilizzato First(column) o Last(column) la query viene resa compatibile con MySQL, ma ovviamente se ne perdono le funzionalità
</p>

<hr/>

<p>Vediamo i pochi metodi nel dettaglio:</p>
    
    <p><b>Paginate(Conn, database, recordsPerPage, Page, pagsql)</b><br/>
    <em>Conn</em> è una connessione attiva (se utilizzate ASPdBManager potete usare quella)<br />
    <em>database</em> è il tipo di database (valori validi "mdb","mysql")<br />
    <em>recordPerPage</em> è il numero di record per pagina (se impostata a NULL viene valorizzata tramite request.QueryString("perpage"); se da querystring viene passato un valore 0 o negativo viene utilizzato <em>RecordsPerPageDefault</em>)<br />
    &nbsp;&nbsp;<span style="color:red;">impostando il valore <em>recordPerPage</em> a 0 la paginazione non verr&agrave; effettuata e potrete verificare la bontà della query)</span><br />
    <em>page</em> è la pagina corrente (se impostata a NULL viene valorizzata tramite request.QueryString("page"))<br />
    <em>pagsql</em> è la query da paginare (vedi limitazioni indicate sopra)<br /><br />
    
    i valori per page e recordperpage vengono validati prima di essere utilizzati<br />
    se la pagina è superiore al numero delle pagine viene impostata a 1<br /><br />
    
    Paginate restituisce TRUE se l'esecuzione è andata a buon fine, altrimenti FALSE (e impostando il debugging avrete a schermo l'errore)<br />
    Nel caso in cui vada tutto bene avrete disponibili le seguenti variabili:</p>
    <ul>
        <li><strong>Page</strong> pagina corrente</li>
        <li><strong>Record</strong> record corrente</li>
        <li><strong>Pages</strong> numero di pagine</li>
        <li><strong>Records</strong> numero di record totali</li>
        <li><strong>RecordsPerPage</strong> numero di record per pagina</li>
        <li><strong>Recordset</strong> il recordset statico con la pagina richiesta</li>
    </ul>
    
    <p><b>printNavigator(text,add,queryadd,printperpage)</b><br/>
    Se non volete scrivere un navigatore fra le pagine potete utilizzare questa funzione che ne genera uno per voi<br />
    Lasciate text a NULL<br />
    in add potete indicare HTML da immettere sempre all'interno della form che verrà generata<br />
    queryadd è una stringa che verr&agrave; aggiunta a tutti i link (es: queryadd = "&amp;amp;search=qualcosa")<br />
    printperpage è una flag (true/false) che indica se il valore "perpage" deve essere messo nei link (utile solo se lasciate a NULL l'argomento recordsPerPage in PAGINATE)</p>
    
    <p><b>Reset()</b><br/>
    Reimposta tutti i valori per una nuova paginazione (non è necessario poiché Paginate lo esegue)</p>

