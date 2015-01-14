<%
'-------------------------------------------------------------------------------------------
' ASPDbBox
' ^^^^^^^^
'   v0.11
    ' The MIT License (MIT)

    ' Copyright (c) 2008-2015 Simone Cingano

    ' Permission is hereby granted, free of charge, to any person obtaining a copy
    ' of this software and associated documentation files (the "Software"), to deal
    ' in the Software without restriction, including without limitation the rights
    ' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    ' copies of the Software, and to permit persons to whom the Software is
    ' furnished to do so, subject to the following conditions:

    ' The above copyright notice and this permission notice shall be included in all
    ' copies or substantial portions of the Software.

    ' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    ' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    ' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    ' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    ' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    ' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    ' SOFTWARE.
'	  
'   https://github.com/yupswing/aspdbox
'-------------------------------------------------------------------------------------------

Class Class_ASPdBManager

	public database 'support "mdb","mysql"
	private connected
	private currentdb
	public Conn, SQL, Rs
	private Reg
	public debugging,errorstop
	
	private sub Class_initialize()
	
		Set Reg = New RegExp 	
		Reg.Global = True
		Reg.Ignorecase = True
		
		debugging = false
		errorstop = true
		connected = false
		
		database = "mdb"
	end sub
	
	private sub Class_Terminate()
		call Disconnect()
	end sub
	
	'/***********************************************************************************/
	
	public function Connect(db,user,password,options)
		call Disconnect()
		Set conn = Server.CreateObject("ADODB.Connection")
		dim openstring, return
		
		return = true
		select case database
			case "mdb"
				openstring = "Provider = Microsoft.Jet.OLEDB.4.0; " & _
							 "Data Source = " & server.MapPath(db) & "; " & _
							 "Persist Security Info = False;"
				if password <> "" then _
				      openstring = openstring & "Jet OLEDB:Database Password=" & password & ";"
			case "mysql"
				openstring = "db=" & db & "; " & _
							 "driver=MySQL ODBC 3.51 Driver; " & _
							 "uid=" & user & "; " & _
							 "pwd=" & password & ";"
			case else
				return = false
		end select
		
		if options <> "" then openstring = openstring & options
		
		on error resume next
			if return then Conn.Open openstring
			if err.number <> 0 then
				if debugging then call Report("CN",openstring,errorstop)
				return = false
			end if
		on error goto 0
		
		currentdb = db
		
		Connect = return
		connected = return
	end function
	
	public sub Disconnect()
		if (connected) then Conn.Close
		connected = false
		set rs = nothing
		set conn = nothing
	end sub
	
	'/***********************************************************************************/
	
	'create a brand new MDB ready to be filled
	private function createMDB(argPath)
		createMDB  = true
		on error resume next
			Dim engine
			Set engine = CreateObject("DAO.DBEngine.36")
			engine.CreateDatabase server.mappath(argPath), ";LANGID=0x0409;CP=1252;COUNTRY=0", 64
			set engine = Nothing
			if err.number <> 0 then
				if debugging then call Report("NA",server.mappath(argPath),true)
				createMDB  = false
			end if
		on error goto 0
	end function
	
	'delete a file
	private function fileKill(argPath)
		fileKill = false
		on error resume next
			dim fs
			Set fs=Server.CreateObject("Scripting.FileSystemObject") 
			if fs.FileExists(server.mappath(argPath)) then
				fs.DeleteFile(server.mappath(argPath))
			end if
			set fs = nothing
			if err.number = 0 then 
				fileKill  = true
			elseif debugging then
				call Report("FK",server.mappath(argPath),false)
			end if
		on error goto 0
	end function
	
	'check if a file exists
	private function fileExists(argPath)
		fileExists = false
		dim fs
		Set fs = Server.CreateObject("Scripting.FileSystemObject")
		if fs.FileExists(Server.MapPath(argPath)) Then fileExists = true
		set fs = nothing
	end function
	
	'/***********************************************************************************/
	
	public function CreateNew(db,user,password,options,overwrite)
		CreateNew = false
		select case database
			case "mdb"
				if fileExists(db) then
					if not overwrite then exit function 'overwrite not permitted
					if not fileKill(db) then exit function 'deleting not permitted
				end if
				if not createMDB(db) then exit function 'creation failed (permissions?)
			case "mysql"
				if not Connect(db,user,password,options) then exit function 'unable to connect
				Set rs = Conn.OpenSchema(20)
				if not rs.eof then
					if not overwrite then exit function 'overwrite not permitted
					do while not rs.eof
						SQL = "DROP TABLE " & rs("TABLE_NAME")
						Conn.execute SQL
						rs.MoveNext
					loop
				end if
				rs.Close
				Disconnect()
		end select 
		CreateNew = true
	end function
	
	'/***********************************************************************************/
	
	public function CreateTable(tablename,columnid,byval tableoptions,skipifexist)
		if not connected then 
			if debugging then call Report("NC","",errorstop)
			CreateTable = false : exit function
		end if
		CreateTable = true
		
		if (TableExists(tablename)) then
			if skipifexist then exit function
			if not DropTable(tablename) then 
				CreateTable = false
				exit function
			end if
		end if
		
		if trim(tableoptions&"") = "" and database = "mysql" then
			tableoptions = "TYPE=INNODB CHARSET=UTF8" 'also you can use, for example, CHARSET=LATIN1 and/or TYPE=MYISAM
		end if
		
		on error resume next
		SQL = "CREATE TABLE " & tablename & " (" & columnid & " " & columnDef(":type=counter") & ") " & tableoptions
		Conn.execute(SQL)
		if err.number <> 0 then
			if debugging then call Report("TC",SQL,errorstop)
			CreateTable = false
		end if
		on error goto 0
	
	end function
	
	public function DropTable(tablename)
		if not connected then 
			if debugging then call Report("NC","",errorstop)
			DropTable = false : exit function
		end if
		DropTable = true
		
		if (TableExists(tablename)) then
			on error resume next
			SQL = "DROP TABLE " & tablename
			Conn.execute(SQL)
			if err.number <> 0 then 
				if debugging then call Report("TD",SQL,errorstop)
				DropTable = false
			end if
			on error goto 0	
		end if
		
	end function
	
	public function AddColumn(tablename,columnname,columnoptions,fill,skipifexist)
		if not connected then 
			if debugging then call Report("NC","",errorstop)
			AddColumn = false : exit function
		end if
		AddColumn = true
		
		if (ColumnExists(tablename,columnname)) then
			if skipifexist then exit function
			if not DropColumn(tablename,columnname) then 
				AddColumn = false
				exit function
			end if
		end if
		
		on error resume next
		SQL = "ALTER TABLE " & tablename & " ADD " & columnname & " " & columnDef(columnoptions)
		Conn.execute(SQL)
		if err.number <> 0 then
			if debugging then call Report("CA",SQL,errorstop)
			AddColumn = false
		end if
		if not isnull(fill) then
			SQL = "UPDATE " & tablename & " SET " & columnname & " = " & fill
			Conn.execute(SQL)
		end if
		on error goto 0
	
	end function
	
	public function ModifyColumn(tablename,columnname,columnoptions,adapt)
		if not connected then 
			if debugging then call Report("NC","",errorstop)
			ModifyColumn = false : exit function
		end if
		ModifyColumn = true
		
		if (ColumnExists(tablename,columnname)) then
			
			if adapt then
				on error resume next
				if inValue(parseOption(columnoptions,"type"),"|char|varchar|") then
					SQL = "UPDATE " & tablename & " SET " & columnname & " = Left(" & columnname & "," & parseOption(columnoptions,"size") & ")"
					conn.execute(SQL)
				end if
				on error goto 0
			end if
		
			on error resume next
			SQL = "ALTER TABLE " & tablename
			if database = "mdb" then
			SQL = SQL & " ALTER COLUMN"
			elseif database = "mysql" then
			SQL = SQL & " MODIFY COLUMN"
			end if
			SQL = SQL & " " & columnname & " " & columnDef(columnoptions)
			Conn.execute(SQL)
			if err.number <> 0 then
				if debugging then call Report("CM",SQL,errorstop)
				ModifyColumn = false
			end if
			on error goto 0	
		else
			ModifyColumn = AddColumn(tablename,columnname,columnoptions,null,false)
		end if
	
	end function
	
	public function DropColumn(tablename,columnname)
		if not connected then 
			if debugging then call Report("NC","",errorstop)
			DropColumn = false : exit function
		end if
		DropColumn = true
		
		if (ColumnExists(tablename,columnname)) then
			on error resume next
			SQL = "ALTER TABLE " & tablename & " DROP COLUMN " & columnname
			Conn.execute(SQL)
			if err.number <> 0 then
				if debugging then call Report("CD",SQL,errorstop)
				DropColumn = false
			end if
			on error goto 0	
		end if
	
	end function
	
	
	':name= indexname
	':type=
	' 		BTREE (mysql [myisam,innodb,memory]
	' 		HASH (mysql [memory,ndb])
	' 		RTREE (mysql [myisam]) [only for SPATIAL index]
	':index= 
	'		UNIQUE (mysql) (mdb)
	'		FULLTEXT (mysql [myisam {char,varchar,text}])
	'		SPATIAL (mysql [myisam {not null}])
	public function AddIndex(tablename,columnname,indexoptions,skipifexist)
		if not connected then 
			if debugging then call Report("NC","",errorstop)
			AddIndex = false : exit function
		end if
		AddIndex = true
		
		dim ioptions
		set ioptions = parseOptions(indexoptions)
		if (trim(ioptions("name")&"")="") then ioptions("name") = columnname
		
		if (IndexExists(tablename,ioptions("name"))) then
			if skipifexist then exit function
			if not DropIndex(tablename,ioptions("name")) then 
				AddIndex = false
				exit function
			end if
		end if
		
		on error resume next
		
		SQL = "CREATE "
		
		if (trim(ioptions("index")&"")<>"") then
		
			if database = "mdb" and ioptions("index") = "unique" then
				SQL = SQL & "UNIQUE "
			elseif database = "mysql" then
				SQL = SQL & ucase(ioptions("index")) & " "
			end if
		
		end if
		
		SQL = SQL & "INDEX " & ioptions("name")
		
		if (trim(ioptions("type")&"")<>"" and database = "mysql") then
			SQL = SQL & " USING " & ucase(ioptions("type"))
		end if
		
		SQL = SQL & " ON " & tablename & " (" & columnname & ")"
		
		Conn.execute(SQL)
		if err.number <> 0 then
			if debugging then call Report("IA",SQL,errorstop)
			AddIndex = false
		end if
		on error goto 0
		
	end function
	
	public function DropIndex(tablename,indexname)
		if not connected then 
			if debugging then call Report("NC","",errorstop)
			DropIndex = false : exit function
		end if
		DropIndex = true
		
		if (IndexExists(tablename,indexname)) then
			on error resume next
			SQL = "DROP INDEX " & indexname & " ON " & tablename
			Conn.execute(SQL)
			if err.number <> 0 then 
				if debugging then call Report("ID",SQL,errorstop)
				DropIndex = false
			end if
			on error goto 0
		end if
	end function
	
	':onupdate=, :ondelete=
	'		CASCADE		delete or update the row from the parent table and automatically delete or update the matching rows in the child table
	'		SET NULL	delete or update the row from the parent table and set the foreign key column or columns in the child table to NULL.
	'		NO ACTION 	reject delete or update actions on parent table
	'		RESTRICT	same as NO ACTION
	public function AddForeignKey(tablename,columnname,byval foreignkeyname,references,foreignkeyoptions,skipifexist)
		if not connected then 
			if debugging then call Report("NC","",errorstop)
			AddForeignKey = false : exit function
		end if
		AddForeignKey = true
		
		dim foptions
		set foptions = parseOptions(foreignkeyoptions)
		if (trim(foreignkeyname&"")="") then foreignkeyname = "FK_" & tablename & "_" & columnname
		
		if (ForeignKeyExists(tablename,foreignkeyname)) then
			if skipifexist then exit function
			if not DropForeignKey(tablename,foreignkeyname) then 
				AddForeignKey = false
				exit function
			end if
		end if
		
		on error resume next
		
		SQL = "ALTER TABLE " & tablename & " ADD CONSTRAINT " & foreignkeyname & " FOREIGN KEY (" & columnname & ") " & _
			  "REFERENCES " & references
		if (database = "mysql") then
			if (trim(foptions("onupdate")&"")<>"") then
				SQL = SQL & " ON UPDATE " & ucase(foptions("onupdate"))
			else
				SQL = SQL & " ON UPDATE NO ACTION"
			end if
			if (trim(foptions("ondelete")&"")<>"") then
				SQL = SQL & " ON DELETE " & ucase(foptions("ondelete"))
			else
				SQL = SQL & " ON DELETE NO ACTION"
			end if
		end if
		Conn.execute(SQL)
		if err.number <> 0 then
			if debugging then call Report("FA",SQL,errorstop)
			AddForeignKey = false
		end if
		on error goto 0
		
	end function
	
	public function DropForeignKey(tablename,foreignkeyname)
		if not connected then 
			if debugging then call Report("NC","",errorstop)
			DropForeignKey = false : exit function
		end if
		DropForeignKey = true
		
		if (ForeignKeyExists(tablename,foreignkeyname)) then
			on error resume next
			if database = "mdb" then
			SQL = "ALTER TABLE " & tablename & " DROP CONSTRAINT " & foreignkeyname
			elseif database = "mysql" then
			SQL = "ALTER TABLE " & tablename & " DROP FOREIGN KEY " & foreignkeyname
			end if
			Conn.execute(SQL)
			if err.number <> 0 then 
				if debugging then call Report("FD",SQL,errorstop)
				DropForeignKey = false
			end if
			on error goto 0
		end if
	end function
	
	'/***********************************************************************************/
	'LOOK AT http://www.w3schools.com/ADO/met_conn_openschema.asp
	'adSchemaTables  			20 (Returns the tables defined in the catalog that are accessible)
	'		Constraints			TABLE_CATALOG,TABLE_SCHEMA,TABLE_NAME, TABLE_TYPE
	'adSchemaColumns  			4 (Returns the columns of tables defined in the catalog)
	'		Constraints			TABLE_CATALOG,TABLE_SCHEMA,TABLE_NAME,COLUMN_NAME
	'adSchemaProviderTypes  	22 		(Returns the data types supported by the data provider)
	'		Constraints  		DATA_TYPE, BEST_MATCH
	
	public function TableExists(tablename)
		if not connected then 
			if debugging then call Report("NC","",errorstop)
			TableExists = false : exit function
		end if
		Set rs = Conn.OpenSchema(20,Array(Empty,Empty,tablename,Empty))
		TableExists = not rs.eof
		rs.Close
	end function
	
	public function TablesOnDatabase()
		if not connected then 
			if debugging then call Report("NC","",errorstop)
			TablesOnDatabase = false : exit function
		end if
		dim return
		Set rs = Conn.OpenSchema(20,Array(Empty,Empty,Empty,Empty))
		
		if not rs.eof then
			while not rs.eof
				if (database <> "mdb" or lcase(left(rs("TABLE_NAME"),4)) <> "msys") then
				return = return & rs("TABLE_NAME") & "|"
				end if
				rs.movenext
			wend
			rs.Close
			TablesOnDatabase = split(left(return,len(return)-1),"|")
		else
			TablesOnDatabase = Array(null)
		end if
	end function
	
	public function ColumnExists(tablename,columnname)
		ColumnExists = false
		if not connected then 
			if debugging then call Report("NC","",errorstop)
			exit function
		end if
		Set rs = Conn.OpenSchema(4,Array(Empty,Empty,tablename,columnname))
		ColumnExists = not rs.eof
		rs.Close
	end function
	
	public function ColumnsOnTable(tablename)
		if not connected then 
			if debugging then call Report("NC","",errorstop)
			ColumnsOnTable = Array(null) : exit function
		end if
		dim return
		Set rs = Conn.OpenSchema(4,Array(Empty,Empty,tablename,Empty))
		
		if not rs.eof then
			while not rs.eof
				return = return & rs("COLUMN_NAME") & "|"
				rs.movenext
			wend
			rs.Close
			
			ColumnsOnTable = split(left(return,len(return)-1),"|")
		else
			ColumnsOnTable = Array(null)
		end if
	end function
	
	public function IndexExists(tablename,indexname)
		if not connected then 
			if debugging then call Report("NC","",errorstop)
			IndexExists = false : exit function
		end if
		if (database = "mdb") then
			Set rs = Conn.OpenSchema(10,Array(Empty,Empty,indexname,Empty,Empty,tablename,"UNIQUE"))
		elseif database = "mysql" then
			Set rs = Conn.Execute("SELECT * FROM information_schema.TABLE_CONSTRAINTS WHERE TABLE_CONSTRAINTS.TABLE_SCHEMA = '" & fString(currentdb) & "' AND TABLE_CONSTRAINTS.TABLE_NAME = '" & fString(tablename) & "' AND TABLE_CONSTRAINTS.CONSTRAINT_TYPE = 'UNIQUE' AND TABLE_CONSTRAINTS.CONSTRAINT_NAME = '" & fString(indexname) & "';")
		end if
		IndexExists = not rs.eof
	end function
	
	public function IndexesOnTable(tablename)
		if not connected then 
			if debugging then call Report("NC","",errorstop)
			IndexesOnTable = Array(null) : exit function
		end if
		dim return
		
		on error resume next
		if (database = "mdb") then
			Set rs = Conn.OpenSchema(10,Array(Empty,Empty,Empty,Empty,Empty,tablename,"UNIQUE"))
		elseif database = "mysql" then
			Set rs = Conn.Execute("SELECT * FROM information_schema.TABLE_CONSTRAINTS WHERE TABLE_CONSTRAINTS.TABLE_SCHEMA = '" & fString(currentdb) & "' AND TABLE_CONSTRAINTS.TABLE_NAME = '" & fString(tablename) & "' AND TABLE_CONSTRAINTS.CONSTRAINT_TYPE = 'UNIQUE';")
		end if
		
		if not rs.eof then
			while not rs.eof
				return = return & rs("CONSTRAINT_NAME") & "|"
				rs.movenext
			wend
			rs.Close
			
			IndexesOnTable = split(left(return,len(return)-1),"|")
		else
			IndexesOnTable = Array(null)
		end if
	end function
	
	public function ForeignKeyExists(tablename,foreignkeyname)
		if not connected then 
			if debugging then call Report("NC","",errorstop)
			ForeignKeyExists = false : exit function
		end if
		if (database = "mdb") then
			Set rs = Conn.OpenSchema(10,Array(Empty,Empty,foreignkeyname,Empty,Empty,tablename,"FOREIGN KEY"))
		elseif database = "mysql" then
			Set rs = Conn.Execute("SELECT * FROM information_schema.TABLE_CONSTRAINTS WHERE TABLE_CONSTRAINTS.TABLE_SCHEMA = '" & fString(currentdb) & "' AND TABLE_CONSTRAINTS.TABLE_NAME = '" & fString(tablename) & "' AND TABLE_CONSTRAINTS.CONSTRAINT_TYPE = 'FOREIGN KEY' AND TABLE_CONSTRAINTS.CONSTRAINT_NAME = '" & fString(foreignkeyname) & "';")
		end if
		ForeignKeyExists = not rs.eof
	end function
	
	public function ForeignKeysOnTable(tablename)
		if not connected then 
			if debugging then call Report("NC","",errorstop)
			ForeignKeyExists = Array(null) : exit function
		end if
		dim return
		
		if (database = "mdb") then
			Set rs = Conn.OpenSchema(10,Array(Empty,Empty,Empty,Empty,Empty,tablename,"FOREIGN KEY"))
		elseif database = "mysql" then
			Set rs = Conn.Execute("SELECT * FROM information_schema.TABLE_CONSTRAINTS WHERE TABLE_CONSTRAINTS.TABLE_SCHEMA = '" & fString(currentdb) & "' AND TABLE_CONSTRAINTS.TABLE_NAME = '" & fString(tablename) & "' AND TABLE_CONSTRAINTS.CONSTRAINT_TYPE = 'FOREIGN KEY';")
		end if
		
		if not rs.eof then
			while not rs.eof
				return = return & rs("CONSTRAINT_NAME") & "|"
				rs.movenext
			wend
			rs.Close
			
			ForeignKeysOnTable = split(left(return,len(return)-1),"|")
		else
			ForeignKeysOnTable = Array(null)
		end if
	end function
	
	'/***********************************************************************************/
	
	public function compatibleSQL(byval SQL)
		dim Matches,Match0,Match1
	
		'rende compatibile le DELETE
		Reg.Pattern = "^(DELETE(?: [*])?)"
		if Reg.Test(SQL) then
			set Matches = reg.Execute(SQL)
			Match0 = trim(ucase(Matches(0).SubMatches(0)))
			Set Matches = nothing
			if database = "mdb" and Match0 = "DELETE" then
				SQL =Reg.replace(SQL,"DELETE *")
			elseif database = "mysql" and Match0 = "DELETE *" then
				SQL = Reg.replace(SQL,"DELETE")
			end if
		end if
		
		'rende compatibile le TOP
		if database = "mdb" then
			Reg.Pattern = "(LIMIT) 0[,](\d+)[;]?$"
		elseif database = "mysql" then
			Reg.Pattern = "^SELECT (TOP) (\d+)"
		end if
		if Reg.Test(SQL) then
			set Matches = reg.Execute(SQL)
			Match0 = trim(ucase(Matches(0).SubMatches(0)))
			Match1 = trim((Matches(0).SubMatches(1)))
			Set Matches = nothing
			if database = "mdb" and Match0 = "LIMIT" then
				SQL = Reg.replace(SQL,"")
				Reg.pattern = "^SELECT"
				SQL = Reg.replace(SQL,"SELECT TOP " & Match1)
			elseif database = "mysql" and Match0 = "TOP" then
				SQL = Reg.replace(SQL,"SELECT") & " LIMIT 0," & Match1
			end if
		end if
		
		compatibleSQL = SQL
	end function
	
	'/***********************************************************************************/
	
	public function fDate(byval argDate)
		dim tmpdateoutput
		select case database
			case "mdb" : tmpdateoutput = "#yyyy-mm-dd hh:nn:ss#"
			case "mysql" : tmpdateoutput = "'yyyy-mm-dd hh:nn:ss'"
		end select
		tmpdateoutput = replace(tmpdateoutput,"dd",right("0" & day(argDate),2))
		tmpdateoutput = replace(tmpdateoutput,"mm",right("0" & month(argDate),2))
		tmpdateoutput = replace(tmpdateoutput,"yyyy",year(argDate))
		tmpdateoutput = replace(tmpdateoutput,"yy",right(year(argDate),2))
		tmpdateoutput = replace(tmpdateoutput,"hh",right("0" & hour(argDate),2))
		tmpdateoutput = replace(tmpdateoutput,"nn",right("0" & minute(argDate),2))
		tmpdateoutput = replace(tmpdateoutput,"ss",right("0" & second(argDate),2))
		
		fDate = tmpdateoutput
	end function
	
	public function fString(argValue)
		fString = replace(trim(argValue)&"","'","''")
	end function
	
	public function fNumber(argValue)
		on error resume next
		if argValue = "" or isnull(argvalue) then
			fNumber = cdbl(0)
		else
			fNumber = replace(cstr(cdbl(trim(argValue))),",",".")
		end if
		if err.number = 6 then fNumber = 0 'overflow
		if err.number = 13 then fNumber = 0 'tipo non corrispondente
		on error goto 0
	end function
	
	'/***********************************************************************************/
	
	private function parseOption(byval columnoptions, columnoption)
	
		dim Matches,return
		return = ""
		
		Reg.pattern = ":(" & columnoption & ")[ ]*=[ ]*((?:,(?![ ]*:)|[^,])+)"
		if reg.test(columnoptions) then
			set Matches = reg.Execute(columnoptions)
			return = trim(Matches(0).submatches(1))
		end if
		
		parseOption = return
	
	end function
	
	private function parseOptions(byval columnoptions)
	
		dim Matches,match,return
		set return = CreateObject("Scripting.Dictionary")
		return.CompareMode = 1
		
		Reg.pattern = ":([a-z0-9_-]+)[ ]*=[ ]*((?:,(?![ ]*:)|[^,])+)"
		if reg.test(columnoptions) then
			
			set Matches = reg.Execute(columnoptions)
			for each match in matches
				return(trim(match.submatches(0))) = trim(match.submatches(1))
			next
		end if
		
		'signed
		if return("signed") = "true" then : return("signed") = true : else : return("signed") = false : end if
		if return("null") = "false" then : return("null") = false : else : return("null") = true : end if
		if return("binary") = "true" then : return("binary") = true : else : return("binary") = false : end if
		
		set parseOptions = return
	
	end function
	
	'value = a, values = "|a|b|c|d|" => true
	private function inValue(value,values)
		inValue = false
		if instr(values,"|" & value & "|") > 0 then inValue = true
	end function
	
	'/***********************************************************************************/
	':name= column name
	':type= column type (look at compatibleColumnType)
	':size= column size (compatible with CHAR, VARCHAR)
	private function columnDef(byval columnoptions)
		dim coptions
		set coptions = parseOptions(columnoptions)
		columndef = compatibleColumnType(coptions("type"))
		'char/varchar size
		if inValue(coptions("type"),"|char|varchar|") and coptions("size") <> "" then
			columndef = columndef & "(" & coptions("size") & ")"
		end if
		
		'numbers options
		if inValue(coptions("type"),"|boolean|byte|short|long|currency|single|double|") then
			if database = "mysql" then
				'not supported in MDB
				if coptions("signed") then 
					columndef = columndef & " SIGNED"
				else
					columndef = columndef & " UNSIGNED"
				end if
			end if
			if coptions("null") = false then
				columndef = columndef & " NOT NULL"
				if coptions("default") = "" then coptions("default") = "0"
			end if
			if coptions("default") <> "" then columndef = columndef & " DEFAULT " & fNumber(coptions("default"))

		'texts options
		elseif inValue(coptions("type"),"|char|varchar|memo|") then
			if coptions("binary") = true then columndef = columndef & " BINARY"
			if coptions("null") = false then columndef = columndef & " NOT NULL"
			if coptions("default") <> "" or not coptions("null") then columndef = columndef & " DEFAULT '" & fString(coptions("default")) & "'"

		'others options
		elseif coptions("type") <> "counter" then
			if coptions("null") = false then columndef = columndef & " NOT NULL"
			if coptions("default") <> "" or not coptions("null") then columndef = columndef & " DEFAULT 0"
		end if
		
		
	end function
	
	private function compatibleColumnType(columntype)
		select case database
			case "mdb"
				select case columntype
					case "counter" 		: compatibleColumnType = "COUNTER UNIQUE NOT NULL PRIMARY KEY"
					case "boolean" 		: compatibleColumnType = "BIT"
					case "byte" 		: compatibleColumnType = "BYTE"
					case "date" 		: compatibleColumnType = "DATETIME"
					case "short" 		: compatibleColumnType = "SHORT"
					case "long" 		: compatibleColumnType = "LONG"
					case "single" 		: compatibleColumnType = "SINGLE"
					case "double" 		: compatibleColumnType = "DOUBLE"
					case "currency" 	: compatibleColumnType = "CURRENCY" 'fifteen integer, four decimal
					case "char" 		: compatibleColumnType = "TEXT"
					case "varchar" 		: compatibleColumnType = "TEXT"
					case "memo" 		: compatibleColumnType = "LONGTEXT"
					case "blob" 		: compatibleColumnType = "LONGBINARY"
				end select
			case "mysql"
				select case columntype
					case "counter" 		: compatibleColumnType = "INT(10) AUTO_INCREMENT UNIQUE NOT NULL PRIMARY KEY"
					case "boolean" 		: compatibleColumnType = "BOOLEAN"
					case "byte" 		: compatibleColumnType = "TINYINT(3)"
					case "date" 		: compatibleColumnType = "DATETIME"
					case "short" 		: compatibleColumnType = "SMALLINT"
					case "long" 		: compatibleColumnType = "INT"
					case "single" 		: compatibleColumnType = "FLOAT"
					case "double" 		: compatibleColumnType = "DOUBLE"
					case "currency" 	: compatibleColumnType = "DECIMAL(19,4)" 'fifteen integer, four decimal
					case "char" 		: compatibleColumnType = "CHAR"
					case "varchar" 		: compatibleColumnType = "VARCHAR"
					case "memo" 		: compatibleColumnType = "LONGTEXT"
					case "blob" 		: compatibleColumnType = "LONGBLOB"
				end select
		end select
	end function
	
	'/***********************************************************************************/
	
	private sub Report(what,exec,ending)
		dim title,description
		select case what
			case "NC" 'connection
			title = "Impossibile eseguire operazioni"
			description = "Attualmente non è attiva alcuna connessione."
			case "CN" 'connection
			title = "Impossibile connettersi"
			description = exec
			case "NA" 'new access
			title = "Impossibile creare un nuovo database"
			description = "path: " & exec & "<br/>controlla i permessi di scrittura"
			case "FK" 'delete access
			title = "Impossibile cancellare il file"
			description = "path: " & exec & "<br/>controlla i permessi di scrittura"
			case "TC" 'create table
			title = "Impossibile creare tabella"
			description = "<strong>sql:</strong> " & exec
			case "TD" 'drop table
			title = "Impossibile eliminare tabella"
			description = "<strong>sql:</strong> " & exec
			case "CA" 'add column
			title = "Impossibile creare campo"
			description = "<strong>sql:</strong> " & exec
			case "CD" 'drop column
			title = "Impossibile eliminare campo"
			description = "<strong>sql:</strong> " & exec & "<br/>forse c'è un indice o una foreign key. prova ad eliminare prima quelli!"
			case "CM" 'drop column
			title = "Impossibile modificare campo"
			description = "<strong>sql:</strong> " & exec & "<br/>forse vengono troncati dei dati?"
			case "IA" 'add index
			title = "Impossibile creare index"
			description = "<strong>sql:</strong> " & exec
			case "ID" 'drop index
			title = "Impossibile eliminare index"
			description = "<strong>sql:</strong> " & exec
			case "FA" 'add constraint
			title = "Impossibile creare foreign key"
			description = "<strong>sql:</strong> " & exec
			case "FD" 'drop constraint
			title = "Impossibile eliminare foreign key"
			description = "<strong>sql:</strong> " & exec
		end select
		response.write "<div style=""font-size:0.9em;font-family:Arial;background-color:#FFDFDF;color:#000;border:2px solid #800000;clear:both;margin:5px;padding:10px;""><strong>ASPDbManager ::: " & title & "</strong><br/>" & description & "</div>"
		if ending then response.end
	end sub
	
	'/***********************************************************************************/
	
end class

Class Class_ASPdBPagination
	
	private Reg,SQL
	private m_Page,m_Pages,m_Records,m_Record
	private m_RecordsPerPage
	
	public RecordsPerPageDefault
	
	public Recordset
	public debugging
	
	private sub Class_initialize()
	
		Set Reg = New RegExp 	
		Reg.Global = True
		Reg.Ignorecase = True
		
		debugging = false
		RecordsPerPageDefault = 10
		call Reset()
	end sub
	
	private sub Class_Terminate()
		call Reset()
	end sub
	
	'/***********************************************************************************/
	
   Public Property Get Page
		Page = m_Page
   End Property
   
   Public Property Get Pages
		Pages = m_Pages
   End Property
   
   Public Property Get Record
		Record = m_Record
   End Property
   
   Public Property Get Records
		Records = m_Records
   End Property
   
   Public Property Get RecordsPerPage
		RecordsPerPage = m_RecordsPerPage
   End Property
   
	'/***********************************************************************************/
   
   public sub Reset()
		m_Pages = 0
		m_Record = 0
		m_Records = 0
		m_RecordsPerPage = 10
		Recordset = null
   end sub
   
	'/***********************************************************************************/
	
	public function Paginate(byref Conn, byval database, recordsPerPage, Page, pagsql)
	
		Paginate = false
			
		dim godirect, rs
	
		call Reset()
		
		if isnull(recordsPerPage) then
			'from querystring
			m_RecordsPerPage = cleanLong(request.QueryString("perpage"))
			if m_RecordsPerPage = 0 then m_RecordsPerPage = RecordsPerPageDefault
		else
			m_RecordsPerPage = cleanLong(recordsPerPage)
		end if
		if m_RecordsPerPage < 0 then m_RecordsPerPage = RecordsPerPageDefault
		
		'check if no pagination
		if m_RecordsPerPage = 0 then
			godirect = true
			SQL = CompatibleLIMITandTOP(pagsql,database)
			if database = "mysql" then SQL = CompatibleFieldNames(SQL)
			m_Pages = 1
			m_Page = 1
			m_Records = 0
			m_Record = 1
		else
			godirect = false
			'*clean SQL
			SQL = pagsql
			SQL = CompatibleLIMITandTOP(SQL,database)
		
			if isnull(Page) then : m_Page = cleanLong(request.QueryString("page")) : else : m_Page = cleanLong(Page) : end if
			if m_Page <= 0 then m_Page = 1
		
			if trim(database&"") = "" then database = "mdb"
			database = lcase(database)
			
			dim SQLfirst, SQLsecond
			
			if database = "mysql" then
			
				'---------------------
				' PREPARE COUNT SQL
				'---------------------
				dim multiselectordistinct
				Reg.Pattern = "(UNION(?: ALL)?[( ]+SELECT|SELECT[ ]*DISTINCT)"
				multiselectordistinct = Reg.Test(SQL)
				
				if multiselectordistinct then
					SQLfirst = "SELECT Count(*) AS APRecords FROM (" & SQL & ") AS ATRecords"
				else
					Reg.Pattern = "SELECT ([^, ]+[ ,]+)+FROM"
					SQLfirst = Reg.Replace(SQL,"SELECT Count(*) AS APRecords FROM")
				end if
				
				'---------------------
				' EXEC COUNT SQL
				'---------------------
				
				m_Records = 0
				on error resume next
				set rs = conn.execute(SQLfirst) 'first query
				if debugging then call Report("1/2 Query preparatoria","<code>" & SQLfirst & "</code>",database,err.number <> 0)
				if err.number <> 0 then exit function 'sql error
				on error goto 0
				m_Records = clng(rs("APRecords"))
				
				'---------------------
				' SET NAVIGATION VARIABLE
				'---------------------
				m_Pages = excessEver(m_Records / m_RecordsPerPage)
				m_Record = ((m_Page - 1) * m_RecordsPerPage) + 1
				if m_Record > m_Records then
					m_Page = 1 : m_Record = 1
				end if
				lastRecord = m_Record + m_RecordsPerPage - 1
				if lastRecord > m_Records then lastRecord = m_Records
				
				'---------------------
				' PREPARE FINAL QUERY
				'---------------------
				'* second query (complete recordset, with LIMIT thanks to first query)
				SQLsecond = CompatibleFieldNames(SQL)
				if not(m_Records=0) then SQLsecond = SQLsecond & " LIMIT " & m_Record-1 & "," & m_RecordsPerPage
				
				on error resume next
				set rs = conn.execute(SQLsecond) 'second query
				if debugging then call Report("2/2 Query di selezione","<code>" & SQLsecond & "</code>",database,err.number <> 0)
				if err.number <> 0 then exit function 'sql error
				on error goto 0
				
				'* That's all for MYSQL
			
			elseif database = "mdb" then
				dim lastrecord
				
				SQLfirst = SQL
				set rs = Server.CreateObject("ADODB.Recordset")			
				on error resume next
				rs.open SQLfirst, Conn, 3, 3 'low performance query
				if debugging then call Report("1/1 Query di selezione","<code>" & SQLfirst & "</code>",database,err.number <> 0)
				if err.number <> 0 then exit function 'sql error
				on error goto 0
		
				'---------------------
				' SET NAVIGATION VARIABLE
				'---------------------
				
				m_Records = cint(rs.recordcount)
				m_Pages = excessEver(m_Records / m_RecordsPerPage)
				m_Record = ((m_Page - 1) * m_RecordsPerPage) + 1
				if m_Record > m_Records then
					m_Page = 1
					m_Record = 1'((m_Page - 1) * m_RecordsPerPage) + 1
				end if
				lastRecord = m_Record + m_RecordsPerPage - 1
				if lastRecord > m_Records then lastRecord = m_Records
				
				if not rs.eof then
					rs.pagesize = m_RecordsPerPage
					rs.absolutepage = m_Page
				end if
				
			end if
		
		
		end if
		
		if godirect then
			on error resume next
			set rs = conn.execute(SQL) 'second query
			if debugging then call Report("Query di selezione diretta","<code>" & SQL & "</code>",database,err.number <> 0)
			if err.number <> 0 then exit function 'sql error
			on error goto 0
		end if
		
		set Recordset = rs
		set rs = nothing
		
		Paginate = true
	
	end function
	
	'/***********************************************************************************/
	
	public function printNavigator(atext,add,queryadd,printperpage)
		dim text
		text = atext
		if trim(text&"")="" then text = "<div class=""s"">mostra record da %RS a %RE di %RT</div><div class=""n"">%PC pagine [ %PN ]</div>"
		dim lastrecord
		text = replace(text,"%RS",m_Record)
		if m_RecordsPerPage <=0 then
		lastrecord = 0
		text = replace(text,"%RT","tutti")
		text = replace(text,"%RE","ultimo")
		else
		lastrecord = m_Record+m_RecordsPerPage-1
		if (lastrecord>m_Records) then lastrecord = m_Records
		text = replace(text,"%RT",m_Records)
		text = replace(text,"%RE",lastrecord)
		end if
		
		text = replace(text,"%PC",m_Pages)
		
		dim pagenav
		pagenav = ""
		
		if m_Page > 1 then
			if m_Page > 4 then pagenav = pagenav & printNavigatorPage("&laquo;",1,queryadd,printperpage)
			pagenav = pagenav & printNavigatorPage("&lsaquo;",m_Page-1,queryadd,printperpage)
			if m_Page - 3 > 0 then pagenav = pagenav & printNavigatorPage(m_Page-3,m_Page-3,queryadd,printperpage)
			if m_Page - 2 > 0 then pagenav = pagenav & printNavigatorPage(m_Page-2,m_Page-2,queryadd,printperpage)
			if m_Page - 1 > 0 then pagenav = pagenav & printNavigatorPage(m_Page-1,m_Page-1,queryadd,printperpage)
		end if
		pagenav = pagenav & "<strong>" & m_Page & "</strong> "
		if m_Page < m_Pages then
			if m_Page + 1 < m_Pages + 1 then pagenav = pagenav & printNavigatorPage(m_Page+1,m_Page+1,queryadd,printperpage)
			if m_Page + 2 < m_Pages + 1 then pagenav = pagenav & printNavigatorPage(m_Page+2,m_Page+2,queryadd,printperpage)
			if m_Page + 3 < m_Pages + 1 then pagenav = pagenav & printNavigatorPage(m_Page+3,m_Page+3,queryadd,printperpage)
			pagenav = pagenav & printNavigatorPage("&rsaquo;",m_Page+1,queryadd,printperpage)
			if m_Page <= m_Pages - 4 then pagenav = pagenav & printNavigatorPage("&raquo;",m_Pages,queryadd,printperpage)
		end if
		
		text = replace(text,"%PN",pagenav)
				
		response.write "<form class=""pagenavigator"" action="""" method=""GET"">" & text & add & "</form>"
	end function
	
	private function printNavigatorPage(text, page,queryadd, printperpage)
		dim return
		return = "<a href=""?page=" & page
		if printperpage then return = return & "&amp;perpage=" & m_RecordsPerPage
		return = return & queryadd & """>" & text & "</a> "
		printNavigatorPage = return
	end function
	
	'/***********************************************************************************/
	
	private function excessEver(argValue)
		'(1.2 => 2, 1.6 => 2, 1.0 => 1)
		dim output : output = cdbl(argValue)
		dim isnegative : isnegative = (output<0) : if isnegative then output = output * (-1)
		if cbyte(right(int(output*10),1))>0 then : output = int(output)+1 : else : output = int(output) : end if
		if isnegative then output = output * (-1)
		excessEver = output
	end function
	
	private function cleanLong(argValue)
		on error resume next
		if argValue = "" then : cleanLong = clng(0) : else : cleanLong = clng(trim(argValue)) : end if
		if err.number = 6 then cleanLong = 0
		if err.number = 13 then cleanLong = 0
		on error goto 0
	end function
	
	'/***********************************************************************************/
	
	private sub Report(title,description,database,iserror)
		response.write "<div style=""font-size:0.9em;font-family:Arial;"
		if iserror then
			response.write "background-color:#FFDFDF;border:2px solid #800000;"
		else
			response.write "background-color:#DFFFDF;border:2px solid #008000;"
		end if
		response.write "color:#000;clear:both;margin:5px;padding:10px;"">"
		response.write "<strong>ASPDbPagination ::: "
		if iserror then : response.write "Error ::: " : else : response.write "Debug ::: " : end if
		response.write title & " [" & ucase(database) & "]</strong><br/>" & description & "</div>"
	end sub
	
	'/***********************************************************************************/
	
	public function CompatibleLIMITandTOP(SQL,database)
	
		'* removes all limit incompatible with ACCESS
		Reg.Pattern = "LIMIT \d+,\d+"
		SQL = Reg.Replace(SQL,"")
		
		'* compatible TOP
		if database = "mysql" then
			Reg.Pattern = "SELECT TOP (\d+)((?:.(?:(?![ )]+UNION(?: ALL)?[( ]+SELECT|[) ]+$)))+.)"
			if reg.test(SQL&" ") then
				CompatibleLIMITandTOP = reg.replace(SQL,"SELECT $2 LIMIT $1")
				exit function
			end if
		end if
		
		'* compatible LIMIT
		if database = "mdb" then
			Reg.Pattern = "SELECT ((?:.(?!LIMIT|UNION(?: ALL)?[( ]+SELECT|[) ]+$))+.)LIMIT (\d+)([^,]|$)"
			if reg.test(SQL) then
				CompatibleLIMITandTOP = reg.replace(SQL,"SELECT TOP $2 $1 $3")
				exit function
			end if
		end if
		CompatibleLIMITandTOP = SQL
	
	end function
	
	'remove FIRST and LAST (to be mysql compatible)
	public function CompatibleFieldNames(SQL)
		'*tested!!
		CompatibleFieldNames = SQL
		Reg.Pattern = "(?:first|last)\(([^), ]+)\)"
		if Reg.test(trim(SQL)) then CompatibleFieldNames = Reg.Replace(SQL,"$1")
	end function
	
	'/***********************************************************************************/
	
'	public function KeyWordsAfter(keyword,include)
'		'*tested!!
'		'* used in "(?:.(?!" & KeyWordsAfter & ")+.)"
'		'* es: WHERE => |SELECT ... FROM ...|
'		select case keyword
'			case "LIMIT"
'				KeyWordsAfter = "PROCEDURE"
'			case "ORDER BY"
'				KeyWordsAfter = "LIMIT|PROCEDURE"
'			case "HAVING"
'				KeyWordsAfter = "ORDER BY|LIMIT|PROCEDURE"
'			case "GROUP BY"
'				KeyWordsAfter = "HAVING|ORDER BY|LIMIT|PROCEDURE"
'			case "WHERE"
'				KeyWordsAfter = "GROUP BY|HAVING|ORDER BY|LIMIT|PROCEDURE"
'			case "FROM"
'				KeyWordsAfter = "WHERE|GROUP BY|HAVING|ORDER BY|LIMIT|PROCEDURE"
'			case else
'				KeyWordsAfter = "FROM"
'		end select
'		KeywordsAfter = KeywordsAfter & "|[) ]+UNION"
'		if include then KeyWordsAfter = keyword & "|" & KeyWordsAfter
'	end function
	
'	public function cleanRegString(value)
'		'*tested!!
'		Reg.Pattern = "([.\[\]()\\])"
'		cleanRegString = Reg.Replace(value,"\$1")
'	end function
	
end class
%>
