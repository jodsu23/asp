<%

		dim objConn
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.ConnectionString = "Provider=SQLOLEDB; Data Source=DESKTOP-J2EQRUJ\DEVELOPER; Database=GenericDB; User ID=usuario2; Password=usuario2"
		objConn.Open

		'if IsObject(objConn) then
		'	response.write"The coneccion is active <br />"
			'if Con.State = 1 then
			'	response.write"A connection is made, and is open"
			'end if
		'end if
		
		objConn.execute("insert into articulos(descripcion,precio)" & _
                 "values ('" & request.form("descripcion") & _
                 "'," & request.form("precio") & ")")
				 
		response.write("Se han guardado los datos con exito!")
		objConn.close
		
	%>