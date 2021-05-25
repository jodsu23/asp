<% option explicit %>
<html>
<head>
	<title>Problema Pagina 4</title>
</head>
<body>

	<%
		response.write("Nombre de Usuario cargado:" & session.contents("usuario") & "<br>")		
		response.write("Clave cargada:" & session.contents("clave") & "<br>")
	%>

	<br><br>
	
	<%
		function MultiplosTres(valor)
		dim numero, a
		numero = ""
			for a = a to valor step 3
				numero = numero & a & " - "
			next
			MultiplosTres = numero
		end function
		
		response.write("Multiplos de 3 hasta el 100</br>")
		response.write(MultiplosTres(100))
		response.write("<br><br>")
		
		function ValidarDatos()
			if isnumeric(request.form("valor1")) and isnumeric(request.form("valor2")) then
				ValidarDatos = true
			else
				ValidarDatos = false
			end if
		end function
		
		
		if ValidarDatos() then
			response.write("Primer Valor:" & request.form("valor1") & "<br>")
			response.write("Segundo Valor:" & request.form("valor2") & "<br>")
		else
			response.write("Deben ingresarse dos valores numericos")
		end if
	%>
	
	<br><br>
	
	<%
	
		sub MostrarSerie(inicio, fin)
			dim f
			for f = inicio to fin
				response.write(f & "<br>")
			next
		end sub
		
		MostrarSerie 5, 12
	
		sub CabeceraPagina()
			response.write("<div " & _
				"style=""background:#c3d9ff;text-align:center;font-size:40px"">")
			response.write("Cabecera de la Pagina")
			response.write("</div>")
		end sub
		
		sub FechaHora()
			dim fecha
			fecha = date()
			response.write("La hora es:" & fecha)
		end sub
		
		sub CuerpoPagina()
		  response.write("<div " & _
			"style=""background:#eeeeee;font-size:18px"">")
		  dim f
		  for f=1 to 20 
			response.write("Cuerpo de la pagina.<br>")
		  next
		  response.write("</div>")
		end sub
		
		sub PiePagina()
		  response.write("<div " & _
		  "style=""background:#cdeb8b;text-align:center;font-size:13px"">")
		  response.write("Pie de la pagina")
		  response.write("</div>")
		end sub
		
		'CabeceraPagina()
		'FechaHora()
		'CuerpoPagina()
		'PiePagina()
	%>

	<br><br><br>
	Add Include Libreria.asp
	<br>
	<!--#include file="libreria.asp"-->
	<%
		dim registros
		set registros = Server.CreateObject("ADODB.RecordSet")
		registros.open "select codigo, descripcion, precio from articulos", objConn
		
		do while not registros.eof
		response.write("Codigo :" & registros("codigo"))
		response.write("<br>")
		response.write("Descripcion: " & registros("descripcion"))
		response.write("<br>")
		response.write("Precio: " & registros("precio"))
		response.write("<br>")
		response.write("------------------------------")
		response.write("<br>")
		registros.movenext
		loop
		objConn.close
	%>

</body>
</html>