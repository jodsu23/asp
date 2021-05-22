<html>
<head>
	<title>Problema</title>
</head>
<body>

	<% response.write("Hola Mundo!") %>

	<br>

	<%
		fecha = date()
		dia = day(fecha)
		
		if dia <= 23 then
			response.write("Sitio Activo")
		else
			response.write("Sitio Fuera de Servicio")
		end if
	
	%>


	<br>
	
	<% response.write("Tipos de Variables") %>
	
	<%
		edad 	= 22
		pi		= 3.1416
		nombre	= "Canela"
		fechahoy= #12/25/2008#
		existe	= true
		
		response.write("Variable Entera: ")
		response.write(edad)
		
		response.write("<br>")
		
		response.write("Variable Real: ")
		response.write(pi)
		
		response.write("<br>")
		
		response.write("Variable Cadena: ")
		response.write(nombre)
		
		response.write("<br>")
		
		response.write("Variable Fecha: ")
		response.write(fechahoy)
		
		response.write("<br>")
		
		response.write("Variable logica: ")
		response.write(existe)
		
	
	%>
	
	
</body>
</html>