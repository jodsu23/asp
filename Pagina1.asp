<% option explicit %>
<html>
<head>
	<title>Problema</title>
</head><% option explicit %>
<html>
<head>
	<title>Problema</title>
</head>
<body>

	<% response.write("Hola Mundo!")
		'response.buffer=false
		'response.clear
	%>

	<br>

	<%
		dim fecha, dia
	
		fecha = date()
		dia = day(fecha)
		
		if dia <= 23 then
			response.write("Sitio Activo")
		else
			response.write("Sitio Fuera de Servicio")
		end if
	
	%>

	<br><br>
	
	<% response.write("Tipos de Variables") %>
	
	<%
		dim edad, pi, nombre, fechahoy, existe
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
	
	<br>
	
	<%  
		dim variable1, variable2, variable3, variable4
		variable1 = 5
		variable2 = 10
		variable3 = 15
		variable4 = variable1 + variable2 + variable3
		
		response.write("La sumatoria de las 3 variables es: ")
		response.write(variable4)
	
	%>
	
	<br><br>
	
	<% response.write("Formulario HTML (Control Text)") %>
	
	<br>
	
	<%
		dim v1, v2, suma
		v1 = request.form("valor1")
		v2 = request.form("valor2")
		
		suma = cint(v1) + cint(v2)
		response.write("La suma de los dos valores es: ")
		response.write(suma)
	%>
	
	<br><br>
	
	<%
		dim valLado, multi
		valLado = request.form("ladoA")
		multi = cint(valLado) * cint(valLado)
	
		response.write("El calculo del Cuadrado es: ")
		response.write(multi)
	%>
	
	<br><br>
	
	<% response.write("Formulario HTML (Control Select)") %>
	
	<br>
	
	<%
		dim v3, v4, sumar, operacion, resta
		v3 = request.form("valor3")
		v4 = request.form("valor4")
		operacion = request.form("operacion")
		
		'if(operacion="suma") then
		if(instr(operacion, "suma")<>0) then
			sumar = cint(v3) + cint(v4)
			response.write("La suma de los dos valores es:")
			response.write(sumar)
		end if
		'if(operacion="resta") then
		if(instr(operacion, "resta")<>0) then
			resta = cint(v3) - cint(v4)
			response.write("La diferencia de los dos valores es:")
			response.write(resta)
		end if
	%>
	
	<br><br>
	
	<% response.write("Estructura Condicionales Anidadas") %>
	
	<br>
	
	<%
		dim v
		v = cint(request.form("valor5"))
		
		if v < 10 then
			response.write("Tiene un dígito")
		else
			if v < 100 then
				response.write("Tiene dos dígito")
			else
				if v < 1000 then
					response.write("Tiene tres dígito")
				else 
					if v < 10000 then
					response.write("Tiene cuatro dígito")
					end if
				end if
			end if
		end if
	%>
	
	<br><br>
	
	<% response.write("Estructura Repetitiva For/Next") %>
	
	<br>
	
	<%
		dim x
		for x = 1 to 10
			response.write(x & "-")
		next
			response.write("<br>")
		for x = 1 to 10 step 2
			response.write(x & "-")
		next
	%>
	
	<br><br>
	
	<% response.write("Estructura repetitiva do/while/loop") %>
	
	<br>
	
	<%
		dim valor6, par
		valor6 = cint(request.form("valor6"))
		response.write("<h1>Valores pares hasta el nro " & valor6 & "</h1>")
		par = 2
		do while par <= valor6
			response.write(par & " - ")
			par = par + 2
		loop
		
	
	%>
	
	<br><br>
	
	<% response.write("Parametros en un Hipervinculo") %>
	
	<br>
	
	<%
		dim num, f
		num = request.querystring("tabla")
		response.write("<h1>Tabla de multiplicar del "& num &"</h1>")
		
		for f = num to num*10 step num
			response.write(f & "-")
		next
	%>
		<br>
	<%
		num = request.querystring("tabla5")
		for f = num to num*10 step num
			response.write(f & "-")
		next	
	%>
	
	<br><br>
	
	<% response.write("Insert") %>
	
	<br>
	
	<%
		'dim conexion
		'set conexion = Server.CreateObject("ADODB.Connection")
		'conexion.ConnectionString 	= "Provider=DESKTOP-J2EQRUJ\DEVELOPER;" & _
		'							  "Data Source=.;" & _
		'							  "Integrated Security=SSPI;" & _
		'							  "Initial Catalog=GenericDB"
		'response.Flush
		'conexion.Open
		'conexion.execute("insert into articulos(descripcion,precio)" & _
        '         "values ('" & request.form("descripcion") & _
        '         "'," & request.form("precio") & ")")
		'conexion.close
		
	%>
	
	<br><br>
	
	<% response.write("Redireccionamiento") %>
	
	<br>
	
	<%
		response.redirect("http://" & request.form("direccion"))
	%>
	
</body>
</html>
<body>

	<% response.write("Hola Mundo!") %>

	<br>

	<%
		dim fecha, dia
	
		fecha = date()
		dia = day(fecha)
		
		if dia <= 23 then
			response.write("Sitio Activo")
		else
			response.write("Sitio Fuera de Servicio")
		end if
	
	%>

	<br><br>
	
	<% response.write("Tipos de Variables") %>
	
	<%
		dim edad, pi, nombre, fechahoy, existe
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
	
	<br>
	
	<%  
		dim variable1, variable2, variable3, variable4
		variable1 = 5
		variable2 = 10
		variable3 = 15
		variable4 = variable1 + variable2 + variable3
		
		response.write("La sumatoria de las 3 variables es: ")
		response.write(variable4)
	
	%>
	
	
</body>
</html>