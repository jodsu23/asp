<% option explicit %>
<html>
<head>
	<title>Problema Pagina 3</title>
</head>
<body>
	
	<% response.write("Redireccionamiento") %>
	
	<br>
	
	<%
		response.redirect("http://" & request.form("direccion"))
	%>
	
	<br><br>
	
	<% response.write("Variable de Sessions") %>
	
	<br>
	
	<%
		session.contents("usuario") = request.form("usuario")
		session.contents("clave") = request.form("clave")
	%>
	
	<a href="Pagina4.asp">Ir a la otra Pagina 4</a>
	
</body>
</html>