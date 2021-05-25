<%
		dim objConn
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.ConnectionString = "Provider=SQLOLEDB; Data Source=DESKTOP-J2EQRUJ\DEVELOPER; Database=GenericDB; User ID=usuario2; Password=usuario2"
		objConn.Open
%>