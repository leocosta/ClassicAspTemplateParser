<%
'*********************************************************************
'** Titulo.............: Classe Template
'** Nome do arquivo....: class.template.asp
'** Autor..............: Leonardo Costa (email: leotasco@gmail.com)
'** Versao.............: 1.3
'** Data de Criacao....: 03/08/2001
'** Ultima modificacao.: 16/08/2005
'*********************************************************************

CONST BREAK = "¹"

Class Template 'iniciando a classe template
  
  Public vEnd
  Public vStart
  Public var_names  'array    (guarda variaveis registradas, id_do_arquivo)
  Public vars				'array    (guarda variaveis definidas no template)
  Public files      'array    (guarda id_do_arquivo, conteudo_arquivo, tamanho_arquivo )
  Public Registreds 'array    (guarda as variaveis registradas com a funcao register)
  
  Private Sub Class_Initialize 
     vStart = "{" : vEnd = "}" 'Setando variaveis na inicializacao da classe
  End Sub

  ' Função para carregar um template numa classe
  ' Parametros file_id, var
  ' Onde: file_id = "chave referente ao arquivo de template carredado" 
  '       filename = "caminho/nome do arquivo de template a ser carregado"
  Public Sub load_file( file_id, filename )
	Dim ObjFso, contentFile, ObjAbreArq, FileSize, nPos, Found, C, EndPos
	Set ObjFso = Server.CreateObject("Scripting.FileSystemObject")

		If not ObjFso.FileExists(filename) then 
			 response.write "<b>ERRO:</b> Arquivo de template inexistente (" &filename& ")"
			 response.end
		End if

		Set ObjAbreArq = ObjFso.OpenTextFile(filename,1, false, false)
		If ObjAbreArq.AtEndOfStream = True Then 
			Response.Write "O Arquivo de template esta vazio!"
		Else
			contentFile = ObjAbreArq.ReadAll()
			set filesize = ObjFso.GetFile(filename)

			If Isarray(files) Then
				 Redim Preserve files(file_id)
			Else 
				 Redim files(file_id)
			End if
			
			files(file_id) = array( contentFile, filesize.size )
			get_var( file_id )

			ObjAbreArq.Close
			Set ObjAbreArq = Nothing
			Set ObjFso = Nothing
		End if	
	
		Call Include_File(File_ID, Verify_IncludeFileName(File_ID))		

  End Sub

  Public Sub Include_File( file_id, filename )
		Dim nfilename, ObjFso, ObjAbreArq, include, ntag, contentFile, pos, tag
		If file_id = "" Or filename = "" Then
			 'response.write "ERRO: Parâmetro solicitado faltando"
			 exit sub
		End if
		
		nFilename = server.MapPath(filename)
		Set ObjFso = Server.CreateObject("Scripting.FileSystemObject")

		If ObjFso.FileExists(nfilename) then

			Set ObjAbreArq = ObjFso.OpenTextFile(nfilename,1, false, false)
			contentFile = ObjAbreArq.ReadAll()
			include = contentFile

			ntag = "<include filename=" & chr(34) & filename & chr(34) & ">"
			pos = instr(files(file_id)(0), ntag)
			If pos > 0 Then
				 tag = mid( files(file_id)(0), pos, len(ntag) )
				 files(file_id)(0) = replace( files(file_id)(0), tag, include )
			End If
			ObjAbreArq.Close
		Else
			include = "ERRO: Arquivo " & filename & " não existe !"
		End If
		
		Set ObjAbreArq = Nothing

		Call Include_File(File_ID, Verify_IncludeFileName(File_ID))				
  
  End Sub

	Private Function Verify_IncludeFileName(File_ID)

		Dim PosI, PosF, tmpFileName, tmpTagI, tmpTagF

		tmpTagI = "<include filename=" & chr(34)
		tmpTagF = chr(34) & ">"

		PosI = InStr(LCase(Files(File_ID)(0)),tmpTagI)
		If (PosI>0) Then
			PosI = PosI + Len(tmpTagI)
			PosF = InStr(PosI, LCase(Files(File_ID)(0)),tmpTagF)
			tmpFileName = Mid(Files(File_ID)(0),PosI, PosF-PosI)
			'----------
			Verify_IncludeFileName = tmpFileName
		End If	

	End Function
  
  ' Função que registra as variáveis em um array 
  ' Parametros file_id, var_name
  ' Onde: file_id = "chave referente ao arquivo de template" 
  '       filename = "string com nome das variáveis que serão substituidas"
  '       obs.: os nomes devem ser separados por vírgula 
  '             ex.:  register 0, "nome, endereco, estado, telefone"
  Public Sub register( file_id, vars )
	Dim vetVars, contRegister, Pos, registerVariable

	if Isarray(vars) then
		vetVars = vars
	else
		vetVars = Split(vars, ",")
	end if
	
	for contRegister = 0 to Ubound(vetVars)
		If IsArray(Registreds) then
			Pos = Ubound(Registreds) + 1
			Redim Preserve Registreds(pos)
		else
			Redim Registreds(0)
		end if	
		registerVariable = Eval(vetVars(contRegister))
		if IsNull(registerVariable) then
			registerVariable = ""
		end if
		Registreds(pos) = Array(vetVars(contRegister), registerVariable)
	next
  End Sub
  
  ' Função substituir elementos passados pelo parametro tag no arquivo file_id0 pelo conteudo
  ' do arquivo passado por file_id1 .
  ' Parametros: file_id0, file_id1, tag
  Public Sub concat_class( file_id0, file_id1, tag )
	 files(file_id0)(0) = replace( files(file_id0)(0), tag, files(file_id1)(0)) 
  End Sub
  
  ' Função que substitui variaveis decaradas no template pelo conteudo definido no codigo asp
  ' Parametros: file_id, vars
  ' Onde: file_id - é o arquivo de template que devera ser substituido
  '       vars - variáveis passadas que deverao conter no template para serem substituidas
  ' Dim nome, endereco
  ' nome     = "Maria"
  ' endereco = "Rua das Flores, 123"
  'Exemplo:
  ' tpl.register "nome, endereco"
  ' tpl.parse 0
  Public Sub parse( file_id )
	Dim contParse
	If file_id <> "" Then
		For contParse = 0 to Ubound(Registreds)
			files(file_id)(0) = Replace(files(file_id)(0), vstart & Trim(Registreds(contParse)(0)) & vend, Registreds(contParse)(1))
		Next
		Registreds = ""
	End if
  End sub

  ' Imprime o conteudo do arquivo de template
  Public Sub print_file( file_id )
     Dim vet, i
	If instr( file_id, "," ) <> 0 Then
	    vet = split(file_id,",")
		for i = 0 to ubound(vet)
		   response.write files(trim(vet(i)))(0)
		next
	 Else
	    response.write files(file_id)(0)
	 End If
  End Sub

  ' Retorna o conteudo do arquivo de template
  Function return_file( file_id )
	Dim Ret
	ret = ""
	If instr( file_id, "," ) <> 0 Then
		vet = split(file_id,",")
		for i = 0 to ubound(vet)
			ret = ret & files(trim(vet(i)))(0)
		next
	Else
		ret = files(file_id)(0)
	End If
	return_file = ret
  End Function

  Public Function parse_if( file_id, result_name, kill )
	Dim sTag, eTag, Start_Pos, End_Pos, If_Code, Start_Tag, End_Tag

	stag = "<if name=" & chr(34) & result_name & chr(34) & ">"
	etag = "</if name=" & chr(34) & result_name & chr(34) & ">"

	start_pos = instr( lcase( files(file_id)(0) ), lcase(stag) ) + len(stag)
  	end_pos = instr( lcase( files(file_id)(0) ), lcase(etag) )
	
	if instr(files(file_id)(0), stag) <> 0 Then
		if_code = mid( files(file_id)(0), start_pos, end_pos - start_pos )
		start_tag = mid( files(file_id)(0), instr(lcase(files(file_id)(0)), lcase(stag)), len(stag) )
		end_tag = mid( files(file_id)(0), instr(lcase(files(file_id)(0)), lcase(etag)), len(etag) )
	else
		exit function
	end if	

	If if_code <> "" Then
		If kill then
			files(file_id)(0) = replace( files(file_id)(0), start_tag & if_code & end_tag, "" )
			if Instr( files(file_id)(0), "<if name=" & chr(34) & result_name & chr(34) & ">" ) <> 0 then
				parse_if file_id, result_name, kill
			end if 
		Else
			files(file_id)(0) = replace( files(file_id)(0), start_tag & if_code & end_tag, trim(if_code) )
			if Instr( files(file_id)(0), "<if name=" & chr(34) & result_name & chr(34) & ">" ) <> 0 then
				parse_if file_id, result_name, kill
			end if 
		End If
	End If
  End Function

  ' Recebe um array e substitui as posicoes do array definidas no template 
  ' Ex.: tpl.parse_loop 0, "flag", "nome, endereco"
  '      Onde "nome, endereco" sao variaveis do html e serao substituidas por seus 
  '      respectivos valores.  
  Public Sub parse_loop( file_id, result_name, array_name )

	Dim sTag, eTag, lower_result_name, start_pos, end_pos, loop_code, start_tag, end_tag, _
	new_code, name, x, i, Temp_Code, Row_data

	if not IsArray(array_name) then
		array_name = template_vetor(array_name)
	end if
	lower_result_name = lcase(  result_name )
	
	stag = "<loop name=" & chr(34) & lower_result_name & chr(34) & ">"
	etag = "</loop name=" & chr(34) & lower_result_name & chr(34) & ">"
	start_pos = instr( lcase( files(file_id)(0) ), stag ) + len(stag)
 	end_pos = instr( lcase( files(file_id)(0) ), etag )
	loop_code = mid( files(file_id)(0), start_pos, end_pos - start_pos )
	start_tag = mid( files(file_id)(0), instr(lcase(files(file_id)(0)), stag), len( stag ) )
	end_tag = mid( files(file_id)(0), instr(lcase(files(file_id)(0)), etag), len( etag ) )

	If loop_code <> "" Then
		new_code = ""
		name = 0
		if Isarray(array_name(0,1)) then
			for x = 0 To ubound( array_name(name,1) )
				name = 0
				temp_code = loop_code
				For i = 0 to ubound ( array_name )
					If UBound(array_name(i,1)) = -1 Then
						row_data = ""
					Else
						row_data = array_name(i,1)(x)
					End if
					temp_code = replace( temp_code, vstart & array_name(i,0) & vend, row_data )
				next
				name = name + 1
				new_code = new_code & temp_code
			next
		End if	
		files(file_id)(0) = replace( files(file_id)(0), start_tag & loop_code & end_tag, new_code )
		'----------
		Call ErrorHandler("&lt;loop name=""" & result_name & """&gt;")
		'----------
		if instr(files(file_id)(0),"<loop name=" & chr(34) & result_name & chr(34) & ">") then
			parse_loop file_id, result_name, array_name
		end if
	End If

  End Sub

	Private Sub ErrorHandler(Local)
		If ((Not Err.Number=0) And (Not Err.Number=5)) Then
			Response.Clear
			'----------
			Response.Write "<B>Erro</B>&nbsp;" & Err.Number & "<BR>"
			Response.Write "<B>Origem</B>&nbsp;" & Err.Source & "<BR>"
			Response.Write "<B>Descição</B>&nbsp;" & Err.Description & "<BR>"
			Response.Write "<B>Local</B>&nbsp;" & Local & "<BR>"
			'----------
			Response.End
		End If
	End Sub
	
	Public Sub parse_mssql( file_id, result_name )

	Dim sTag, eTag, Start_Pos, End_Pos, If_Code, Lower_Result_Name, Start_Tag, End_Tag
	Dim Loop_Code, Rsparse_mssql, new_code, field_names, newIndex, Temp_Code, x, Row_Data

	lower_result_name = lcase( result_name )

	stag = "<loop name=" & chr(34) & lower_result_name & chr(34) & ">"
	etag = "</loop name=" & chr(34) & lower_result_name & chr(34) & ">"
	start_pos = instr( lcase( files(file_id)(0) ), stag ) + len(stag)
  	end_pos = instr( lcase( files(file_id)(0) ), etag )
	if start_pos <> 0 And end_pos <> 0 Then
		loop_code = mid( files(file_id)(0), start_pos, end_pos - start_pos )
		start_tag = mid( files(file_id)(0), instr(lcase(files(file_id)(0)), stag), len( stag ) )
		end_tag = mid( files(file_id)(0), instr(lcase(files(file_id)(0)), etag), len( etag ) )
	Else
	    exit sub
	End If	
	
	set rsparse_mssql = eval(result_name)
	If loop_code <> "" Then
		new_code = ""
		field_names = array()
		for i = 0 to rsparse_mssql.fields.count-1
			newIndex = UBound(field_names)+1
			redim preserve field_names( newIndex  )
			field_names(newIndex) = rsparse_mssql.fields(i).name
		next
		
		do while not rsparse_mssql.eof
			temp_code = loop_code
			For x = 0 to ubound ( field_names  )
				row_data = rsparse_mssql ( field_names(x) ) 
				if Isnull(row_data) then
					row_data = " "
				end if
				
				if instr(temp_code, vstart & field_names(x) & vend) <> 0 then
					temp_code = replace( temp_code, vstart & field_names(x) & vend, row_data )
				end if	
			next
			new_code = new_code & temp_code
			rsparse_mssql.movenext
		loop
		files(file_id)(0) = replace( files(file_id)(0), start_tag & loop_code & end_tag, new_code )
	End If
	
	set rsparse_mssql = nothing
	'----------
	Call ErrorHandler("&lt;sql name=""" & result_name & """&gt;")
	
  End Sub

  Private Sub get_var( file_id )
	Dim sTag, eTag, Start_Pos, End_Pos, If_Code, Lower_Result_Name, Start_Tag, End_Tag
	Dim var_code, real_var_code, cab, posi, posf, vars_temp

	stag = "<vars>"
	etag = "</vars>"
	start_pos = instr( lcase( files(file_id)(0) ), stag ) + len(stag)
  	end_pos = instr( lcase( files(file_id)(0) ), etag )
	if instr(files(file_id)(0), stag) <> 0 Then
		var_code = mid( files(file_id)(0), start_pos, end_pos - start_pos )
		start_tag = mid( files(file_id)(0), instr(lcase(files(file_id)(0)), stag), len( stag ) )
		end_tag = mid( files(file_id)(0), instr(lcase(files(file_id)(0)), etag), len( etag ) )
	else
		exit sub
	end if

	real_var_code = var_code
	var_code = replace(var_code, vbCrlf, "" )
	do while instr(var_code, "[") <> 0
		cab = mid(var_code, instr(var_code, "["), instr(var_code, "]") - instr(var_code, "[") + 1 )
		posi = instr( var_code, cab )
		var_code = replace(var_code, cab, "" )
		posf = instr( var_code, "[")-1
		if posf <= 0 then
			posf = end_pos
		end if
		if not posf <= 0 and not posi <= 0 then
			vars_temp = cab & ";" &mid(var_code, posi, posf)
			if Isarray(vars) then
				redim preserve vars( ubound(vars)+1 )
			else
				redim vars(0)
			end if
			vars(ubound(vars)) = split(vars_temp, ";")
		end if
	Loop

	files(file_id)(0) = replace( files(file_id)(0), start_tag & real_var_code & end_tag, "" )

   end sub
   
   'Pega o Vars pelo nome definido entre as tags Var
   '<VARS>
   '[mensagem]
   'Ocorreu um erro;	
   'Incluido com sucesso;
   '</VARS>
   '<HTML>
   '<BODY>
   '</BODY>
   '</HTML>
   '
   'Ex.: tpl.GetVars 0, "[mensagem]", 2
   'Retorna -> Incluido com sucesso
   Public Function GetVars( name, position)
	 dim x	

	For x = 0 to Ubound(Vars)
		If UCase(Vars(x)(0)) = UCase(name) then
			If position > 0 And position <= Ubound(Vars(x)) then
				GetVars = Vars(x)(position)
				exit for
			else
				response.write "Você está tentando acessar uma posição que não existe !"
				exit function
			end if	
		End if
	Next
   End Function

   Public Function GetNextFileID()
		 GetNextFileID = 0
		 If IsArray(Files) Then
			 GetNextFileID = Ubound(Files)+1
		 End If
   End Function

	 Public Function GetFile(Type_File)
			
			Dim Path_File, aParameters, Language
			
			If(Type_File="") Then
				Type_File = "html"
			End If

			Type_File = LCase(Type_File)	

			If (InStr(Type_File, ";")) Then
				aParameters = Split(Type_File, ";")
				Type_File = aParameters(0)
				Language = aParameters(1) & "_"
			ElseIf (InStr(Type_File, "/")) Then
				Path_File = Replace(Type_File,"/","\")
				Type_File = "virtual"
			End If

			Select Case Type_File
				
				Case "html"

					Path_File = LCase(Replace(Request.ServerVariables("PATH_TRANSLATED"),Server.MapPath("."),""))
					Path_File = Replace(Path_File,".asp",".htm")
					If Len(Language) Then
						Path_File = Replace(Path_File,"\", "\" & Language)
					End If
					Path_File = Server.MapPath(".") & "\templates" & Path_File

				Case "virtual"
					
					Path_File = Server.MapPath("\") & Replace(Path_File,".asp",".htm")

				Case Else
					
					Path_File = Server.MapPath(".") & "\templates\" & Type_File

			End Select
			
			GetFile = Path_File

	 End Function
   
	 Function parse_internal_loop( file_id, result_name, if_result_name_false, if_result_name_true, array_name  )
	Dim sTag, eTag, Start_Pos, End_Pos, If_Code, Lower_Result_Name, Start_Tag, End_Tag

	array_name = template_vetor(array_name)
	lower_result_name = lcase(  result_name )
	lower_if_result_name_false = lcase(  if_result_name_false )
	lower_if_result_name_true = lcase(  if_result_name_true )
	
	stag = "<loop name=" & chr(34) & lower_result_name & chr(34) & ">"
	etag = "</loop name=" & chr(34) & lower_result_name & chr(34) & ">"
	start_pos = instr( lcase( files(file_id)(0) ), stag ) + len(stag)
  	end_pos = instr( lcase( files(file_id)(0) ), etag )

	loop_code = mid( files(file_id)(0), start_pos, end_pos - start_pos )
	
	start_tag = mid( files(file_id)(0), instr(lcase(files(file_id)(0)), stag), len( stag ) )
	end_tag = mid( files(file_id)(0), instr(lcase(files(file_id)(0)), etag), len( etag ) )

	If loop_code <> "" Then
		if Isarray(array_name(0,1)) then
			for x = 0 To ubound( array_name(name,1) )
				name = 0
				temp_code = loop_code

				For i = 0 to ubound( array_name )
					row_data = array_name(i,1)(x)
					temp_code = replace( temp_code, vstart & array_name(i,0) & vend, row_data )
				next
			
				name = name + 1
				new_code = new_code & temp_code
			next
		end if
	End if

	files(0)(0) = replace( files(file_id)(0), start_tag & loop_code & end_tag, new_code )
	parse_if file_id, if_result_name_false, true
	parse_if file_id, if_result_name_true, false

  End function

  'Funcao auxiliar ao parse_loop e parse_internal_loop
  Private Function Template_Vetor( vars ) 
     Dim varsTemplate_vetor, countLines, contForTemplate_vetor

	vars = replace(vars, " ", "")
   
     varsTemplate_vetor = split( vars,"," )
     countLines = UBound( varsTemplate_vetor )
     ReDim vetReturnTemplate_Vetor( countLines, 1 )
	
     For contForTemplate_vetor = 0 To countLines
	   If Len(Eval(varsTemplate_vetor(contForTemplate_vetor))) Then
		   vetReturnTemplate_Vetor(contForTemplate_vetor, 0) = varsTemplate_vetor(contForTemplate_vetor)
		   vetReturnTemplate_Vetor(contForTemplate_vetor, 1) = Split(Left(Eval(varsTemplate_vetor(contForTemplate_vetor)),Len(Eval(varsTemplate_vetor(contForTemplate_vetor)))-1), BREAK )
	   End If
     Next
     Template_Vetor = vetReturnTemplate_Vetor
  End Function
  '--------------------------------------------------------------------------------------------
  ' Funções de atalho
  '--------------------------------------------------------------------------------------------
  Public Sub PPrint( file_id, replacements )
  	parse file_id, register( replacements )
	print_file( file_id )
  End Sub

  Public Sub PParse( file_id, replacements )
	register file_id, replacements
	parse file_id
  End Sub

  Function PGet( file_id, replacements )
  	parse file_id, register( replacements )
	pget = files(file_id)(0)
  End Function
  '--------------------------------------------------------------------------------------------
  ' Fim das Funções de atalho
  '--------------------------------------------------------------------------------------------
  Private Sub Class_Terminate   ' Setup Terminate event.
  End Sub

End Class 'fim da classe template
%>
