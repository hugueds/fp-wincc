
'Script para validacao de login de usuario no posto
'Author : Hugo SSBHPE, Rafael SSBROX
'Created: Nov/2015
'Last Significative Modification: Out/2016, Jan/2017

'Impotant Tags
'trigger : 
'1 = Solicitacao, 
'2 = Aguardando Leitura, 
'3 = Leitura Ok, 
'4 = Leitura Adm Ok
'5 = Leitura Colaborador Ok
'6 = Leitura Lider Ok

'trigger errors : -1 = timeout, -2 = contagem invalida, -3 = identificado/nao treinado, -4 = n�o identificado, -5 nivel abaixo Do requisitado

Dim conn
Dim rstWorkplace, rstTerminal, rstEnter, rstTraining
Dim tagAccessLevel,tagTrigger, tagUserReg, tagUserSSB, tagIdWorkstation, tagIdentified, tagLogged, tagTrainingLevel
Dim tagAdminSSB, tagAdminReg, tagAdminIdentified, tagAdminLogged, tagAdminAcessLevel 
Dim sqlDbTag
Dim idTerminal
Dim timeCounter, invalidCounter
Dim timeCounterTag, invalidCounterTag
Dim tEnterStatus, tEnterValid
Dim userId, userReg, userSSB, idWorkstation, adminLevel, trainLevel
Dim popid
Dim securityCounter

On Error Resume Next	

'TraceMsg "------ Function Biometria Terminal, Station: " & CStr(station) & ", Position: " & CStr(workplace) & vbLf	

'Get the trigger value
If station > 0 Then 
	Set tagTrigger = HMIRuntime.Tags("DB_OPERADOR_POSTO[" & CStr(index) & "]_CMD_WR_Trigger")
	tagTrigger.Read()
	trigger = tagTrigger.Value
End If

If trigger = 3 Then
	biometriaTerminal = 0		
	Exit Function
End If

If trigger <= 0 Or trigger > 10 Then
	biometriaTerminal = 0
	TraceMsg "------ Function Biometria Terminal, exiting...  trigger <= 0 || trigger >= 10 " & vbLf	
	Exit Function
End If

If station = "" Or  position = "" Then
	biometriaTerminal = 0
	TraceMsg "------ Function Biometria Terminal, station or position empty  " & vbLf	
	Exit Function
End If


'Get the Database from the configuration
Set sqlDbTag = HMIRuntime.Tags("SQL_DATABASE")
sqlDbTag.Read()
sqlDb = sqlDbTag.Value

'Create the Recordset objects (one for each table)
Set conn = CreateObject("ADODB.Connection")
Set rstWorkplace = CreateObject("ADODB.Recordset")	
Set rstEnter = CreateObject("ADODB.Recordset")		

conn.Open "Provider=SQLOLEDB;Initial Catalog=" & "LTS" & ";Data Source=" & sqlDb & ";User ID=LTS;Password=lts123;" 

'Error routine for connection
If Err.Number <> 0 Then
	TraceMsg "Error Opening Database #" & Err.Number & " " & Err.Description
	Err.Clear
	Set conn = Nothing  				
	biometriaTerminal = -1 ' Not Ok
	Exit Function
End If		
	
'Execute - Select Id Workstation
Set rstWorkplace = conn.Execute("SELECT W.ID FROM TB_WORKPLACE W ,TB_STATION S WHERE W.ID_STATION=S.ID AND S.NAME = '" & station & "' AND W.NAME = '" & position & "'")

'Error routine
If Err.Number <> 0 Then
	TraceMsg "Error #" & Err.Number & " " & Err.Description
	Err.Clear
	'Close data source 
	conn.close
	Set conn = Nothing
	Set rstWorkplace = Nothing 	
	biometriaTerminal = -1 ' Not Ok
	Exit Function
End If	

idWorkstation = rstWorkplace.Fields("ID").Value

'Execute - Select Id Terminal
Set rstTerminal = conn.Execute("SELECT IdTerminal FROM TB_Terminal WHERE IdWorkstation = '" & idWorkstation & "'")

'Error routine 
If Err.Number <> 0 Then
	TraceMsg "Error #" & Err.Number & " " & Err.Description
	Err.Clear
	'Close record set
	rstWorkplace.Close	
	'Close data source
	conn.close
	Set conn = Nothing
	Set rstWorkplace = Nothing
	Set rstTerminal = Nothing 	
	biometriaTerminal = -1 ' Not Ok	
	Exit Function
End If		

idTerminal = rstTerminal.Fields("IdTerminal").Value

Set timeCounterTag = HMIRuntime.Tags("DB_OPERADOR_POSTO[" & CStr(index) & "]_CMD_R_timeCounter")	
timeCounterTag.Read()
timeCounter = timeCounterTag.Value 
	
Set invalidCounterTag = HMIRuntime.Tags("DB_OPERADOR_POSTO[" & CStr(index) & "]_CMD_R_invalidCounter")	

If trigger = 2 Then	'Enviado para IHM 

	Set rstEnter = conn.Execute("SELECT * FROM tEnter WHERE	L_TID=" & CStr(idTerminal) & "ORDER BY ID")
	
	'Error routine
	If Err.Number <> 0 Then
		TraceMsg "Error #" & Err.Number & " " & Err.Description
		Err.Clear

		'Close record set
		rstWorkplace.Close	
		rstTerminal.Close			
		'Close data source - Datenquelle schlie?en
		conn.close			
		Set conn = Nothing
		Set rstWorkplace = Nothing
		Set rstTerminal = Nothing 	
		Set rstEnter = Nothing
		biometriaTerminal = -1 ' Not Ok	
		Exit Function
	End If				

	tEnterValid = 0
	invalidCounter = 0
	securityCounter = 0
	
	Do While Not(rstEnter.EOF) 'And tEnterValid = 0 And securityCounter < 100 'for each tEnter event
	
		tEnterStatus  = rstEnter.Fields("L_Result").Value
		
		If tEnterStatus  = 0 Then
			userId = rstEnter.Fields("L_UID").Value
			userReg = rstEnter.Fields("C_Unique").Value
			userSSB = rstEnter.Fields("C_Name").Value
			adminLevel = rstEnter.Fields("C_Office").Value
			tEnterValid = 1				
		Else
			invalidCounter = invalidCounter + 1
		End If
		
		rstEnter.MoveNext			
		
		If tEnterValid = 1 Then
			Exit Do
		End If
	Loop	

	If securityCounter > 99 Then
		TraceMsg "Saiu do loop pelo security counter" & Err.Number & " " & Err.Descriptionn
	End If

	Set tagIdentified = HMIRuntime.Tags("DB_OPERADOR_POSTO[" & index & "]_POSICAO[" & position & "]_IDENTIFICADO")
	Set tagLogged = HMIRuntime.Tags("DB_OPERADOR_POSTO[" & index & "]_POSICAO[" & position & "]_HABILITADO")
	
	If tEnterValid = 1 Then		
	
		'Check login method (1 = Operator in Workstation, 2 = Admin, 3 = To Update, 4 = Leader)
		Set tagAccessLevel = HMIRuntime.Tags("DB_OPERADOR_POSTO[" & CStr(index) & "]_CMD_W_AccessLevel")
		tagAccessLevel.Read()
		
		Set popid = HMIRuntime.Tags("POP_POSTO[" & index & "]")
		popid.Read()

		If  tagAccessLevel.Value = 1 Then	 'Caso o operador esteja na linha de montagem
							
			tagIdentified.Value = 0
			tagIdentified.Write()
			tagLogged.Value = 0 
			tagLogged.Write()
	
			Set tagUserReg = HMIRuntime.Tags("DB_OPERADOR_POSTO[" & index & "]_POSICAO[" & position & "]_REGISTRO")
			tagUserReg.Value = 0
			tagUserReg.Write()
				
			Set tagUserSSB = HMIRuntime.Tags("DB_OPERADOR_POSTO[" & index & "]_POSICAO[" & position & "]_SSB")
			tagUserSSB.Value = "NONE"
			tagUserSSB.Write()

			Set tagTrainingLevel = HMIRuntime.Tags("DB_OPERADOR_POSTO[" & index & "]_POSICAO[" & position & "]_TRAINING_LEVEL")
			tagTrainingLevel.Value = 0
			tagTrainingLevel.Write()			
					
			'Verifica treinamento da pessoa na posicao
			Set rstTraining = conn.Execute( "SELECT * FROM TB_PESSOAS P ,TB_WORKPLACE_TR WTR WHERE WTR.ID_PESSOA = P.ID AND SSB = '" & userSSB & "' AND ID_WORKPLACE = '" & idWorkstation & "' ")
	
			'Error routine
			If Err.Number <> 0 Then
				TraceMsg "Error #" & Err.Number & " " & Err.Description
				Err.Clear	
				'Close record set
				rstWorkplace.Close	
				rstTerminal.Close			
				'Close data source 
				conn.close			
				Set conn = Nothing
				Set rstTraining = Nothing
				Set rstWorkplace = Nothing
				Set rstTerminal = Nothing 	
				Set rstEnter = Nothing
				biometriaTerminal = -1 ' Not Ok	
				Exit Function
			End If	
						
			trainLevel = rstTraining.Fields("TRAIN_LEVEL").Value 'Seta nivel de treinamento

			If trainLevel = "" Or trainLevel = 0  Then 'Colaborador n�o treinado / identificado /a matriz
				tagIdentified.Value = 0
				tagIdentified.Write()
				tagLogged.Value = 0 
				tagLogged.Write()

				trigger = -3 
				tagTrigger.Value = -3
				tagTrigger.Write()
				
				conn.Execute "EXEC LTS.[dbo].INS_MATRIZ_EVENT " & idWorkstation & "," & userId & "," & 36 & "," & popid.Value
			

			Elseif trainLevel = 1 Then  'Colaborador n�o possui n�vel suficiente para logar						
			
				tagIdentified.Value = 1
				tagIdentified.Write()
				tagLogged.Value = 0 
				tagLogged.Write()

				trigger = -4 
				tagTrigger.Value = -4
				tagTrigger.Write()
				conn.Execute "EXEC LTS.[dbo].INS_MATRIZ_EVENT " & idWorkstation & "," & userId & "," & 35 & "," & popid.Value
			
			Elseif trainLevel >= 2 Then 'Colaborador possui n�vel para logar
			
				'Set popid = HMIRuntime.Tags("POP_POSTO[" & index & "]") -- modificado 19-12-2016
				'popid.Read()
				
				tagIdentified.Value = 1
				tagIdentified.Write()
				tagLogged.Value = 1 
				tagLogged.Write()

				Set tagUserReg = HMIRuntime.Tags("DB_OPERADOR_POSTO[" & index & "]_POSICAO[" & position & "]_REGISTRO")
				tagUserReg.Value = CInt(userReg)
				tagUserReg.Write()
				
				Set tagUserSSB = HMIRuntime.Tags("DB_OPERADOR_POSTO[" & index & "]_POSICAO[" & position & "]_SSB")
				tagUserSSB.Value = userSSB
				tagUserSSB.Write()
				
				Set tagIdWorkstation = HMIRuntime.Tags("DB_OPERADOR_POSTO[" & index & "]_POSICAO[" & position & "]_ID_WORKSTATION")
				tagIdWorkstation.Value = idWorkstation
				tagIdWorkstation.Write()

				Set tagTrainingLevel = HMIRuntime.Tags("DB_OPERADOR_POSTO[" & index & "]_POSICAO[" & position & "]_TRAINING_LEVEL")
				tagTrainingLevel.Value = trainLevel
				tagTrainingLevel.Write()

				trigger = 3
				tagTrigger.Value = 3 ' Ok
				tagTrigger.Write()	
				
				Set tagUserSSB = Nothing
				Set tagIdWorkstation = Nothing
				Set tagTrainingLevel = Nothing
									
				'CRIAR INSERT NA TABELA EVENTOS (SUCESSO)
				conn.Execute "EXEC LTS.[dbo].INS_MATRIZ_EVENT " & idWorkstation & "," & userId & "," & 33 & "," & popid.Value
				
				
				'UPDATE TABELA PESSOA NO POSTO
				conn.Execute "EXEC UPT_PESSOA_POS " & userId  & "," & idWorkstation
				
			End If   	
		
		Elseif tagAccessLevel.Value = 2 Then 'Colaborador logado como ADM
		
			'Seta Tags para DB de Controle do terminal
			Set tagAdminAcessLevel = HMIRuntime.Tags("TERM[" & index  & "]_NIVEL_ACESSO")
			tagAdminAcessLevel.Value = adminLevel
			tagAdminAcessLevel.Write()	
			
			Set tagAdminIdentified = HMIRuntime.Tags("TERM[" & index  & "]_IDENTIFICADO")
			tagAdminIdentified.Value = 1
			tagAdminIdentified.Write()
			
			Set tagAdminSSB = HMIRuntime.Tags("TERM[" & index  & "]_SSB")
			Set tagAdminReg = HMIRuntime.Tags("TERM[" & index  & "]_REGISTRO")
			Set tagAdminLogged = HMIRuntime.Tags("TERM[" & index  & "]_HABILITADO")
			Set tagAdminAcessLevel = HMIRuntime.Tags("TERM[" & index  & "]_NIVEL_ACESSO")
		
				If adminLevel <= 1 Then 'Se nao for, apaga e cai fora
					userId = 0 'zera para garantir registro no banco
									
					tagAdminLogged.Value = 0
					tagAdminLogged.Write()	
					
					tagAdminSSB.Value = "NONE"
					tagAdminSSB.Write()					

					tagAdminReg.Value = 0
					tagAdminReg.Write()														
					
					trigger = -5
					tagTrigger.Value = -5
					tagTrigger.Write()		
					'Insert Nivel de Login invalido	
					conn.Execute "EXEC LTS.[dbo].INS_MATRIZ_EVENT " & idWorkstation & "," & userId & "," & 37 & "," & popid.Value
				Else					
				'Se for loga usuario no posto, IHM fara o controle do tempo e LogOff					
					tagAdminSSB.Value = userSSB
					tagAdminSSB.Write()								

					tagAdminReg.Value = CInt(userReg)
					tagAdminReg.Write()
					
					tagAdminLogged.Value = 1
					tagAdminLogged.Write()					

					tagAdminIdentified.Value = 1
					tagAdminIdentified.Write()																																		
					
					trigger = 4 'Admin Logado
					tagTrigger.Value = 4
					tagTrigger.Write()	
					
					'Insert(SQL) SSB logado no posto como admin	
					conn.Execute "EXEC LTS.[dbo].INS_MATRIZ_EVENT " & idWorkstation & "," & userId & "," & 31 & "," & popid.Value
			
				End If
			
			Set tagAdminSSB = Nothing
			Set tagAdminReg = Nothing
			Set tagAdminLogged = Nothing
			Set	tagAdminIdentified = Nothing
			Set tagAdminAcessLevel = Nothing
		
		Elseif tagAccessLevel.Value = 3 Then 'Colaborador logou para mudar de nivel			

			Set tagUserRegUpdate = "TAG PARA CRIAR DEPOIS"
			tagUserRegUpdate.Value = 0
			tagUserRegUpdate.Write()

			Set tagUserSSBUpdate = "TAG PARA CRIAR DEPOIS"
			tagUserRegUpdate.Value = "NONE"
			tagUserRegUpdate.Write()

			Set tagUserLevelUpdate = "TAG PARA CRIAR DEPOIS"
			tagUserRegUpdate.Value = "NONE"
			tagUserRegUpdate.Write()

			Set tagUserLevelUpdate = "TAG PARA CRIAR DEPOIS"
			tagUserRegUpdate.Value = "NONE"
			tagUserRegUpdate.Write()

			Set rstTraining = conn.Execute( "SELECT * FROM TB_PESSOAS P ,TB_WORKPLACE_TR WTR WHERE WTR.ID_PESSOA = P.ID AND SSB = '" & userSSB & "' AND ID_WORKPLACE = '" & idWorkstation & "' ")
			trainLevel = rstTraining.Fields("TRAIN_LEVEL").Value

		End if

		conn.Execute("DELETE FROM tEnter WHERE L_TID=" & CStr(idTerminal))			
				
		timeCounterTag.Value = 0
		timeCounterTag.Write()
				
		invalidCounterTag.Value = 0
		invalidCounterTag.Write()
		
		rstTraining.Close
		Set rstTraining = Nothing 			
				
	Else 'Ainda nao identificado
	
		timeCounterTag.Value = timeCounterTag.Value + 1
		timeCounterTag.Write()
		
		invalidCounterTag.Value = invalidCounter
		invalidCounterTag.Write()	
		
		If timeCounterTag.Value >= 30 Then 'Expirou tempo
			tagTrigger.Value = -1 ' Timeout
			tagTrigger.Write()	
			'Insert na tabela eventos
		End If		 
		
		If invalidCounterTag.Value > 3 Then
			tagTrigger.Value = -2 ' Tentativas Invalidas
			tagTrigger.Write()	
			'Insert na tabela eventos
		End If
	
	End If			
	
	rstEnter.Close
	Set rstEnter = Nothing		
	
End If	'Trigger 2

If trigger = 1 Then

	conn.Execute("DELETE FROM tEnter WHERE L_TID=" & CStr(idTerminal))
	
	'Error routine - Fehler Routine
	If Err.Number <> 0 Then
		TraceMsg "Error #" & Err.Number & " " & Err.Description
		Err.Clear

		'Close record set
		rstWorkplace.Close	
		rstTerminal.Close			
		'Close data source - Datenquelle schlie?en
		conn.close			
		Set conn = Nothing
		Set rstWorkplace = Nothing
		Set rstTerminal = Nothing 	
		trigger = -6
		tagTrigger.Value = -6 ' Waiting 	
		tagTrigger.Write()
		
		biometriaTerminal = -1 ' Not Ok	
		Exit Function
	End If		
	
	timeCounterTag.Value = 0
	timeCounterTag.Write()
	
	invalidCounterTag.Value = 0
	invalidCounterTag.Write()
	
	
	trigger = 2
	tagTrigger.Value = 2 ' Waiting 	
	tagTrigger.Write()			


End If	'Trigger  1

If trigger = 10 Then 

	Set popid = HMIRuntime.Tags("POP_POSTO[" & index & "]")
	popid.Read()

	timeCounterTag.Value = 0
	timeCounterTag.Write()
	
	invalidCounterTag.Value = 0
	invalidCounterTag.Write()
	
	Set tagIdentified = HMIRuntime.Tags("DB_OPERADOR_POSTO[" & index & "]_POSICAO[" & position &  "]_IDENTIFICADO")
	Set tagLogged = HMIRuntime.Tags("DB_OPERADOR_POSTO[" & index & "]_POSICAO[" & position &  "]_HABILITADO")
	
	tagIdentified.Value = 0
	tagLogged.Value = 0
	tagIdentified.Write()
	tagLogged.Write()
	
	Set tagUserReg = HMIRuntime.Tags("DB_OPERADOR_POSTO[" & index & "]_POSICAO[" & position &  "]_REGISTRO")
	tagUserReg.Value = 0
	tagUserReg.Write()
				
	Set tagUserSSB = HMIRuntime.Tags("DB_OPERADOR_POSTO[" & index & "]_POSICAO[" & position &  "]_SSB")
	tagUserSSB.Value = "NONE"
	tagUserSSB.Write()

	Set tagTrainingLevel = HMIRuntime.Tags("DB_OPERADOR_POSTO[" & index & "]_POSICAO[" & position &  "]_TRAINING_LEVEL")
	tagTrainingLevel.Value = 0
	tagTrainingLevel.Write()
	
	'INSERT EVENTO LOGOFF
	conn.Execute "EXEC LTS.[dbo].INS_MATRIZ_EVENT " & idWorkstation & "," & userId & "," & 34 & "," & popid.Value
	
	'REMOVE OPERADOR DO POSTO
	conn.Execute "DELETE FROM TB_PESSOA_POSTO WHERE ID_WORKPLACE = " & idWorkstation
	
	trigger = 0
	tagTrigger.Value = 0 ' Waiting 	
	tagTrigger.Write()	

End If 'Trigger 10


'Close the recordset
rstTerminal.Close
rstWorkplace.Close		

'Close data source 
conn.close

Set rstTerminal = Nothing		
Set rstWorkplace = Nothing		
Set conn = Nothing
Set tagTrigger = Nothing	
Set popid = Nothing
