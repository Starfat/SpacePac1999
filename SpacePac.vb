
Option Explicit
Dim intKey As Byte 'Variabel for å motta hvilken piltast som er holdt nede
Dim intLife As Byte 

Private Sub Form KeyDown (KeyCode As Integer, Shift As Integer)
Dim i As Byte 'Løkketeller
	For i = 0 To 2 'Starter Iøkka
		If imgPac. Top = IinHor (i) .YI Then
			'Undersøker om PacMan befinner seg i en av de horisontale banene
			If KeyCode 37 Then
			'Undersøker om venstre piltast er holdt nede
				intKey = KeyCode
				'Overfører KeyCodeverdien til variabelen intKey hvis ovenstående argumenter er sanne
			End If
		End If
	Next i
	For i 0 To 2 'Starter Iøkka
		If imgPac.Top = IinHor (i).Yl Then
		'Undersøker om PacMan befinner seg i en av de horisontale banene
			If KeyCode — 39 Then
			'Undersøker om høyre piltast er holdt nede
				intKey = KeyCode
				'Overfører KeyCodeverdien til variabelen intKey hvis ovenstående argumenter er sanne
			End If
		End If
	Next i
	For i = 0 To 2 'Starter Iøkka
		If imgPac.Left = IinVer (i).Xl Then
		'Undersøker om PacMan befinner seg i en av de vertikale banene
			If KeyCode — 40 Then
			'Undersøker om nedre pi1tast er holdt nede
				intKey = KeyCode
				'Overfører KeyCodeverdientil variabelenintKey hvis ovenstående argumenter er sanne
			End If
		End If
	Next i
	For i = 0 To 2 'Starter Iøkka
		If imgPac.Left = IinVer (i).Xl Then
		'Undersøker om PacMan befinner seg i en av de vertikale banene
			If KeyCode — 38 Then
			'Undersøker om nedre pi1tast er holdt nede
				intKey = KeyCode
				'Overfører KeyCodeverdientil variabelenintKey hvis ovenstående argumenter er sanne
			End If
		End If
	Next i
End Sub

Private Sub Form_Load()
	intLife = 2
End Sub

Private Sub mnuAbout_Click ( )
	frmAbout . Show 'Viser Aboutmenyen
End Sub

Private Sub mnuExit Click ()
	Dim intSvar As Integer 'deklarering
	intSvar = MsqBox(prompt:="Are you sure you want to quit?",
				Buttons:=vbYesNo + vbQuestion,
				Title:="Quit?")
				'Viser MsgBox med spørsmål om du vil avslutte
	If intSvar = vbYes Then
		End
		'Avslutter hvis svaret er ja
	Else
		Unload frmGratuIerer
		Unload frmBoardl
		frmBoard1.Show
		'Fjerner Gratulererformen oq starter spillet på nytt hvis svaret er nei
	End If
End Sub
			
Private Sub mnuNew Click()
	frmNewGame . Show 'Viser NewGame-menyen
End Sub

Private Sub Click()
	frmStatus.Show 'Viser Statusmenyen
End Sub

Private Sub tmr Timer ()
Dim i As Byte 'Deklarerinq av teller

	If intLife 0 Then
		End
	End If
	
'Rutine som styrer Ghosts beveqelse nedover og oppover
	If tmrGhostEscape.EnabIed = False Then
		If imgPac.Left — imqGhost.Left
			And (imgPac.Left + ImgPac.Width) = (imqGhost.Left + imqGhost.width))
			And imqPac.Top = imgGhost.Top
			And (imqPac.Top + imqPac.Height) = imgGhost.Top + imgGhost.Height Then
			'Undersøker om ghost befinner seq innenfor PacMans fire sider
			intLife = intLife - 1
			Call PacStartPosition
		End if
		For i = 0 To 2 'Starter teller
			If imgGhost. Left — IinVer(i).X1 And imgGhost.Top < imgPac.Top And imgGhost.Top < IinHor(2).Y1 Then
			'Undersøker om Ghost befinner seg på en av de vertikale sidene og høyere enn PacMan 
				imgGhost. Top = imgGhost. Top + 200
				'Flytter ghost nedover hvis ovenstående argument er sant
			Else lf imgGhost.Left — linVer(i).X1 And imgGhost.Top > imgPac.Top And imgGhost.Top > linHor(O).Y1 Then
			'Undersøker om Ghost befinner seg på en av de vertikale sidene og  lavere enn PacMan 
				imgGhost . Top = imgGhost . Top - 200
				'Flytter ghost oppover hvis ovenstående argument er sant
			End If
	Next i
 
'Rutine som styrer Ghosts bevegelse mot høyre og venstre
	For i = 0 To 2 'Starter teller
		If imgGhost.Top = IinHor(i).Y1 And imgGhost.Left < imgPac.Left And imgGhost.Left < lin Ver(2).X1 Then
		' Undersøker om Ghost befinner seg på en av de horisontale sidene og til
		' venstre for PacMan 
			imgGhost.Left = imgGhost.Left + 200
			'Flytter ghost mot høyre hvis ovenstående argument er sant
		Else If imgGhost.Top — linHor(i).Y1 And   imgGhost.Left > imgPac.Left And imgGhost.Left > linVer(0).X1 Then
		'Undersøker om Ghost befinner seg på en av de horisontale sidene og til høyre for PacMan 
			imgGhost.Left = imgGhost.Left - 200
			'Flytter ghost mot venstre hvis ovenstående argument er sant
		End If
	Next i
	
'Rutine som styrer PacMans bevegelse mot venstre
	If imgPac.Left > IinVer(0).XI Then
	'Undersøker om PacMan befinner seg til høyre for venstre vertikale bane 
		If intKey = 37 Then
		'Undersøker om venstre piltast er holdt nede 
			imgPac.Left = imgPac.Left - 200
			'Beveger PacMan mot venstre hvis ovenstående argumenter er sanne
			Call EatDot
			'Kaller opp prosedyren som undersøker om en Dot har blitt spist og som leverer poeng tilbake til lblPoints
			Call EatBonusdot
			'Kaller opp prosedyren som undersøker om en Bonusdot har blitt spist og som leverer poeng tilbake til IblPoints
		End If		
	End If
	
'Rutine som styrer PacMans bevegelse mot høyre
	If imgPac.Left < IinVer(2).XI Then
	'Undersøker om PacMan befinner seg til høyre for venstre vertikale bane 
		If intKey = 39 Then
		'Undersøker om venstre piltast er holdt nede 
			imgPac.Left = imgPac.Left + 200
			'Beveger PacMan mot høyre hvis ovenstående argumenter er sanne
			Call EatDot
			'Kaller opp prosedyren som undersøker om en Dot har blitt spist og som leverer poeng tilbake til lblPoints
			Call EatBonusdot
			'Kaller opp prosedyren som undersøker om en Bonusdot har blitt spist og som leverer poeng tilbake til IblPoints
		End If		
	End If
	
'Rutine som styrer PacMans bevegelse oppover
	If imgPac.Left > IinHor(0).Y1 Then
	'Undersøker om PacMan befinner seg under øvre horisontale bane 
		If intKey = 38 Then
		'Undersøker om øvre piltast er holdt nede 
			imgPac.Top = imgPac.top - 200
			'Beveger PacMan mot oppover hvis ovenstående argumenter er sanne
			Call EatDot
			'Kaller opp prosedyren som undersøker om en Dot har blitt spist og som leverer poeng tilbake til lblPoints
			Call EatBonusdot
			'Kaller opp prosedyren som undersøker om en Bonusdot har blitt spist og som leverer poeng tilbake til IblPoints
		End If		
	End If
	
'Rutine som styrer PacMans bevegelse nedover
	If imgPac.Top < IinHor(2).Y1 Then
	'Undersøker om PacMan befinner seg under øvre horisontale bane 
		If intKey = 40 Then
		'Undersøker om nedre piltast er holdt nede 
			imgPac.Top = imgPac.top + 200
			'Beveger PacMan mot oppover hvis ovenstående argumenter er sanne
			Call EatDot
			'Kaller opp prosedyren som undersøker om en Dot har blitt spist og som leverer poeng tilbake til lblPoints
			Call EatBonusdot
			'Kaller opp prosedyren som undersøker om en Bonusdot har blitt spist og som leverer poeng tilbake til IblPoints
		End If		
	End If
	
' Rutine som styrer PacMans bevegelse fra overkant til underkant
	If imgPac. Top = IinVer(1).Y2 And imgPac.Left = IinVer(1).XI Then  
	'Undersøker om PacMan befinner seg i øvre horisontale bane 
		If intKey = 38 Then
		'Undersøker om øvre piltast er holdt nede 
			imgPac.Top = IinVer(1).Yl
			'Lar PacMan passere fra overkant til underkant hvis ovenstående  argumenter er sanne
		End If
	End If
	
End Sub

Private Sub tmrGhostEscape_Timer()
Static bteTe11er As Byte
Dim i As Integer

 'Rutine som styrer Ghosts flukt fra PacMan
	If imgPac.Left = imgGhost.Left  
	And (imgPac.Left + imgPac.Width) = (imgGhost.Left + imgGhost.Width)  
	And imgPac.Top = imgGhost.Top
	And (imgPac.Top + imgPac.Height) = imgGhost.Top + imgGhost.Height Then
		intPoints = intPoints + 20 
		Call GhostStartPosition 
	End If
	
	For i — 0 To 2 'Starter teller
		If imgGhost.Left = IinVer(i).XI And imgGhost.Top < imgPac.Top And imgGhost.Top > linHor(0).Yl Then
		'Undersøker om Ghost befinner seg på en av de vertikale sidene og høyere enn PacMan 
			imgGhost.Top = imgGhost.Top - 200
			' Flytter ghost nedover hvis ovenstående argument er sant
		Else if imgGhost.Left = linVer(i).XI And imgGhost.Top > imgPac.Top  
		And imgGhost.Top < linHor(2).Yl Then
		'Undersøker om Ghost befinner seg på en av de vertikale sidene og  lavere enn PacMan 
			imgGhost.Top = imgGhost.Top + 200
			' Flytter ghost oppover hvis ovenstående argument er sant
		End If
	Next i

	For i — 0 To 2 'Starter teller
		If imgGhost.Top = linHor (i).Yl And imgGhost.Left < imgPac.Left  
		And imgGhost.Left > IinVer.XI Then
		'Undersøker om Ghost befinner seg på en av de horisontale sidene og til venstre for PacMan 
		imgGhost.Left = imgGhost.Left - 200
		'Flytter ghost mot høyre hvis ovenstående argument er sant
		Else If imgGhost.Top = linHor(i).Yl And imgGhost.Left > imgPac.Left  
		And imgGhost.Left < linVer(2).XI Then
		'Undersøker om Ghost befinner seg på på en av de horisontale sidene og til høyre for PacMan 
			imgGhost.Left = imgGhost.Left + 200
			'Flytter ghost mot venstre hvis ovenstående argument er sant
		End If 
	Next i

	For i = O To 2
		If imgPac.Left = imgGhost.Left And imgGhost.Left < linVer(2).X1  
		And imgGhost.Top = linHor(i).Yl Then 
			imgGhost.Left = imgGhost.Left + 200
		End If 
	Next i
	
	For i — 0 To 2
		If imgPac.Left = imgGhost.Left And imgGhost.Left > linVer(0).Xl  
		And imgGhost.Top = linHor(i).Yl Then 
			imgGhost.Left = imgGhost.Left - 200
		End If
	Next i

	For i = 0 To 2
		If imgPac.Top = imgGhost.Top And imgGhost.Top < linHor(2).Yl
		And imgGhost.Left = linVer(i).Xl Then imgGhost.Top = imgGhost.Top + 200
		End If 
	Next i

	For i = O To 2
		If imgPac.Top = imgGhost.Top And imgGhost.Top > linHor(0).YI
		And imgGhost.Left = IinVer(i).Xl Then 
		imgGhost.Top = imgGhost.Top — 200
		End If 
	Next i
	
bteTeller = bteTe11er + 1

	If bteTeIIer = 35 Then 
		tmrGhostEscape.Enabled = False 
		bteTe11er =  0
	End If
End Sub
