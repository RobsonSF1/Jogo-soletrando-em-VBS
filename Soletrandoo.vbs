Dim palavras_sorteadas(25), i, n
Dim acertos, resposta,resp
Dim audio, message, titulo
dim pular_palavra



message = "Digite a palavra que acabou de ouvir"
titulo = "JOGO SOLETRANDO"

Call carregar_voz

Sub carregar_voz()
    Set audio = CreateObject("SAPI.SpVoice")
    audio.Volume = 100
    audio.Rate = 3
End Sub

Call iniciar

Sub iniciar()
   pular_palavra=1
     
        'audio.Speak "Bem vindo ao Soletrar. Prepare-se para digitar a palavra que ouvirá logo a seguir!"
        resp = MsgBox("Deseja ouvir a palavra?", vbInformation + vbYesNo, "AVISO")
        If resp = vbYes Then
            Call palavras_array
        Else
            WScript.Quit
        End If
end sub

Sub sortear()
    acertos = 0
    For i = 1 To 25 Step 1 
        Do
            n = Int(Rnd * 25) + 1 ' Gera um número aleatório entre 1 e 25
        Loop While palavras_sorteadas(n) = "" ' Verifica se a palavra na posição (n) ainda não foi sorteada

        audio.Speak (palavras_sorteadas(n))
       resposta = UCase(InputBox(message+vbnewline&_
					    "[P]PARA PULAR", acertos))

        If resposta = "P" Then
            If pular_palavra > 0 Then
                pular_palavra = 0
                Call sortear
            Else
                MsgBox("Função já utilizada!!"), vbExclamation + vbOKOnly, "AVISO"
                resposta = UCase(InputBox(message+vbnewline&_
				 "[P]PARA PULAR", titulo))
            End If
        End If

        If resposta = UCase(palavras_sorteadas(n)) Then
            MsgBox("Parabéns, PALAVRA CORRETA!!")
            acertos = acertos + 1
        if acertos = 15 Then
			msgbox("PARABÉNS, VOCÊ GANHOU!!VOCÊ ACERTOU:"& acertos),vbexclamation+vbokonly,"FIM DE JOGO"
			WScript.Quit
		end if
		else 
            MsgBox("PALAVRA INCORRETA! Fim do jogo"), vbInformation + vbOKOnly, "Seus Acertos: " & acertos
            WScript.Quit
        End If
    Next
End Sub
			
			

Sub palavras_array()
    palavras_sorteadas(1) = "casa"
    palavras_sorteadas(2) = "tesouro"
    palavras_sorteadas(3) = "cavalo"
    palavras_sorteadas(4) = "cachorro"
    palavras_sorteadas(5) = "carro"
    palavras_sorteadas(6) = "avião"
    palavras_sorteadas(7) = "chapéu"
    palavras_sorteadas(8) = "escada"
    palavras_sorteadas(9) = "óculos"
    palavras_sorteadas(10) = "cadeira"
    palavras_sorteadas(11) = "almofada"
    palavras_sorteadas(12) = "esponja"
    palavras_sorteadas(13) = "dado"
    palavras_sorteadas(14) = "torneira"
    palavras_sorteadas(15) = "faca"
    palavras_sorteadas(16) = "copo"
    palavras_sorteadas(17) = "balde"
    palavras_sorteadas(18) = "pincel"
    palavras_sorteadas(19) = "bolo"
    palavras_sorteadas(20) = "sapo"
    palavras_sorteadas(21) = "cobra"
    palavras_sorteadas(22) = "iguana"
    palavras_sorteadas(23) = "macaco"
    palavras_sorteadas(24) = "lâmpada"
    palavras_sorteadas(25) = "ferrari"
    
    Randomize 
    n=int(rnd * 25) + 1 ' Gera um numero aleatório entre 1 e 25
	call sortear
End Sub