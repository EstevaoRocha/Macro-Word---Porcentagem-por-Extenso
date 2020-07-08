'*****************Código original de: Eduardo (eduzsrj) 07/02/2011*****************'
'https://www.clubedohardware.com.br/topic/817427-dica-ms-office-valor-por-extenso-no-word/'
''
'*****************Código adaptado: Luan (luangabs) 10/03/2017*****************'
'https://www.clubedohardware.com.br/topic/1219410-dica-ms-office-valor-percentual-por-extenso-no-word/'
''
'*****************Código adaptado: Estevão (estevaogsr) 08/07/2020*****************'
''
'*****************Notas de nova versão:*****************'
'Escreve por extenso até trés casas decimais(milésimos)'
'Só escreve o sufixo inteiro quando é um numero decimal'
'Distingue décimos de centésimos'

Sub WValorExtensoPerc()

'    Selection.MoveLeft unit:=wdWord, Count:=1, Extend:=wdExtend

    On Error GoTo Erro

    Selection.MoveStartUntil cset:=" ", Count:=wdBackward
    Selection.TypeText Selection.Text & "%" & " (" & ConverterParaExtensoPerc(Selection.Text) & " por cento" & ")"

    GoTo Pula

Erro:

    MsgBox "O valor deve ser informado sem ponto e sem 'R$'." & Chr$(10) & "O cursor deve estar imediatamente após o valor." _
    & Chr$(10) & "O valor não pode estar em início de parágrafo." & Chr$(10) & _
    "Exemplo: 1250,35", vbCritical, "Dados inválidos!"

    Exit Sub

Pula:

End Sub

Public Function ConverterParaExtensoPerc(NumeroParaConverter As String) As String
Dim sExtensoFinal As String, sExtensoAtual As String
Dim i As Integer
Dim iQtdGrupos As Integer
Dim Original As String
Dim sDecimais As String
Dim sTipoSing As String, sTipoPlu As String, sDecimos As String, sConector As String
Dim bSufTipo As Boolean
Dim vArrCenten As Variant

Original = NumeroParaConverter

'Separa os Decimais
If InStr(1, NumeroParaConverter, ",") > 0 Then
sDecimais = Right(NumeroParaConverter, Len(NumeroParaConverter) - InStr(1, NumeroParaConverter, ","))
NumeroParaConverter = Mid(NumeroParaConverter, 1, InStr(1, NumeroParaConverter, ",") - 1)
End If

'Obtém a separação de milhares
iQtdGrupos = Fix(Len(NumeroParaConverter) / 3)
If Len(NumeroParaConverter) Mod 3 > 0 Then
iQtdGrupos = iQtdGrupos + 1
End If

'Chama as funções para escrever o número
If iQtdGrupos > 2 Then bSufTipo = True

For i = iQtdGrupos To 1 Step -1
sExtensoAtual = DesmembraValor(NumeroParaConverter, i)
If i = 1 Then
If sExtensoAtual = "" Then
sExtensoFinal = sExtensoFinal & sExtensoAtual
Else
If sExtensoFinal = "" Then
sExtensoFinal = sExtensoFinal & sExtensoAtual
Else

vArrCenten = Array("cem", "duzentos", "trezentos", "quatrocentos", _
"quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos")

sConector = ""

For w = 0 To 8
If Len(NumeroParaConverter) >= 4 And Right(NumeroParaConverter, 2) = "00" _
And sExtensoAtual <> vArrCenten(w) Then sConector = "e "
Exit For
Next w

If Len(NumeroParaConverter) >= 4 And Left(Right(NumeroParaConverter, 3), 1) = "0" Then sConector = " e "

If Len(NumeroParaConverter) >= 4 And sExtensoAtual = "cem" Then sConector = " e "

sExtensoFinal = sExtensoFinal & sConector & sExtensoAtual
End If
End If
Else
sExtensoFinal = sExtensoFinal & sExtensoAtual
End If

If iQtdGrupos > 2 Then
Select Case i
Case 1, 2
If sExtensoAtual <> "" Then
bSufTipo = False
End If
End Select
End If
Next i

'Define o tipo
If Original = Int(Original) Then
    sTipoPlu = ""
    sTipoSing = ""
Else
    sTipoPlu = " inteiros"
    sTipoSing = ""
End If

If bSufTipo = True Then sTipoPlu = " de inteiros"

'Escreve os décimos
sDecimos = EscreveDecimosPerc(sDecimais)

'Adiciona os décimos e se é plural ou singular
sExtensoFinal = IIf((sExtensoFinal = ""), "", sExtensoFinal & IIf((sExtensoFinal = "um"), sTipoSing, sTipoPlu)) _
& IIf((sExtensoFinal = ""), sDecimos, IIf((sDecimos = ""), "", " e " & sDecimos))

'retorna o resultado

sExtensoFinal = Replace(sExtensoFinal, "  ", " ", 1, , vbTextCompare)

ConverterParaExtensoPerc = Replace(sExtensoFinal, " e e ", " e ", 1, , vbTextCompare)

End Function

Private Function DesmembraValor(sValor As String, iGrupoDiv As Integer) As String
Dim iValor As Integer
Dim sExtenso As String
Dim iDivResto As Integer
Dim iDivInteiro As Integer
Dim iPosInicMid As Integer
Dim iTamMid As Integer
Dim sComplemento As String
Dim vArrDez1 As Variant
Dim vArrDez2 As Variant
Dim vArrCentena As Variant

vArrDez1 = Array("", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove", _
"dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", _
"dezoito", "dezenove")

vArrDez2 = Array("vinte", "trinta", "quarenta", "cinquenta", "sessenta", _
"setenta", "oitenta", "noventa")

vArrCentena = Array("cem", "cento", "duzentos", "trezentos", "quatrocentos", _
"quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos")

'Pega o Valor a ser escrito e desmembra para o grupo numérico correto
iPosInicMid = Len(sValor) - ((3 * iGrupoDiv) - 1)
If iPosInicMid <= 1 Then
iTamMid = 2 + iPosInicMid
Else
iTamMid = 3
End If

If iPosInicMid < 1 Then iPosInicMid = 1

iValor = CInt(Mid(sValor, iPosInicMid, iTamMid))

Select Case iGrupoDiv
Case 2
sComplemento = " mil "
Case 3
If iValor = 1 Then
sComplemento = " milhão "
Else
sComplemento = " milhões "
End If
Case 4
If iValor = 1 Then
sComplemento = " bilhão "
Else
sComplemento = " bilhões "
End If
Case 5
If iValor = 1 Then
sComplemento = " trilhão "
Else
sComplemento = " trilhões "
End If
End Select

Select Case iValor
Case 0 To 19
sExtenso = vArrDez1(iValor)
Case 20 To 99
iDivInteiro = Fix(iValor / 10)
iDivResto = iValor Mod 10

If iDivResto = 0 Then
sExtenso = vArrDez2(iDivInteiro - 2)
Else
sExtenso = vArrDez2(iDivInteiro - 2) & " e " & vArrDez1(iDivResto)
End If
Case 100 To 999
iDivInteiro = Fix(iValor / 100)
iDivResto = iValor Mod 100

If iDivResto = 0 Then
If iDivInteiro = 1 Then
sExtenso = vArrCentena(0)   'Cem
Else
sExtenso = vArrCentena(iDivInteiro) 'inteiro maior que 100
End If
Else
sExtenso = vArrCentena(iDivInteiro) & " e "
Select Case iDivResto
Case 0 To 19
sExtenso = sExtenso & vArrDez1(iDivResto)
Case 20 To 99
iDivInteiro2 = Fix(iDivResto / 10)
iDivResto2 = iDivResto Mod 10

If iDivResto2 = 0 Then
sExtenso = sExtenso & vArrDez2(iDivInteiro2 - 2)
Else
sExtenso = sExtenso & vArrDez2(iDivInteiro2 - 2) & " e " & vArrDez1(iDivResto2)
End If
End Select
End If

End Select

If sExtenso = "um" And sComplemento = " mil " And Len(sValor) < 7 Then
sComplemento = "mil "
sExtenso = ""
End If

smilx = Right(sValor, 6)

If sComplemento = " milhão " Then
If Left(smilx, 2) = "00" And Right(smilx, 5) <> "00000" Then sComplemento = " milhão e " Else sComplemento = " milhão "
End If

If sComplemento = " milhões " Then
If Right(smilx, 6) = "000000" Then
sComplemento = " milhões "
Else
If Left(smilx, 2) = "00" And Right(smilx, 5) <> "00000" Then sComplemento = " milhões e " Else sComplemento = " milhões "
End If
End If

DesmembraValor = sExtenso & IIf(iValor > 0, sComplemento, "")

End Function

Private Function EscreveDecimosPerc(sCent As String) As String
Dim sExtenso As String
Dim iDivResto As Integer
Dim iDivResto2 As Integer
Dim iDivInteiro As Integer
Dim iDivInteiro2 As Integer
Dim sComplemento As String
Dim vArrDez1 As Variant
Dim vArrDez2 As Variant
Dim vArrCen As Variant
Dim iCent As Integer
Dim iLenght As Integer

vArrDez1 = Array("", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove", _
"dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", _
"dezoito", "dezenove")

vArrDez2 = Array("vinte", "trinta", "quarenta", "cinquenta", "sessenta", _
"setenta", "oitenta", "noventa")

vArrCen = Array("cem", "cento", "duzentos", "trezentos", "quatrocentos", _
"quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos")

'Separa os Decimais
iLenght = Len(sCent)

If iLenght = 1 Then
    'Adequando para uma casa decimal
    iCent = Fix(sCent & String(1 - Len(sCent), "0"))
    
    'Escrevendo Singular ou plural
    If iCent = 1 Then
    sComplemento = " décimo"
    Else
    sComplemento = " décimos"
    End If
ElseIf iLenght = 2 Then
    'Adequando para duas casas decimais
    iCent = Fix(sCent & String(2 - Len(sCent), "0"))
    
    'Escrevendo Singular ou plural
    If iCent = 1 Then
    sComplemento = " centésimo"
    Else
    sComplemento = " centésimos"
    End If
ElseIf iLenght = 3 Then
    'Adequando para tres casas decimais
    iCent = Fix(sCent & String(3 - Len(sCent), "0"))
    
    'Escrevendo Singular ou plural
    If iCent = 1 Then
    sComplemento = " milésimo"
    Else
    sComplemento = " milésimos"
    End If
End If

'Calculando os valores
Select Case iCent
    Case 0 To 19
        sExtenso = vArrDez1(iCent)
    Case 20 To 99
        iDivInteiro = Fix(iCent / 10)
        iDivResto = iCent Mod 10
        
        If iDivResto = 0 Then
            sExtenso = vArrDez2(iDivInteiro - 2)
        Else
            sExtenso = vArrDez2(iDivInteiro - 2) & " e " & vArrDez1(iDivResto)
        End If
    Case 100 To 999
        iDivInteiro = Fix(iCent / 100)
        iDivResto = iCent Mod 100
        
        If iDivResto = 0 Then
            If iDivResto = 1 Then
            sExtenso = vArrCen(0) 'cem
            Else
            sExtenso = vArrCen(iDivInteiro - 1) 'inteiro maior que 100
            End If
        Else
            sExtenso = vArrCen(iDivInteiro) & " e " 'Concatena a centena com a dezena
            Select Case iDivResto
                Case 0 To 19
                    sExtenso = sExtenso & vArrDez1(iDivResto)
                Case 20 To 99
                    iDivInteiro2 = Fix(iDivResto / 10)
                    iDivResto2 = iDivResto Mod 10
                If iDivResto2 = 0 Then
                    sExtenso = sExtenso & vArrDez2(iDivInteiro2 - 2)
                Else
                    sExtenso = sExtenso & vArrDez2(iDivInteiro2 - 2) & " e " & vArrDez1(iDivResto2)
                End If
            End Select
        End If
End Select

EscreveDecimosPerc = IIf(iCent > 0, sExtenso & sComplemento, "")

End Function