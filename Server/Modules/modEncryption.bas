Attribute VB_Name = "modEncryption"
'********************************************
'*  Name:  modEncryption                    *
'*  Function:  Encrypt any message with     *
'*             Random numbers generated at  *
'*             Encryption time and then en- *
'*             crypt the number used and    *
'*             store it with the message.   *
'*  Affect:  Will never have the same En-   *
'*           crypted message twice          *
'*  Author:  James Miller                   *
'*  E-mail:  usssssy@yahoo.com              *
'********************************************

Public Function Encrypt(sString As String) As String
    Dim NewString As String
    Dim RandNum As Long
    RandNum = Rnd() * 40
    Encrypt = RandNum & Chr(3)
    For i = 1 To Len(sString)
        Encrypt = Encrypt & Encrypt_Letter(Mid(sString, i, 1), RandNum)
    Next i
End Function
Private Function Encrypt_Letter(lLetter As String, RandNum As Long) As String
    Dim RN As Long
    RN = Rnd() * 40
    If RN <= 2 Then RN = 3
    RN = Format(RN, "###")
    Dim RNS As String
    RNS = RN
    Encrypt_Letter = Chr(Asc(lLetter) + RN) & Chr(1)
    Encrypt_Letter = Encrypt_Letter & Encrypt_Number(RNS, RandNum)
End Function
Private Function Encrypt_Number(nNumber As String, RandNum As Long) As String
    Dim NewNum As String
    For x = 1 To Len(nNumber)
        NewNum = NewNum & Chr(Asc(Mid(nNumber, x, 1)) + RandNum)
    Next x
    Encrypt_Number = NewNum & Chr(2)
End Function
Public Function Decrypt(sString As String, cCall As String) As String
    Dim RandNum As Long
    Dim sStop As Long
    For i = 1 To Len(sString)
        If Mid(sString, i, 1) = Chr(3) Then
            sStop = i
        End If
    Next i
    RandNum = Mid(sString, 1, sStop - 1)
    sString = Mid(sString, sStop + 1, Len(sString))
    Decrypt = Decrypt_Numbers(sString, RandNum)
End Function
Private Function Decrypt_Numbers(sString As String, RndNum As Long) As String
    Dim TempString As String
    Dim TempNumber As String
    Dim lLast As Long
    Dim sStart As Long
    Dim sStop As Long
    lLast = 1
    For i = 1 To Len(sString)
        If Mid(sString, i, 1) = Chr(1) Then
            sStart = i
        End If
        If Mid(sString, i, 1) = Chr(2) Then
            sStop = i
            TempString = TempString & Mid(sString, lLast, 1)
            TempString = TempString & Chr(1) & Get_Number(Mid(sString, sStart + 1, (sStop - 1) - sStart), RndNum) & Chr(2)
            lLast = sStop + 1
        End If
    Next i
    Decrypt_Numbers = TempString
    Decrypt_Numbers = Decrypt_Letters(Decrypt_Numbers)
End Function
Private Function Decrypt_Letters(DecNum As String) As String
    Dim tString As String
    Dim tNum As Long
    Dim S1 As Long
    Dim S2 As Long
    Dim lNum As Long
    lNum = 1
    For i = 1 To Len(DecNum)
        If Mid(DecNum, i, 1) = Chr(1) Then
            S1 = i
        End If
        If Mid(DecNum, i, 1) = Chr(2) Then
            S2 = i
            Dim eEnd As Long
            eEnd = (S2 - 1) - S1
            tNum = Mid(DecNum, S1 + 1, eEnd)
            tString = Mid(DecNum, lNum, 1)
            Decrypt_Letters = Decrypt_Letters & Get_Letter(tString, tNum)
            lNum = S2 + 1
        End If
    Next i
End Function
Private Function Get_Letter(lLetter As String, RN As Long) As String
    Get_Letter = Chr(Asc(lLetter) - RN)
End Function
Private Function Get_Number(sString As String, RadNum As Long) As String
    Dim TempStrng As String
    For i = 1 To Len(sString)
        TempStrng = TempStrng & Chr(Asc(Mid(sString, i, 1)) - RadNum)
    Next i
    Get_Number = TempStrng
End Function
