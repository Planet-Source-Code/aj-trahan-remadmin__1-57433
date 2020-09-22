Attribute VB_Name = "modEncryption"
'**********************************************
'**  Program Name:  SubNet® (©2003-2004)     **
'**  Program Description:  Remotely Control  **
'**       A Computer Via LAN/WAN/Dial-Up     **
'**  Program Ports:  6800-6805               **
'**  GUI Designer:  Steven Tanyi             **
'**     E-mail:  sniper6oo@hotmail.com       **
'**  Programmer:  James Miller               **
'**     E-mail:  usssssy@yahoo.com           **
'**********************************************

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
'frmEnrcryption.Text2.Text will now = the "Encrypted," frmEncrypted.Text1.Text
Dim NewString As String
'So we don't have to change our original sString, We make up a new one.
'You also Create them at app load to save time when running, use this time
'   in the "Splash" screen.
Dim RandNum As Long
'creating a variable
RandNum = Rnd() * 40
'making it a random number
Encrypt = RandNum & Chr(3)
For I = 1 To Len(sString)
'Let's go through every letter/variable 1 at a time and not change it
    Encrypt = Encrypt & Encrypt_Letter(Mid(sString, I, 1), RandNum)
Next I
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
For X = 1 To Len(nNumber)
    NewNum = NewNum & Chr(Asc(Mid(nNumber, X, 1)) + RandNum)
Next X
Encrypt_Number = NewNum & Chr(2)
End Function
Public Function Decrypt(sString As String) As String
Dim RandNum As Long
Dim sStop As Long
For I = 1 To Len(sString)
    If Mid(sString, I, 1) = Chr(3) Then
        sStop = I
    End If
Next I
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
For I = 1 To Len(sString)
    If Mid(sString, I, 1) = Chr(1) Then
        sStart = I
    End If
    If Mid(sString, I, 1) = Chr(2) Then
        sStop = I
        TempString = TempString & Mid(sString, lLast, 1)
        TempString = TempString & Chr(1) & Get_Number(Mid(sString, sStart + 1, (sStop - 1) - sStart), RndNum) & Chr(2)
        lLast = sStop + 1
    End If
Next I
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
For I = 1 To Len(DecNum)
    If Mid(DecNum, I, 1) = Chr(1) Then
        S1 = I
    End If
    If Mid(DecNum, I, 1) = Chr(2) Then
        S2 = I
        Dim eEnd As Long
        eEnd = (S2 - 1) - S1
        tNum = Mid(DecNum, S1 + 1, eEnd)
        tString = Mid(DecNum, lNum, 1)
        Decrypt_Letters = Decrypt_Letters & Get_Letter(tString, tNum)
        lNum = S2 + 1
    End If
Next I
End Function
Private Function Get_Letter(lLetter As String, RN As Long) As String
Get_Letter = Chr(Asc(lLetter) - RN)
End Function
Private Function Get_Number(sString As String, RadNum As Long) As String
Dim TempStrng As String
For I = 1 To Len(sString)
    TempStrng = TempStrng & Chr(Asc(Mid(sString, I, 1)) - RadNum)
Next I
Get_Number = TempStrng
End Function


