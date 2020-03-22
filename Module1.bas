Attribute VB_Name = "Module1"
Option Explicit

Function Additionner(a As Integer, b As Integer) As Integer
Dim somme As Integer

    somme = a + b

    Additionner = somme
End Function

Sub FaireDesCalculs()
Dim a As Integer, o As Integer, resultatDuCalcul As Integer

    a = 16
    o = 9
    resultatDuCalcul = Additionner(a, o)
    
    MsgBox "Le résultat de l'addition de " & a & " + " & o & " = " & resultatDuCalcul, vbInformation, "ELEOB"
    
End Sub
