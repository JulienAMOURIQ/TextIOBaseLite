# TextIOBaseLite
TextIOBase porté en VBScript. L'objectif est d'implémenter en VBS la classe python [TextIOBase](https://docs.python.org/fr/3/library/io.html#id1) .

## Usage

    dim ofichier : set ofichier = New TextIOBaseLite
    ofichier.open "C:\exemple.txt", "r", "utf-8", "\r\n"
    dim lignes, i
    lignes = ofichier.readlines()
    For i = 0 To UBound(lignes)
      WScript.Echo lignes(i)
    Next
