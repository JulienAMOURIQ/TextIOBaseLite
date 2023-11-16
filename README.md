# TextIOBaseLite
TextIOBase porté en VBScript. L'objectif est d'implémenter en VBS la classe python [TextIOBase](https://docs.python.org/fr/3/library/io.html#id1) .

## Usage
En lecture :

    dim ofichier : set ofichier = open("C:\exemple.txt", "r", "utf-8", "\r\n")
    dim lignes, i
    lignes = ofichier.readlines()
    For i = 0 To UBound(lignes)
      WScript.Echo lignes(i)
    Next

    ofichier.seek(0)
    WScript.Echo ofichier.readline()

En écriture :

    dim ofichier : set ofichier = open ("C:\exemple2.txt", "w", "utf-8", "\r\n")
    
    ofichier.write("une première ligne")
    ofichier.write("une deuxième ligne")
    
    ofichier.close()
    
    ofichier.write("ne sera pas écrit...")
