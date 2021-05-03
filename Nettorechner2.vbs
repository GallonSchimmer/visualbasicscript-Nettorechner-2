' Eingabe
brutto = InputBox("Geben Sie den Brutto-Betrag ein!","Eingabe Brutto")
IF NOT IsNumeric(brutto) THEN 'wenn Buchstaben eingegeben wurden 
   MsgBox "Keine Zahl eingegeben!",vbCritical,"Fehler"
   WScript.quit()  'Programmende
END IF
IF isEmpty(brutto) THEN  'wenn Abbrechen, dann ist brutto leer
   MsgBox "Abbrechen geklickt",,"Abbruch"
   WScript.quit()  ' Programmende
END IF
' Verarbeitung (Berechnung)
netto = brutto / 119 * 100
'Ausgabe
ergebnis = "Brutto:" & vbTab & vbTab & FormatCurrency(brutto) & _
           vbNewline & _
           "/ 119 * 100:" & vbTab & FormatCurrency(netto) & _
           vbNewline & _
           "Gesamt: " & vbTab & vbTab & FormatCurrency(netto)


MsgBox ergebnis,,"Ergebnis"