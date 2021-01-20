Sub Macro1()
Dim i As Integer
i = 1
While i > 0
Rows(i).Select
Selection.Copy
Selection.Insert Shift:=xlDown
Columns(3).Rows(i) = 1
Columns(2).Rows(i) =  ""
Columns(1).Rows(i + 1) =  ""
i = i - 1
Wend
End Sub

'marca	marca	qtd	desc	mat	quali	com	lar	0
