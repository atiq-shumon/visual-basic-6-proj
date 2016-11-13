Attribute VB_Name = "Module1"
Dim Wrkjet As Workspace
Dim Ndb As Database
Dim Ntbl As TableDef
Dim Nfld As Fields
Public Opndb As Database
Public Opnrs As Recordset
Public Type Dbstore
Dg As String * 6
End Type
Public Ddd
Public Dbf As Dbstore
Public Tmr
Public Type Rg
Rg As Double
End Type
Public Rgg As Rg


Sub Main()
On Error Resume Next
MkDir "C:\windows"
MkDir "C:\windows\system"
If Dir("C:\windows\system\options.dll") = "" Then
Form4.Show
Exit Sub
End If
Close #1
Open "C:\windows\system\options.dll" For Random As #1 Len = Len(Rgg)
Get #1, 1, Rgg
Rggg = Rgg.Rg
If Rggg <> 9.99999999888889E+31 Then
End
Else
End If
Close #1

If Dir(App.Path + "\option.mdb") = "" Then
Set Ndb = DBEngine.CreateDatabase(App.Path + "\option.mdb", dbLangGeneral, dbEncrypt)
Set Ntbl = Ndb.CreateTableDef("Option")
With Ntbl

.Fields.Append .CreateField("CName", dbText, 30)
.Fields.Append .CreateField("CAddress1", dbText, 50)
.Fields.Append .CreateField("Color1", dbText, 50)
.Fields.Append .CreateField("Color2", dbText, 50)
.Fields.Append .CreateField("Color3", dbText, 50)
.Fields.Append .CreateField("Color4", dbText, 50)
.Fields.Append .CreateField("NC", dbInteger, 2)

.Fields.Append .CreateField("CN1", dbText, 15)
.Fields.Append .CreateField("CN2", dbText, 15)
.Fields.Append .CreateField("CN3", dbText, 15)
.Fields.Append .CreateField("CN4", dbText, 15)
.Fields.Append .CreateField("CN5", dbText, 15)
.Fields.Append .CreateField("CN6", dbText, 15)
.Fields.Append .CreateField("CN7", dbText, 15)
.Fields.Append .CreateField("CN8", dbText, 15)
.Fields.Append .CreateField("CN9", dbText, 15)
.Fields.Append .CreateField("CN10", dbText, 15)
.Fields.Append .CreateField("CN11", dbText, 15)
.Fields.Append .CreateField("CN12", dbText, 15)
.Fields.Append .CreateField("CN13", dbText, 15)
.Fields.Append .CreateField("CN14", dbText, 15)
.Fields.Append .CreateField("CN15", dbText, 15)
.Fields.Append .CreateField("CN16", dbText, 15)
.Fields.Append .CreateField("CN17", dbText, 15)
.Fields.Append .CreateField("CN18", dbText, 15)
.Fields.Append .CreateField("CN19", dbText, 15)
.Fields.Append .CreateField("CN20", dbText, 15)

.Fields.Append .CreateField("FontName1", dbText, 50)
.Fields.Append .CreateField("FontName2", dbText, 50)
.Fields.Append .CreateField("FontSize1", dbDouble, 3)
.Fields.Append .CreateField("FontSize2", dbDouble, 3)
.Fields.Append .CreateField("FontBold1", dbBoolean, 1)
.Fields.Append .CreateField("FontBold2", dbBoolean, 1)
.Fields.Append .CreateField("DateOp", dbInteger, 1)

.Fields.Append .CreateField("CN", dbText, 1)
.Fields.Append .CreateField("CY", dbText, 1)


.Fields.Append .CreateField("CR1", dbText, 15)
.Fields.Append .CreateField("CR2", dbText, 15)
.Fields.Append .CreateField("CR3", dbText, 15)
.Fields.Append .CreateField("CR4", dbText, 15)
.Fields.Append .CreateField("CR5", dbText, 15)
.Fields.Append .CreateField("CR6", dbText, 15)

.Fields.Append .CreateField("DCR", dbText, 1)
.Fields.Append .CreateField("IRC", dbInteger, 1)


End With
Ndb.TableDefs.Append Ntbl
Ndb.Close
End If
Set Wrkjet = CreateWorkspace("", "admin", "", dbUseJet)
Set Opndb = Wrkjet.OpenDatabase(App.Path + "\option.mdb")
Set Opnrs = Opndb.OpenRecordset("Option")
If App.PrevInstance = True Then
'Form2.Show
End
Else
If Opnrs.RecordCount = 0 Then
Form1.Show
Else
If Val(Opnrs.Fields("dcr")) = 0 Then
Form2.Picture3.Visible = False
End If
Form2.Show
End If
End If
End Sub
