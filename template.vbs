Option Explicit

Dim outputFolder
outputFolder = WScript.Arguments.Item(1)

With CreateObject("Excel.Application")
  .AutomationSecurity = 3 ' msoAutomationSecurityForceDisable
  Dim workbook
  Set workbook = .Workbooks.Open(WScript.Arguments.Item(0),, True)
End With

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

If Not fso.FolderExists(outputFolder) Then
  fso.CreateFolder outputFolder
End If

Dim vbc
For Each vbc In workbook.VBProject.VBComponents
  Dim extension

  Select Case vbc.Type
      Case 1 ' vbext_ct_StdModule
      Case 100 ' vbext_ct_Document
          extension = ".bas"
      Case 2 ' vbext_ct_ClassModule
          extension = ".cls"
      Case 3 ' vbext_ct_MSForm
          extension = ".frm"
      Case Else
          Err.Raise 9
  End Select

  vbc.Export fso.BuildPath(outputFolder, vbc.name & extension)
Next

workbook.Close