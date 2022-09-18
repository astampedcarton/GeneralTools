VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrFind 
   Caption         =   "Word Enhanced Find"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14115
   OleObjectBlob   =   "usrFind.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usrFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'17SEP2022 - Hentie du Plessis
'Initial Creation
'
'Purpose:
'Enhanced Search for WORD. Search through a given folder for a specific word or phrase
'
'Updates
'
Option Explicit

Private Sub cmdFind_Click()
  Dim arFiles() As Variant
  Dim iar As Integer
  Dim ofso As Object
  Dim tmpDoc As Document
  Dim tmpRng As Range
    
  'E:\Documents\sas\SASUniversityEdition\myfolders\
  
  Set ofso = CreateObject("Scripting.FileSystemObject")
  
  arFiles = FindAllFiles(txtFileLoc.Text)
  
  lstResults.Clear
    
  For iar = 0 To UBound(arFiles)
    If ofso.FileExists(txtFileLoc.Text + "\" + arFiles(iar)) Then
        Documents.Open txtFileLoc.Text + "\" + arFiles(iar)
        Set tmpDoc = ActiveDocument
        Set tmpRng = ActiveDocument.Content
        
        With tmpRng.Find
          .MatchCase = chkMatchCase.Value
          .MatchWholeWord = chkWholeWords.Value
        End With
        
        tmpRng.Find.Execute FindText:=txtFind.Text
        If tmpRng.Find.Found Then
          lstResults.AddItem txtFileLoc.Text + "\" + arFiles(iar)
        End If
        tmpDoc.Close
        'lstResults.AddItem txtFileLoc.Text + "\" + arFiles(iar) + " - Close"
    End If
  Next iar
  'FindText (arFiles)
  
End Sub

Function FindText(arFlist() As Variant) As String
  Dim tmpDoc As Document
  Dim iar As Integer
  'for each arFlist
  For iar = 0 To UBound(arFlist)
    Documents.Open stloc + "\" + arFlist(iar)
    lstResults.AddItem stloc + "\" + arFlist(iar)
    Documents.Close
   
  Next iar
End Function

Function FindAllFiles(stloc As String) As Variant()
  Dim oFolder As Object
  Dim oFile, oFiles As Object
  Dim arFlist() As Variant
  Dim icnt As Integer
  Dim ofso As Object
  Dim res As Integer
  
  'E:\Documents\sas\SASUniversityEdition\myfolders\
  
  Set ofso = CreateObject("Scripting.FileSystemObject")
  
  If ofso.FolderExists(stloc) = True Then
    stloc = stloc
    Set oFolder = ofso.Getfolder(stloc)
    Set oFiles = oFolder.Files
    
    icnt = oFiles.Count
    
    ReDim arFlist(icnt)
    icnt = 0
    
    For Each oFile In oFiles
      If UCase(ofso.GetExtensionName(oFile.Name)) = UCase(cbExtensions.Text) Then
        arFlist(icnt) = oFile.Name
        icnt = icnt + 1
        'lstResults.AddItem stloc + "\" + oFile.Name
      End If
      Debug.Print arFlist(icnt)
    Next oFile
  Else
    res = MsgBox("Folder " + stloc + " does not exist", vbOKOnly)
  End If
  
  FindAllFiles = arFlist
  
End Function

Private Sub lstResults_Click()

End Sub

Private Sub lstResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Dim tmpDoc As Document

  Documents.Open lstResults.Text
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
  cbExtensions.AddItem "doc"
  cbExtensions.AddItem "docx"
  cbExtensions.AddItem "rtf"
  
End Sub
