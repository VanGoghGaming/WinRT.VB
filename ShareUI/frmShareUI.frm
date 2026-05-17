VERSION 5.00
Begin VB.Form frmShareUI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ShareUI Demo"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShareUI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShare 
      Caption         =   "Share This"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmShareUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ITypedEventHandlerDataTransferManagerDataRequested
Implements ITypedEventHandlerDataTransferManagerTargetApplicationChosen

Private DataTransferManager As IDataTransferManager, StorageItems As IVector_IStorageItem

Private Sub AddFilesForSharing(vFiles As Variant)
Dim vFile As Variant
    If Not IsArray(vFiles) Then vFiles = Array(vFiles)
    StorageItems.Clear
    With New cAwait
        For Each vFile In vFiles
            If .Await(StorageFileStatics.GetFileFromPathAsync(StrRef(vFile))) = AsyncStatus_Completed Then StorageItems.Append .GetResults
        Next vFile
    End With
End Sub

Private Sub Form_Load()
Dim DataTransferManagerInterop As IDataTransferManagerInterop, FolderLauncherOptions As IFolderLauncherOptions, vArray(0) As IDataTransferManager
    Set DataTransferManagerInterop = DataTransferManagerStatics: Set DataTransferManager = DataTransferManagerInterop.GetForWindow(hWnd, GuidFromArray(vArray))
    Set FolderLauncherOptions = NewObject("FolderLauncherOptions"): Set StorageItems = FolderLauncherOptions.ItemsToSelect
    AddFilesForSharing AppPath & "\These-are-not-the-droids-youre-looking-for.jpg"
    DataTransferManager.AddDataRequested Me: DataTransferManager.AddTargetApplicationChosen Me
End Sub

Private Sub ITypedEventHandlerDataTransferManagerDataRequested_Invoke(ByVal Sender As IDataTransferManager, ByVal Args As IDataRequestedEventArgs)
    With Args.Request.Data
        .Properties.Title = StrRef("These are not the droids you're looking for!")
        .SetStorageItems StorageItems, APIFALSE
    End With
End Sub

Private Sub ITypedEventHandlerDataTransferManagerTargetApplicationChosen_Invoke(ByVal Sender As IDataTransferManager, ByVal Args As ITargetApplicationChosenEventArgs)
    Debug.Print "You have shared to the " & GetStr(Args.ApplicationName) & " application!", App.Title
End Sub

Private Sub cmdShare_Click()
Dim DataTransferManagerInterop As IDataTransferManagerInterop
    Set DataTransferManagerInterop = DataTransferManagerStatics
    DataTransferManagerInterop.ShowShareUIForWindow hWnd
End Sub
