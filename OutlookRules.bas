Attribute VB_Name = "OutlookRules"
' Procedure to run Outlook rules
Sub RunRules()
    Dim oStores As Outlook.Stores
    Dim oStore As Outlook.Store
     
    Dim olRules As Outlook.Rules
    Dim myRule As Outlook.Rule

    Dim time1 As Date
    Dim time2 As Date
    Dim totaltime As String
    
    Set oStores = Application.Session.Stores

    Call WaitShow   ' display wait dialog
    
    time1 = Now()

    For Each oStore In oStores
    On Error Resume Next
    ' use the display name as it appears in the navigation pane
        If oStore.DisplayName = "JGesa@bsp.com.pg" Then
            Set olRules = oStore.GetRules()
        
            For Each myRule In olRules
                ' inbox belonging to oStore
                myRule.Execute ShowProgress:=False, Folder:=GetFolderPath(oStore.DisplayName & "\Inbox")
                myRule.Execute ShowProgress:=False, Folder:=GetFolderPath(oStore.DisplayName & "\Sent Items")
            Next
        End If
    Next
    
    Call KillWait   ' remove wait dialog
    
    time2 = Now()
    
    totaltime = ElapsedTime(time2, time1)
    
    MsgBox "All rules run successfully " & vbNewLine & "Time Elapsed: " & totaltime, vbOKOnly, "Rules Run Successfully"
End Sub

' Function to retrieve Outlook user folder
Function GetFolderPath(ByVal FolderPath As String) As Outlook.Folder
    Dim oFolder As Outlook.Folder
    Dim FoldersArray As Variant
    Dim i As Integer
        
    On Error GoTo GetFolderPath_Error
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If
    'Convert folderpath to array
    FoldersArray = Split(FolderPath, "\")
    Set oFolder = Application.Session.Folders.Item(FoldersArray(0))
    If Not oFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = oFolder.Folders
            Set oFolder = SubFolders.Item(FoldersArray(i))
            If oFolder Is Nothing Then
                Set GetFolderPath = Nothing
            End If
        Next
    End If
    'Return the oFolder
    Set GetFolderPath = oFolder
    Exit Function
        
GetFolderPath_Error:
    Set GetFolderPath = Nothing
    Exit Function
End Function

' Function to show wait dialog box
Sub WaitShow()
    Wait.Show vbModeless
End Sub

' Function to remove wait dialog box
Sub KillWait()
    Unload Wait
End Sub

Function ElapsedTime(endTime As Date, startTime As Date)
    Dim strOutput As String
    Dim Interval As Date

    ' Calculate the time interval.
    Interval = endTime - startTime

    ' Format and print the time interval in hours, minutes and seconds.
    strOutput = Int(CSng(Interval * 24)) & ":" & Format(Interval, "nn:ss") & " h:mm:ss"
    
    ElapsedTime = strOutput

End Function


