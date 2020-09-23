Attribute VB_Name = "Module1"

Global ip As String
Public Sub SaveListBox(TheList As ListBox, Directory As String)

    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1


    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&

    Close #1
End Sub

'Example: Call LoadListBox(list1, "C:\Temp\MyList.dat")


Public Sub LoadListBox(TheList As ListBox, Directory As String)

    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1


    While Not EOF(1)
        Input #1, MyString$


        DoEvents
            TheList.AddItem MyString$
        Wend

        Close #1
        
    End Sub



Public Sub PrintListBox(TheList As ListBox)

    Dim SaveList As Long
    On Error Resume Next
    Printer.FontSize = 12


    For SaveList& = 0 To TheList.ListCount - 1
        Printer.Print TheList.List(SaveList&)
    Next SaveList&

    Printer.EndDoc
End Sub










