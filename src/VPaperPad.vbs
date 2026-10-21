Dim CurrentPath
Dim WordToFind

Set System=CreateObject("KiXtart.System")
Set fso=CreateObject("Scripting.FileSystemObject")

Set Form=System.Form()
Form.Text="VPaperPad"
Form.Height=512
Form.Width=1024
Form.Center
Set TextBox=Form.Controls.Add("TextBox")
TextBox.MultiLine=True
TextBox.WordWrap=False
TextBox.ScrollBars=3
TextBox.Anchor=15
TextBox.Height=Form.ClientHeight
TextBox.Width=Form.ClientWidth
TextBox.SetFocus

Set MainMenu=System.MainMenu()
Set FileMenu=MainMenu.MenuItems.Add()
FileMenu.Text="File"
Set FileNewMenu=FileMenu.MenuItems.Add()
FileNewMenu.Text="New"
Set FileOpenMenu=FileMenu.MenuItems.Add()
FileOpenMenu.Text="Open"
Set FileSaveMenu=FileMenu.MenuItems.Add()
FileSaveMenu.Text="Save"
Set FileSaveAsMenu=FileMenu.MenuItems.Add()
FileSaveAsMenu.Text="Save As"
Set FileExitMenu=FileMenu.MenuItems.Add()
FileExitMenu.Text="Exit"
Set EditMenu=MainMenu.MenuItems.Add()
EditMenu.Text="Edit"
Set EditSelectAllMenu=EditMenu.MenuItems.Add()
EditSelectAllMenu.Text="SelectAll"
Set EditGotoMenu=EditMenu.MenuItems.Add()
EditGotoMenu.Text="Goto"
Set EditCopyMenu=EditMenu.MenuItems.Add()
EditCopyMenu.Text="Copy"
Set EditCutMenu=EditMenu.MenuItems.Add()
EditCutMenu.Text="Cut"
Set EditPasteMenu=EditMenu.MenuItems.Add()
EditPasteMenu.Text="Paste"
Set EditDeleteMenu=EditMenu.MenuItems.Add()
EditDeleteMenu.Text="Delete"
Set EditFindMenu=EditMenu.MenuItems.Add()
EditFindMenu.Text="Find"
Set EditFindNextMenu=EditMenu.MenuItems.Add()
EditFindNextMenu.Text="Find Next"
Set EditReplaceMenu=EditMenu.MenuItems.Add()
EditReplaceMenu.Text="Replace"
Set OptionsMenu=MainMenu.MenuItems.Add()
OptionsMenu.Text="Options"
Set OptionsWordWrapMenu=OptionsMenu.MenuItems.Add()
OptionsWordWrapMenu.Text="Word Wrap"
Set HelpMenu=MainMenu.MenuItems.Add()
HelpMenu.Text="Help"
Set HelpAboutMenu=HelpMenu.MenuItems.Add()
HelpAboutMenu.Text="About"
Form.Menu=MainMenu

FileNewMenu.OnClick="FileNewMenu_Click()"
FileOpenMenu.OnClick="FileOpenMenu_Click()"
FileSaveMenu.OnClick="FileSaveMenu_Click()"
FileExitMenu.OnClick="Wscript.Quit"
HelpAboutMenu.OnClick="HelpAboutMenu_Click()"
TextBox.OnTextChanged="TextChanged()"
EditSelectAllMenu.OnClick="EditSelectAllMenu_Click()"
EditGotoMenu.OnClick="EditGotoMenu_Click()"
EditCopyMenu.OnClick="TextBox.Copy"
EditCutMenu.OnClick="TextBox.Cut"
EditPasteMenu.OnClick="TextBox.Paste"
EditDeleteMenu.OnClick="EditDeleteMenu_Click()"
EditFindMenu.OnClick="EditFindMenu_Click()"
OptionsWordWrapMenu.OnClick="OptionsWordWrapMenu_Click()"
FileSaveAsMenu.OnClick="FileSaveAsMenu_Click()"
EditFindNextMenu.OnClick="EditFindNextMenu_Click()"
EditReplaceMenu.OnClick="EditReplaceMenu_Click()"

Form.Visible="True"

Do While Form.Visible
    Wscript.Sleep 1
    Execute(Form.DoEvents)
Loop

Sub FileOpenMenu_Click()
    Set Form.OpenFileDialog=System.OpenFileDialog()
    Form.OpenFileDialog.Filter="txt files (*.txt)|*.txt|All files (*.*)|*.*"
    Form.OpenFileDialog.ShowDialog()
    If Form.OpenFileDialog.Filename<>"" Then
        Set File=fso.OpenTextFile(Form.OpenFileDialog.Filename)
        FileSize=fso.GetFile(Form.OpenFileDialog.Filename).Size
        If FileSize>0 Then
            TextBox.Text=File.ReadAll
        Else
            TextBox.Text=""
        End If
    End If
    Form.Text="VPaperPad"
    CurrentPath=Form.OpenFileDialog.Filename
End Sub

Sub FileSaveMenu_Click()
    If CurrentPath="" Then
        Set Form.SaveFileDialog=System.SaveFileDialog()
        Form.SaveFileDialog.Filter="txt files (*.txt)|*.txt|All files (*.*)|*.*"
        Form.SaveFileDialog.ShowDialog()
        If Form.SaveFileDialog.Filename<>"" Then
            Set File=fso.CreateTextFile(Form.SaveFileDialog.Filename)
            File.Write TextBox.Text
        End If
        Form.Text="VPaperPad"
        CurrentPath=Form.SaveFileDialog.Filename
    Else
        Set File=fso.CreateTextFile(Form.SaveFileDialog.Filename)
        File.Write TextBox.Text
        Form.Text="VPaperPad"
    End If
End Sub

Sub HelpAboutMenu_Click()
    Wscript.Echo "VPaperPad v0.1"
End Sub

Sub FileNewMenu_Click()
    Wscript.Sleep 10
    If TextBox.Text<>"" Then
        TextBox.Text=""
    End If
    Form.Text="VPaperPad"
    CurrentPath=""
    Wscript.Echo CurrentPath
End Sub

Sub TextChanged()
    Wscript.Sleep 10
    If Form.Text <> "*VPaperPad" Then
        Form.Text="*VPaperPad"
    End If
End Sub

Sub EditSelectAllMenu_Click()
    Wscript.Sleep 10
    TextBox.SelectionStart=0
    TextBox.SelectionLength=Len(TextBox.Text)
End Sub

Sub EditGotoMenu_Click()
    GotoLine=CInt(InputBox("Enter the line you would like to go to.", "VPaperPad"))
    Loops=1
    If GotoLine>1 Then
        Do
            NextLineBreak=Instr(NextLineBreak+1, TextBox.Text, Chr(13) & Chr(10))
            If NextLineBreak=0 Or Loops>=GotoLine-1 Then
                If NextLineBreak<>0 Then
                    TextBox.SelectionStart=NextLineBreak+1
                Else
                    Wscript.Echo "Out of range!"
                End If
                Exit Do
            End If
            Loops=Loops+1
        Loop
    Else
        TextBox.SelectionStart=0
    End If
End Sub

Sub EditDeleteMenu_Click()
    TextBox.SelectedText=""
End Sub

Sub EditFindMenu_Click()
    WordToFind=InputBox("Enter the word to search for.", "VPaperPad", WordToFind)
    If WordToFind<>"" Then
        If Mid(TextBox.Text, TextBox.SelectionStart+1, TextBox.SelectionLength)=WordToFind Then
            WordFound=Instr(TextBox.SelectionStart+1+TextBox.SelectionLength, TextBox.Text, WordToFind)
        Else
            WordFound=Instr(TextBox.SelectionStart+1, TextBox.Text, WordToFind)
        End If
        If WordFound>0 Then
            TextBox.SelectionStart=WordFound-1
            TextBox.SelectionLength=Len(WordToFind)
        End If
    End If
End Sub

Sub OptionsWordWrapMenu_Click()
    If TextBox.WordWrap=0 Then
        TextBox.WordWrap=1
    ElseIf TextBox.WordWrap=-1 Then
        TextBox.WordWrap=0
    End If
End Sub

Sub FileSaveAsMenu_Click()
    Set Form.SaveFileDialog=System.SaveFileDialog()
    Form.SaveFileDialog.Filter="All files (*.*)|*.*|txt files (*.txt)|*.txt"
    Form.SaveFileDialog.ShowDialog()
    If Form.SaveFileDialog.Filename<>"" Then
        Set File=fso.CreateTextFile(Form.SaveFileDialog.Filename)
        File.Write TextBox.Text
    End If
    Form.Text="VPaperPad"
    CurrentPath=Form.SaveFileDialog.Filename
    Form.Text="VPaperPad"
End Sub

Sub EditFindNextMenu_Click()
    If WordToFind="" Then
        WordToFind=InputBox("Enter the word to search for.", "VPaperPad")
    End If
    If Mid(TextBox.Text, TextBox.SelectionStart+1, TextBox.SelectionLength)=WordToFind Then
        WordFound=Instr(TextBox.SelectionStart+1+TextBox.SelectionLength, TextBox.Text, WordToFind)
    Else
        WordFound=Instr(TextBox.SelectionStart+1, TextBox.Text, WordToFind)
    End If
    If WordFound>0 Then
        TextBox.SelectionStart=WordFound-1
        TextBox.SelectionLength=Len(WordToFind)
    End If
End Sub

Sub EditReplaceMenu_Click()
    WordToReplace=InputBox("Enter the word to search for.", "VPaperPad", WordToReplace)
    WordReplaced=InputBox("Enter the word to replace.", "VPaperPad", WordReplaced)
    If WordToReplace<>"" Then
        If Mid(TextBox.Text, TextBox.SelectionStart+1, TextBox.SelectionLength)=WordToReplace Then
            WordFound=Instr(TextBox.SelectionStart+1+TextBox.SelectionLength, TextBox.Text, WordToReplace)
        Else
            WordFound=Instr(TextBox.SelectionStart+1, TextBox.Text, WordToReplace)
        End If
        If WordFound>0 Then
            TextBox.SelectionStart=WordFound-1
            TextBox.SelectionLength=Len(WordToReplace)
            TextBox.SelectedText=WordReplaced
            TextBox.SelectionStart=TextBox.SelectionStart-Len(WordReplaced)
            TextBox.SelectionLength=Len(WordReplaced)
        End If
    End If
End Sub
