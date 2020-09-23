<div align="center">

## Advanced Auto Complete List Combo and List Boxs


</div>

### Description

This is a high class code article

i never be stingy with vb lovers .

I try and success to move the power of access to the power of vb .

The main features are :

1) the sub can be used with combo box and list box controls.

2) High performance.

3) faster with huge lists.

4) fully support of Delete and Backspace keys.

5)Case Sensitive.

6) of course few lines of code point to the purpose.

Try it and tell me.

If you found this really advanced please be faithful and vote.

Thanks.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Said Sowiny](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/said-sowiny.md)
**Level**          |Advanced
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/said-sowiny-advanced-auto-complete-list-combo-and-list-boxs__1-46441/archive/master.zip)





### Source Code

<B><FONT SIZE=2><P ALIGN="LEFT">'==================================</P>
<P ALIGN="LEFT">Public Sub AutoCompleteList(ByVal cboCtl As ComboBox, ByVal KeyCode As Integer)</P>
<P ALIGN="LEFT">'No comments it is clear .</P>
<P ALIGN="LEFT"> Dim Counter As Integer</P>
<P ALIGN="LEFT"> Dim Length As Integer</P>
<P ALIGN="LEFT"> If cboCtl.Text <> "" Then</P>
<P ALIGN="LEFT"> If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then</P>
<P ALIGN="LEFT">  KeyCode = 0</P>
<P ALIGN="LEFT">  Exit Sub</P>
<P ALIGN="LEFT"> End If</P>
<P ALIGN="LEFT"> For Counter = 0 To cboCtl.ListCount - 1</P>
<P ALIGN="LEFT">  Length = Len(cboCtl.Text)</P>
<P ALIGN="LEFT">  If Mid(cboCtl.List(Counter), 1, Len(cboCtl)) = cboCtl.Text Then</P>
<P ALIGN="LEFT">  cboCtl.Text = cboCtl.List(Counter)</P>
<P ALIGN="LEFT">  cboCtl.SelStart = Length</P>
<P ALIGN="LEFT">  cboCtl.SelLength = Len(cboCtl.Text)</P>
<P ALIGN="LEFT">  cboCtl.ListIndex = Counter</P>
<P ALIGN="LEFT">  Exit For</P>
<P ALIGN="LEFT">  End If</P>
<P ALIGN="LEFT"> Next</P>
<P ALIGN="LEFT"> End If</P>
<P ALIGN="LEFT">End Sub</P>
<P ALIGN="LEFT">'===========================================</P>
<P ALIGN="LEFT">Typical usage :</P>
<P ALIGN="LEFT">Private Sub cboCategory_Change()</P>
<P ALIGN="LEFT"> Call AutoCompleteList(cboCategory, mintKeyCode)</P>
<P ALIGN="LEFT">End Sub</P>
<P ALIGN="LEFT">Where mintKeyCode is a form variable given by a Form_KeyDown event like :</P>
<P ALIGN="LEFT">Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)</P>
<P ALIGN="LEFT"> mintKeyCode = KeyCode</P>
<P ALIGN="LEFT">End Sub</P>
<P ALIGN="LEFT">and the cboCategory is the combo box filled by your own items.</P>
<P ALIGN="LEFT">Taking into consideration the header of the form should look like this one :</P>
<P ALIGN="LEFT">Option Explicit</P>
<P ALIGN="LEFT">Private mintKeyCode As Integer</P>
<P ALIGN="LEFT">and your form KeyPreview Property is set to TRUE.</P>
<P ALIGN="LEFT">Enjoy ,</P></B></FONT>

