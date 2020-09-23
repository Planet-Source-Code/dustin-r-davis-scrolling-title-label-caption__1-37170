<div align="center">

## Scrolling Title/Label Caption


</div>

### Description

Simple, 4 lines of code, timer function that will scroll the title of your form, or a label caption, or the caption of a button or whatever. scroll something. this is very simple, and i dont know if anythign else like it is posted, i dont look before i post. I just thought this was neat since i'm bored and at work...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dustin R Davis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dustin-r-davis.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dustin-r-davis-scrolling-title-label-caption__1-37170/archive/master.zip)





### Source Code

```
'Can be form.caption, text1.text, label.caption, button1.caption, whatever.
Private Sub Timer1_Timer()
Dim Length As Long
DoEvents
Length = Len(Form1.Caption) - 1
Form1.Caption = Right$(Form1.Caption, Length) & Left(Form1.Caption, 1)
End Sub
```

