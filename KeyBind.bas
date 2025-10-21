Attribute VB_Name = "KeyBind"
Sub BindKeys()
    Application.OnKey "{LEFT}", "moveLeft"
    Application.OnKey "{RIGHT}", "moveRight"
    Application.OnKey "{UP}", "moveUp"
    Application.OnKey "{DOWN}", "moveDown"
End Sub
Sub Unbindkeys()
Application.OnKey "{UP}"
Application.OnKey "{DOWN}"
Application.OnKey "{LEFT}"
Application.OnKey "{RIGHT}"
Application.OnKey " "
End Sub
Sub MoveLeft()
    If game <> True Then Exit Sub
    If cinc <> 1 Then
        cinc = -1
        rinc = 0
    End If
    MoveSnake

End Sub
Sub MoveRight()
    If game <> True Then Exit Sub
    If cinc <> -1 Then
        cinc = 1
        rinc = 0
    End If
    MoveSnake
End Sub
Sub MoveUp()
    If game <> True Then Exit Sub
    If rinc <> 1 Then
        cinc = 0
        rinc = -1
    End If
    MoveSnake

End Sub
Sub MoveDown()
    
    If game <> True Then Exit Sub
    If rinc <> -1 Then
        cinc = 0
        rinc = 1
    End If
    MoveSnake
End Sub
Sub freeKey()
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
End Sub
