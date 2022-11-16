# MiniGdiPlusForVB6

This is a standard module of VB6, which inherited from VistaSwx's Gdip module.

## How To Use

Create a VB standard EXE project, then import this module into the project.

In startup form, type these codes to use MGP(MiniGdiPlus).

```vb
Private Sub Form_Load()
    InitializeGDIPlus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TerminateGDIPlus
End Sub

```

To show grahpics correctly, the autoredraw of form should be set to true. 
