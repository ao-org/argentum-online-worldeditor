Attribute VB_Name = "modPaneles"
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

''
' modPaneles
'
' @remarks Funciones referentes a los Paneles de Funcion
' @author gshaxor@gmail.com
' @version 0.3.28
' @date 20060530

Option Explicit

''
' Activa/Desactiva el Estado de la Funcion en el Panel Superior
'
' @param Numero Especifica en numero de funcion
' @param Activado Especifica si esta o no activado

Public Sub EstSelectPanel(ByVal Numero As Byte, ByVal Activado As Boolean)
    
    On Error GoTo EstSelectPanel_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 30/05/06
    '*************************************************
    If Activado = True Then
        FrmMain.SelectPanel(Numero).GradientMode = lv_Bottom2Top
        FrmMain.SelectPanel(Numero).HoverBackColor = FrmMain.SelectPanel(Numero).GradientColor

        If FrmMain.mnuVerAutomatico.Checked = True Then

            Select Case Numero

                Case 0

                    If FrmMain.cCapas.Text = 4 Then
                        FrmMain.mnuVerCapa4.Tag = CInt(FrmMain.mnuVerCapa4.Checked)
                        FrmMain.mnuVerCapa4.Checked = True
                    ElseIf FrmMain.cCapas.Text = 3 Then
                        FrmMain.mnuVerCapa3.Tag = CInt(FrmMain.mnuVerCapa3.Checked)
                        FrmMain.mnuVerCapa3.Checked = True
                    ElseIf FrmMain.cCapas.Text = 2 Then
                        FrmMain.mnuVerCapa2.Tag = CInt(FrmMain.mnuVerCapa2.Checked)
                        FrmMain.mnuVerCapa2.Checked = True

                    End If

                Case 2
                    FrmMain.cVerBloqueos.Tag = CInt(FrmMain.cVerBloqueos.Value)
                    FrmMain.cVerBloqueos.Value = True
                    FrmMain.mnuVerBloqueos.Checked = FrmMain.cVerBloqueos.Value

                Case 6
                    FrmMain.cVerTriggers.Tag = CInt(FrmMain.cVerTriggers.Value)
                    FrmMain.cVerTriggers.Value = True
                    FrmMain.mnuVerTriggers.Checked = FrmMain.cVerTriggers.Value

            End Select

        End If

    Else
        FrmMain.SelectPanel(Numero).HoverBackColor = FrmMain.SelectPanel(Numero).BackColor
        FrmMain.SelectPanel(Numero).GradientMode = lv_NoGradient

        If FrmMain.mnuVerAutomatico.Checked = True Then

            Select Case Numero

                Case 0

                    If FrmMain.cCapas.Text = 4 Then
                        If LenB(FrmMain.mnuVerCapa3.Tag) <> 0 Then FrmMain.mnuVerCapa4.Checked = CBool(FrmMain.mnuVerCapa4.Tag)
                    ElseIf FrmMain.cCapas.Text = 3 Then

                        If LenB(FrmMain.mnuVerCapa3.Tag) <> 0 Then FrmMain.mnuVerCapa3.Checked = CBool(FrmMain.mnuVerCapa3.Tag)
                    ElseIf FrmMain.cCapas.Text = 2 Then

                        If LenB(FrmMain.mnuVerCapa2.Tag) <> 0 Then FrmMain.mnuVerCapa2.Checked = CBool(FrmMain.mnuVerCapa2.Tag)

                    End If

                Case 2

                    If LenB(FrmMain.cVerBloqueos.Tag) = 0 Then FrmMain.cVerBloqueos.Tag = 0
                    FrmMain.cVerBloqueos.Value = CBool(FrmMain.cVerBloqueos.Tag)
                    FrmMain.mnuVerBloqueos.Checked = FrmMain.cVerBloqueos.Value

                Case 6

                    If LenB(FrmMain.cVerTriggers.Tag) = 0 Then FrmMain.cVerTriggers.Tag = 0
                    FrmMain.cVerTriggers.Value = CBool(FrmMain.cVerTriggers.Tag)
                    FrmMain.mnuVerTriggers.Checked = FrmMain.cVerTriggers.Value

            End Select

        End If

    End If

    
    Exit Sub

EstSelectPanel_Err:
    Call RegistrarError(Err.Number, Err.Description, "modPaneles.EstSelectPanel", Erl)
    Resume Next
    
End Sub

''
' Muestra los controles que componen a la funcion seleccionada del Panel
'
' @param Numero Especifica el numero de Funcion
' @param Ver Especifica si se va a ver o no
' @param Normal Inidica que ahi que volver todo No visible

Public Sub VerFuncion(ByVal Numero As Byte, ByVal Ver As Boolean, Optional Normal As Boolean)
    
    On Error GoTo VerFuncion_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************
    If Normal = True Then
        Call VerFuncion(vMostrando, False, False)

    End If

    Select Case Numero

        Case 0 ' Superficies
            FrmMain.lListado(0).Visible = Ver
            FrmMain.cFiltro(0).Visible = Ver
            FrmMain.cCapas.Visible = Ver
            FrmMain.cGrh.Visible = Ver
            FrmMain.cQuitarEnEstaCapa.Visible = Ver
            FrmMain.cQuitarEnTodasLasCapas.Visible = Ver
            FrmMain.cSeleccionarSuperficie.Visible = Ver
            FrmMain.lbFiltrar(0).Visible = Ver
            FrmMain.lbCapas.Visible = Ver
            FrmMain.lbGrh.Visible = Ver
            FrmMain.PreviewGrh.Visible = Ver

            If Ver = True Then
                FrmMain.StatTxt.Top = 655
                FrmMain.StatTxt.Height = 37
            Else
                FrmMain.StatTxt.Top = 416
                FrmMain.StatTxt.Height = 270

            End If

        Case 1 ' Translados
            FrmMain.lMapN.Visible = Ver
            FrmMain.lXhor.Visible = Ver
            FrmMain.lYver.Visible = Ver
            FrmMain.tTMapa.Visible = Ver
            FrmMain.tTX.Visible = Ver
            FrmMain.tTY.Visible = Ver
            FrmMain.cInsertarTrans.Visible = Ver
            FrmMain.cInsertarTransOBJ.Visible = Ver
            FrmMain.cUnionManual.Visible = Ver
            FrmMain.cUnionAuto.Visible = Ver
            FrmMain.cQuitarTrans.Visible = Ver

        Case 2 ' Bloqueos
            FrmMain.cQuitarBloqueo.Visible = Ver
            FrmMain.cInsertarBloqueo.Visible = Ver
            FrmMain.cVerBloqueos.Visible = Ver
            Dim i As Integer

            For i = 0 To 3
                FrmMain.chkBloqueo(i).Visible = Ver
            Next
            FrmMain.BloqAll.Visible = Ver
            
        Case 3  ' NPCs
            FrmMain.lListado(1).Visible = Ver
            FrmMain.cFiltro(1).Visible = Ver
            FrmMain.lbFiltrar(1).Visible = Ver
            FrmMain.lNumFunc(Numero - 3).Visible = Ver
            FrmMain.cNumFunc(Numero - 3).Visible = Ver
            FrmMain.cInsertarFunc(Numero - 3).Visible = Ver
            FrmMain.cQuitarFunc(Numero - 3).Visible = Ver
            FrmMain.cAgregarFuncalAzar(Numero - 3).Visible = Ver
            FrmMain.lCantFunc(Numero - 3).Visible = Ver
            FrmMain.cCantFunc(Numero - 3).Visible = Ver

        Case 4 ' NPCs Hostiles

            'frmMain.lListado(1).Visible = Ver
            'frmMain.cFiltro(1).Visible = Ver
            'frmMain.lbFiltrar(1).Visible = Ver
            'frmMain.lNumFunc(Numero - 3).Visible = Ver
            'frmMain.cNumFunc(Numero - 3).Visible = Ver
            'frmMain.cInsertarFunc(Numero - 3).Visible = Ver
            'frmMain.cQuitarFunc(Numero - 3).Visible = Ver
            'frmMain.cAgregarFuncalAzar(Numero - 3).Visible = Ver
            'frmMain.lCantFunc(Numero - 3).Visible = Ver
            'frmMain.cCantFunc(Numero - 3).Visible = Ver
        Case 5 ' OBJs
            FrmMain.lListado(3).Visible = Ver
            FrmMain.cFiltro(3).Visible = Ver
            FrmMain.lbFiltrar(3).Visible = Ver
            FrmMain.lNumFunc(Numero - 3).Visible = Ver
            FrmMain.cNumFunc(Numero - 3).Visible = Ver
            FrmMain.cInsertarFunc(Numero - 3).Visible = Ver
            FrmMain.cQuitarFunc(Numero - 3).Visible = Ver
            FrmMain.cAgregarFuncalAzar(Numero - 3).Visible = Ver
            FrmMain.lCantFunc(Numero - 3).Visible = Ver
            FrmMain.cCantFunc(Numero - 3).Visible = Ver

        Case 6 ' Triggers
            FrmMain.cQuitarTrigger.Visible = Ver
            FrmMain.cInsertarTrigger.Visible = Ver
            FrmMain.cVerTriggers.Visible = Ver
            FrmMain.lListado(4).Visible = Ver
            FrmMain.TiggerEspecial.Visible = Ver

        Case 7 ' Particulas
            FrmMain.insertarParticula.Visible = Ver
            FrmMain.numerodeparticula.Visible = Ver
            FrmMain.quitarparticula.Visible = Ver
            FrmMain.ListaParticulas.Visible = Ver
            FrmMain.mnuVerParticulas.Checked = True
    
        Case 8 ' Luces
            FrmMain.Label8.Visible = Ver
            FrmMain.insertarLuz.Visible = Ver
            FrmMain.RangoLuz.Visible = Ver
            FrmMain.QuitarLuz.Visible = Ver
            FrmMain.LuzColor.Visible = Ver
            FrmMain.mnuVerLuces.Checked = True

    End Select

    If Ver = True Then
        vMostrando = Numero

        If Numero < 0 Or Numero > 8 Then Exit Sub
        If FrmMain.SelectPanel(Numero).Value = False Then
            FrmMain.SelectPanel(Numero).Value = True

        End If

    Else

        If Numero < 0 Or Numero > 8 Then Exit Sub
        If FrmMain.SelectPanel(Numero).Value = True Then
            FrmMain.SelectPanel(Numero).Value = False

        End If

    End If

    
    Exit Sub

VerFuncion_Err:
    Call RegistrarError(Err.Number, Err.Description, "modPaneles.VerFuncion", Erl)
    Resume Next
    
End Sub

''
' Filtra del Listado de Elementos de una Funcion
'
' @param Numero Indica la funcion a Filtrar

Public Sub Filtrar(ByVal Numero As Byte)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 26/05/06
    '*************************************************
    
    On Error GoTo Filtrar_Err
    

    Dim vMaximo As Integer
    Dim vDatos  As String
    Dim NumI    As Integer
    Dim i       As Integer
    Dim j       As Integer
    
    If FrmMain.cFiltro(Numero).ListCount > 5 Then
        FrmMain.cFiltro(Numero).RemoveItem 0

    End If

    FrmMain.cFiltro(Numero).AddItem FrmMain.cFiltro(Numero).Text
    FrmMain.lListado(Numero).Clear
        
    Select Case Numero

        Case 0 ' superficie
            vMaximo = MaxSup

        Case 1 ' NPCs
            vMaximo = NumNPCs - 1

        Case 2 ' NPCs Hostiles

            'vMaximo = NumNPCsHOST - 1
        Case 3 ' Objetos
            vMaximo = NumOBJs - 1

    End Select
    
    For i = 0 To vMaximo
    
        Select Case Numero

            Case 0 ' superficie
                vDatos = SupData(i).Name
                NumI = i

            Case 1 ' NPCs
                vDatos = NpcData(i + 1).Name
                NumI = i + 1

            Case 2 ' NPCs Hostiles

                'vDatos = NpcData(i + 500).name
                'NumI = i + 500
            Case 3 ' Objetos
                vDatos = ObjData(i + 1).Name
                NumI = i + 1

        End Select
        
        If LenB(vDatos) > 0 Then
        
            For j = 1 To Len(vDatos)
    
                If UCase$(mid$(vDatos & str(i), j, Len(FrmMain.cFiltro(Numero).Text))) = UCase$(FrmMain.cFiltro(Numero).Text) Or LenB(FrmMain.cFiltro(Numero).Text) = 0 Then
                    FrmMain.lListado(Numero).AddItem vDatos & " - #" & NumI
                    Exit For
    
                End If
    
            Next
        End If
    Next

    
    Exit Sub

Filtrar_Err:
    Call RegistrarError(Err.Number, Err.Description, "modPaneles.Filtrar", Erl)
    Resume Next
    
End Sub

Public Function DameGrhIndex(ByVal GrhIn As Long) As Long
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo DameGrhIndex_Err
    

    DameGrhIndex = SupData(GrhIn).Grh

    If SupData(GrhIn).Width > 0 Then
        frmConfigSup.MOSAICO.Value = vbChecked
        frmConfigSup.mAncho.Text = SupData(GrhIn).Width
        frmConfigSup.mLargo.Text = SupData(GrhIn).Height
    Else
        frmConfigSup.MOSAICO.Value = vbUnchecked
        frmConfigSup.mAncho.Text = "0"
        frmConfigSup.mLargo.Text = "0"

    End If

    
    Exit Function

DameGrhIndex_Err:
    Call RegistrarError(Err.Number, Err.Description, "modPaneles.DameGrhIndex", Erl)
    Resume Next
    
End Function

Public Sub fPreviewGrh(ByVal GrhIn As Long)
    '*************************************************
    'Author: Unkwown
    'Last modified: 22/05/06
    '*************************************************
    
    On Error GoTo fPreviewGrh_Err
    

    If Val(GrhIn) < 1 Then
        FrmMain.cGrh.Text = MaxGrhs
        Exit Sub

    End If

    If Val(GrhIn) > MaxGrhs Then
        FrmMain.cGrh.Text = 1
        Exit Sub

    End If

    'Change CurrentGrh
    CurrentGrh.GrhIndex = GrhIn
    CurrentGrh.Started = 1
    CurrentGrh.FrameCounter = 1

    
    Exit Sub

fPreviewGrh_Err:
    Call RegistrarError(Err.Number, Err.Description, "modPaneles.fPreviewGrh", Erl)
    Resume Next
    
End Sub

''
' Indica la accion de mostrar Vista Previa de la Superficie seleccionada
'

Public Sub VistaPreviaDeSup()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 26/05/06
    '*************************************************
    On Error Resume Next

    Dim SR  As RECT, DR As RECT
    Dim aux As Long
    Dim Tem As Long
    Dim w As Long, h As Long

    If CurrentGrh.GrhIndex = 0 Then Exit Sub
    frmGrafico.ShowPic = frmGrafico.Picture1

    If frmConfigSup.MOSAICO = vbUnchecked Then
    
        DR.Left = 0
        DR.Top = 0
        DR.Bottom = (GrhData(CurrentGrh.GrhIndex).pixelHeight)
        DR.Right = (GrhData(CurrentGrh.GrhIndex).pixelWidth)
        SR.Left = GrhData(CurrentGrh.GrhIndex).sX
        SR.Top = GrhData(CurrentGrh.GrhIndex).sY
        SR.Bottom = SR.Top + (GrhData(CurrentGrh.GrhIndex).pixelHeight)
        SR.Right = SR.Left + (GrhData(CurrentGrh.GrhIndex).pixelWidth)
        
        h = frmConfigSup.mLargo.Text
        If h <= 0 Then h = 1
        w = frmConfigSup.mAncho.Text
        If w <= 0 Then w = 1
        
        aux = Val(FrmMain.cGrh.Text) + (((1 + 1) Mod h) * w) + ((1 + 1) Mod w)
        Call Grh_Render_To_Hdc(FrmMain.PreviewGrh, (aux), 0, 0, False)
    Else
    
        Dim X As Integer, y As Integer, j As Integer, i As Integer
        Dim ww As Integer, hh As Integer
        Dim Cont As Integer
        
        hh = Val(frmConfigSup.mLargo)
        ww = Val(frmConfigSup.mAncho)
        
        Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)

        For i = 1 To hh
            For j = 1 To ww
                DR.Left = (j - 1) * 32
                DR.Top = (i - 1) * 32
                DR.Right = j * 32
                DR.Bottom = i * 32
                SR.Left = GrhData(CurrentGrh.GrhIndex).sX
                SR.Top = GrhData(CurrentGrh.GrhIndex).sY
                SR.Right = SR.Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
                SR.Bottom = SR.Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
                Call Grh_Render_To_HdcSinBorrar(FrmMain.PreviewGrh, (CurrentGrh.GrhIndex), j * 32, i * 32)

                If Cont < hh * ww Then Cont = Cont + 1
                CurrentGrh.GrhIndex = CurrentGrh.GrhIndex + 1
            Next
        Next
 
        CurrentGrh.GrhIndex = CurrentGrh.GrhIndex - Cont

    End If
 
End Sub
