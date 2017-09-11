Attribute VB_Name = "GameOfLife"
Public Const version_nr As String = "Wersja 0.1"

' Conway's Game of Life
' in Microsoft Excel
' by Konrad Ungeheuer

'Copyright (c) 2017 Konrad Ungeheuer
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

#If VBA7 Then

    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

#Else

    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds as Long)

#End If


' simulation arena size limits, works best when < 100
Public Const rowsize As Integer = 50
Public Const colsize As Integer = 50

' offsets for simulation arena
Public Const row_offset As Integer = 7 ' 2 minimum
Public Const column_offset As Integer = 5 ' 2 minimum

' menu position
Public Const menu_title_r As String = "D2:P2"

Public Const menu_cycles_r As String = "D3:G3"
Public Const menu_living_r As String = "H3:L3"
Public Const menu_delta_r As String = "M3:P3"

Public Const menu_cyclesn_r As String = "D4:G4"

' adress of cycles input for game_of_life_simulation_run()
Public Const menu_input_cycles As String = "$D$4"

Public Const menu_livingn_r As String = "H4:L4"


Public Const menu_deltan_r As String = "M4:P4"

' adress of delay input for game_of_life_simulation_run()
Public Const menu_input_delay As String = "$M$4"

' names of styles used in simulation
Const dead_style_name As String = "Dead"
Const living_style_name As String = "Liv"

' sheet name for simulation
Const sim_sheet_name As String = "Game_of_Life"

' global non const swith indicating if prep phase was done before simulation phase !
Dim is_prepared As Boolean

Option Explicit

Sub flush_arena_e()
'
' Clear "Arena" range and apply dead cell style
'


  With Range("Arena").Cells
    .clear
    .Style = dead_style_name
  End With
End Sub

Sub game_of_life_simulation_prepare()
'
' preparations required for game to run
' this macro need to be run first by user
' can be relaunch when needed
'

  Dim simulation_arena As Range
  Dim user_answer As Variant
  Dim sim_worksheet As Worksheet
  Dim update_scereen_state As Boolean
  Dim sim_workbook As Workbook
  
  Set sim_workbook = ThisWorkbook

  If Not SheetExists(sim_sheet_name, ThisWorkbook) Then
    Set sim_worksheet = Sheets.Add(after:=Sheets(Worksheets.Count))
    sim_worksheet.Name = sim_sheet_name
  Else
    Set sim_worksheet = Worksheets(sim_sheet_name)
  End If

' screen update turned of for menu draw etc
  update_scereen_state = APPLICATION.ScreenUpdating
  APPLICATION.ScreenUpdating = False

' activate workbook
  sim_worksheet.Activate

' clear sheet and delete shapes
  flush_sheet sim_worksheet

' sheet shape preparation
  MakeSquareCells sim_worksheet

' styles initialization
  DeadCellStyle
  LivingCellStyle

' simulation arena prep, rowsize * colsize with (1, 1) offset
  Set simulation_arena = sim_worksheet.Range(Cells(row_offset, column_offset), Cells(rowsize + row_offset, colsize + column_offset))
  
' this name is kinda global if changed flush_arena sub need change too
  simulation_arena.Name = "Arena"

' filling arena with dead cell style
  flush_arena_e ' simulation_arena

' drawing menu
  menu_draw sim_worksheet
  shapes_for_menu sim_workbook, sim_worksheet, simulation_arena

' border around simulation arena
  simple_border_maker Range(Cells(row_offset - 1, column_offset - 1), Cells(rowsize + row_offset + 1, colsize + column_offset + 1))

' screen update turned back
  APPLICATION.ScreenUpdating = update_scereen_state

' instructions for user
  MsgBox "    Now mark living cells within marked range using " _
    & living_style_name & _
    " style (black one), other cells are assumed dead and are marked by " & dead_style_name & _
    ". Then set the number of life cycles in " & menu_input_cycles & " and delay time in miliseconds in " & _
    menu_input_delay & Chr(13) & Chr(13) & _
    "    When you finish run game_of_life_simulation macro.", vbOKOnly, "Game of Life preparation phase"

' select arena
simulation_arena.Select

' prep compelte switch
is_prepared = True

End Sub
Sub game_of_life_simulation_run()
'
' macro running simulation in "Arena" range
' takes arguments from cells menu_input_cycles and menu_input_delay
' delay is best to leave 0
'
'

' amount of cycles to simulate
  Dim cycles As Integer
  ' delay in miliseconds
  Dim delay As Integer
  
  Dim step As Integer
  Dim ilu As Integer
  Dim dead As Boolean
  Dim update_screen_state As Boolean
  Dim kom As Range
  
  If Not is_prepared Then
    MsgBox "The game_of_life_simulation_prepare macro need to be runned first!", vbCritical
    Exit Sub
  End If
  
  If is_integer(Range(menu_input_cycles).Value) And CInt(Range(menu_input_cycles)) >= 0 Then
    cycles = CInt(Range(menu_input_cycles).Value)
  Else
     MsgBox "Input error " & menu_input_cycles & " positive integer required!", vbCritical
     Exit Sub
  End If
  
  If is_integer(Range(menu_input_delay)) And CInt(Range(menu_input_delay)) >= 0 Then
    delay = CInt(Range(menu_input_delay))
  Else
    MsgBox "Input error " & menu_input_delay & " positive integer required!", vbCritical
    Exit Sub
  End If
  
  ' main loop
  For step = 0 To cycles
    
      DoEvents
      
      Sleep (delay)
      
  ' screen update turned of for menu draw etc
      update_screen_state = APPLICATION.ScreenUpdating
      APPLICATION.ScreenUpdating = False
    
  ' number of living cells surrounding kom cell
    For Each kom In Range("Arena")
      ilu = 0
      If kom.Offset(-1, -1).Style = living_style_name Then
        ilu = ilu + 1
      End If
      
      If kom.Offset(0, -1).Style = living_style_name Then
        ilu = ilu + 1
      End If
      
      If kom.Offset(1, -1).Style = living_style_name Then
        ilu = ilu + 1
      End If
    
      If kom.Offset(1, 0).Style = living_style_name Then
        ilu = ilu + 1
      End If
      
      If kom.Offset(1, 1).Style = living_style_name Then
        ilu = ilu + 1
      End If
      
      If kom.Offset(0, 1).Style = living_style_name Then
        ilu = ilu + 1
      End If
      
      If kom.Offset(-1, 1).Style = living_style_name Then
        ilu = ilu + 1
      End If
    
      If kom.Offset(-1, 0).Style = living_style_name Then
        ilu = ilu + 1
      End If
      
      kom.Value = ilu
    Next
      
     ' DoEvents
      
      'life and dead conditions
    For Each kom In Range("Arena")
      With kom
        ilu = .Value
        dead = (.Style <> living_style_name)
      End With
      
      Select Case dead
        Case True 'cell is dead
          If ilu = 3 Then 'the only condition when cell can start living
            kom.Style = living_style_name
          End If
        Case False 'cell is alive, and if value = 2 or 3 it remain alive
          If ilu <> 2 And ilu <> 3 Then
            kom.Style = dead_style_name
          End If
        End Select 'otherwise cell alive/dead state don't change
    Next
      
      ' screen update turned back
      APPLICATION.ScreenUpdating = update_screen_state
      
      DoEvents
      
  Next step
End Sub

Private Sub shapes_for_menu(wrkbook As Workbook, wrksheet As Worksheet, arena As Range)  ', workbook_name As String)
'
' Put arrow and roundedsquare and assign macros to them
'
  Dim run As Shape, clear As Shape
  
  Set run = wrksheet.Shapes.addshape(msoShapeRightArrow, 300, 15, 41, 41)
  Set clear = wrksheet.Shapes.addshape(msoShapeRoundedRectangle, 250, 15, 41, 41)
  
  With run
    .Name = "Run simulation"
    .OnAction = "'" & wrkbook.Name & "'!game_of_life_simulation_run"
  End With
  
  With clear
    .Name = "Clear Arena"
    .OnAction = "'" & wrkbook.Name & "'!flush_arena_e"
  End With

End Sub

Private Sub menu_draw(sht As Worksheet)
'
' Simulation menu and inputs
'
  sht.Activate
  
  ' main title
  With Range(menu_title_r)
    .HorizontalAlignment = xlCenter
    .Merge
    .Style = "Heading 4"
    .Value = "Game of Life"
  End With
  
  ' description part
  With Range(menu_cycles_r)
    .HorizontalAlignment = xlCenter
    .Merge
    .Value = "Cycles"
  End With
  With Range(menu_living_r)
    .HorizontalAlignment = xlCenter
    .Merge
    .Value = "Living cells"
  End With
  With Range(menu_delta_r)
    .HorizontalAlignment = xlCenter
    .Merge
    .Value = "Delay"
  End With
  
  ' stat part
  With Range(menu_cyclesn_r)
    .HorizontalAlignment = xlCenter
    .Style = "Input"
    .Merge
    .Value = 0
  End With
  With Range(menu_livingn_r)
    .HorizontalAlignment = xlCenter
    .Style = "Good"
    .Merge
  End With
  With Range(menu_deltan_r)
    .HorizontalAlignment = xlCenter
    .Style = "Input"
    .Merge
    .Value = 0
  End With
End Sub

Private Sub simple_border_maker(rng As Range)
'
' draw a border around given range, assumes that range is rectangular
'
  With rng
    .Columns(1).Style = "Bad"
    .Columns(rng.Columns.Count).Style = "Bad"
    .Rows(1).Style = "Bad"
    .Rows(rng.Rows.Count).Style = "Bad"
  End With
End Sub

Private Sub flush_sheet(sht As Worksheet)
'
' Clear sheet sht and delete shapes
'

  Dim shp As Shape
  sht.Cells.clear
  For Each shp In sht.Shapes
    shp.Delete
  Next
End Sub

Private Sub MakeSquareCells(sht As Worksheet) ' https://superuser.com/questions/165738/how-to-make-cells-perfect-squares-in-excel
'
' Resize cells to make them Squares
'
    Dim update_scerees_state As Boolean
    update_scerees_state = APPLICATION.ScreenUpdating
    APPLICATION.ScreenUpdating = False
    With sht
        .Columns.ColumnWidth = 2 '// minimum 2, max 400 ; above 7 --> zoom doesn't work nice
        .Rows.EntireRow.RowHeight = .Cells(1).Width
    End With
    APPLICATION.ScreenUpdating = update_scerees_state
End Sub

Private Sub LivingCellStyle()
'
' Living Cell style script
'


' Remove style named <style_name_for_living> if it exist
Dim stl As Style
  For Each stl In ActiveWorkbook.Styles
    If stl.Name = living_style_name Then
      Exit Sub
    End If
  Next stl
    
ActiveWorkbook.Styles.Add Name:=living_style_name
  With ActiveWorkbook.Styles(living_style_name)
      .IncludeNumber = False
      .IncludeFont = True
      .IncludeAlignment = False
      .IncludeBorder = True
      .IncludePatterns = True
      .IncludeProtection = False
      With .Font
          .Name = "Calibri"
          .Size = 11
          .Bold = False
          .Italic = False
          .Underline = xlUnderlineStyleNone
          .Strikethrough = False
          .ThemeColor = 2
          .TintAndShade = 0
          .ThemeFont = xlThemeFontMinor
      End With
      With .Borders(xlLeft)
          .LineStyle = xlContinuous
          .TintAndShade = 0
          .Weight = xlThin
          .ColorIndex = 15
      End With
      With .Borders(xlRight)
          .LineStyle = xlContinuous
          .TintAndShade = 0
          .Weight = xlThin
          .ColorIndex = 15
      End With
      With .Borders(xlTop)
          .LineStyle = xlContinuous
          .TintAndShade = 0
          .Weight = xlThin
          .ColorIndex = 15
      End With
      With .Borders(xlBottom)
          .LineStyle = xlContinuous
          .TintAndShade = 0
          .Weight = xlThin
          .ColorIndex = 15
      End With
      With .Interior
          .Pattern = xlSolid
          .PatternColorIndex = 0
          .ThemeColor = xlThemeColorLight1
          .TintAndShade = 0
          .PatternTintAndShade = 0
      End With
  End With
End Sub

Private Sub DeadCellStyle()
'
' Living Cell style script
'


' Remove style named <dead_style_name> if it exist
Dim stl As Style
  For Each stl In ActiveWorkbook.Styles
    If stl.Name = dead_style_name Then
      Exit Sub
    End If
  Next stl


ActiveWorkbook.Styles.Add Name:=dead_style_name
  With ActiveWorkbook.Styles(dead_style_name)
      .IncludeNumber = False
      .IncludeFont = True
      .IncludeAlignment = False
      .IncludeBorder = True
      .IncludePatterns = True
      .IncludeProtection = False
      With .Font
          .Name = "Calibri"
          .Size = 11
          .Bold = False
          .Italic = False
          .Underline = xlUnderlineStyleNone
          .Strikethrough = False
          .ThemeColor = 1
          .TintAndShade = 0
          .ThemeFont = xlThemeFontMinor
      End With
      With .Borders(xlLeft)
          .LineStyle = xlContinuous
          .TintAndShade = 0
          .Weight = xlThin
          .ColorIndex = 15
      End With
      With .Borders(xlRight)
          .LineStyle = xlContinuous
          .TintAndShade = 0
          .Weight = xlThin
          .ColorIndex = 15
      End With
      With .Borders(xlTop)
          .LineStyle = xlContinuous
          .TintAndShade = 0
          .Weight = xlThin
          .ColorIndex = 15
      End With
      With .Borders(xlBottom)
          .LineStyle = xlContinuous
          .TintAndShade = 0
          .Weight = xlThin
          .ColorIndex = 15
      End With
      With .Interior
          .Pattern = xlSolid
          .PatternColorIndex = 0
          .ThemeColor = xlThemeColorDark1
          .TintAndShade = 0
          .PatternTintAndShade = 0
      End With
  End With
End Sub
