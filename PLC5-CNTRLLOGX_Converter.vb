Sub PLC5_to_CNTL_Logix_Converter()
'''''''''''''''''''PLEASE READ BEFORE RUNNING MACRO'''''''''''''''''''''

'one step must be done by the user
'use Cntl-F to find all XXXPLC values
'replace with XXXPLC but with BASE RED font : RGB (255,0,0)

''''''''''''''Macro Begins''''''''''''''''''
 'Variable declarations: all global
 Dim FindRange As Range, c As Range
 Dim periodloc As Integer, slashloc As Integer, IOint As Integer, State As Integer, Config_State As Integer
 Dim IOnameLen As Integer, ArrayNum As Integer, bitnum As Integer, colonloc As Integer
 Dim y As Long, Z As Long, x As Long, i As Long
 Dim DataSource As String, ReturnString As String, ArrayNum_Str As String, bitnum_Str As String, termL As String
 Dim IOint_Str As String, IOValue As String, NewColumn As String, OrigColumn As String, IOname As String
 Dim Done As Boolean
 
 'initalize variables
 State = 0
 y = 1
 x = 1
 i = 1
 m = 0
 Z = 1
 OrigColumn = "A"
 NewColumn = "B"
 x = 1
 'Create Columns
 Columns(OrigColumn).Insert (xlDown)
 Columns(NewColumn).Insert (xlDown)

'beginning of function
 
 ''''''''''stage 1 - find all instances of DataBase String'''''''''''
 For Each c In ActiveSheet.UsedRange
 If c.Font.Color = RGB(255, 0, 0) Then
          If FindRange Is Nothing Then
                  Set FindRange = c
          Else
                  Set FindRange = Union(FindRange, c)
          End If
 End If
 Next
'copy everything from range into column A
If Not FindRange Is Nothing Then
      For Each c In FindRange
       Cells(x, OrigColumn) = c.Value
       x = x + 1
     Next
     State = 1
 End If
 
''''''''''Stage 2 Create Control Logix tags'''''''''''

While Not y = x

    If State = 1 Then 'grab original tag
            IOValue = Cells(y, OrigColumn)
            State = 2
    End If 'state 1 end
    
    If State = 2 Then 'isolate changeable value
            termL = Len(IOValue)
            DataSource = Left(IOValue, InStr(IOValue, "."))
            IOname = Right(IOValue, termL - Len(DataSource))
            IOnameLen = Len(IOname)
            State = 3
    End If 'state 2 end
    
    If State = 3 Then 'determine IOtype through  pattern - classify as Config_State
            'determine the type of variable based off of characters
            colonloc = InStr(IOname, ":")
            periodloc = InStr(IOname, ".")
            slashloc = InStr(IOname, "/")
            Done = False
            'Classify the variable to Config_State
                If Not periodloc = 0 And Not colonloc = 0 Then
                    Config_State = 1
                ElseIf Not colonloc = 0 And Not slashloc = 0 Then
                    Config_State = 2
                Else:
                    Done = True
                End If
                    
                If Done = True Then
                    If Not colonloc = 0 Then
                        Config_State = 3
                    ElseIf Not slashloc = 0 Then
                        Config_State = 4
                    Else:
                        Config_State = 5
                    End If
                End If
                State = 4
End If 'state 3 end

'state 4 creates a new string based of classification type Config_State
    If State = 4 Then
    
                If Config_State = 1 Then
                    '    XX:YY.ZZ format to XX[YY].ZZ
                    IOname = Replace(IOname, ":", "[")
                    IOname = Replace(IOname, ".", "].")
                    IOint_Str = IOname
                
                ElseIf Config_State = 2 Then
                    '    XX:YY/ZZ format to XX[YY].ZZ
                    IOname = Replace(IOname, ":", "[")
                    IOname = Replace(IOname, "/", "].")
                    IOint_Str = IOname
                ElseIf Config_State = 3 Then
                    '    XX:YY format to XX[YY]
                    ArrayNum_Str = Right(IOname, IOnameLen - colonloc) ' Extract YY
                    IOname = Left(IOname, colonloc - 1)
                    IOint_Str = (IOname + "[" + ArrayNum_Str + "]")
                    
                ElseIf Config_State = 4 Then
                    '    XX/YY format to XX[Quotient(YY)].Mod(YY)
                    IOint = Right(IOname, IOnameLen - slashloc)
                    ArrayNum = CInt(IOint)
                    ArrayNum = ArrayNum \ 16 'Quotient
                    ArrayNum_Str = CStr(ArrayNum)
                    bitnum = CInt(IOint)
                    bitnum = bitnum - (16 * (bitnum \ 16)) 'Modulus Equation
                    bitnum_Str = CStr(bitnum)
                    IOname = Left(IOname, slashloc - 1)
                    IOint_Str = (IOname + "[" + ArrayNum_Str + "]." + bitnum_Str)
        
                ElseIf Config_State = 5 Then
                'if IOValue is a text string then leave it
                    IOint_Str = IOname
                End If
                
            'concatenate new string place each value in column B
            ReturnString = DataSource + IOint_Str
            Cells(y, NewColumn) = ReturnString
            If Not y = x Then
            State = 1
            End If
       End If 'state 4 End
       y = y + 1 'incriment to next row
Wend 'final wend statement


''''''''''Stage 3 Replace and Delete'''''''''''

'For Loop goes through each row replaces each item
For Z = 1 To y
    ReturnString = Cells(i, NewColumn).Value
    IOValue = Cells(i, OrigColumn).Value
    'replace function
    ActiveSheet.Cells.Replace What:=IOValue, Replacement:=ReturnString, _
                    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                    SearchFormat:=False, ReplaceFormat:=False
    i = i + 1
Next
'Deletes created Columns
Debug.Print ("all done")
Columns(NewColumn).Delete
Columns(OrigColumn).Delete
End Sub