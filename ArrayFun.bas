Attribute VB_Name = "ArrayFun"
'The purpose of this module is to make working with Arrays
'in VBA easier by providing functions to perform common data
'manipulation task.
'---------------------------------------------------------------
'Reference:
' Microsoft Scripting Runtime
'---------------------------------------------------------------
'List of Function
'  Transpose:
'  RemoveDuplicates:
'  Append:
'
'
' Developer: Hadi Ali
' Email: ali.hadi071@gmail.com
'
Option Explicit
Option Base 0

Public Function Transpose(ByVal varArray As Variant) As Variant

   Dim varResult As Variant
   Dim varTranposed() As Variant
   Dim dRow As Double
   Dim dCol As Double
      
   'make sure the imput is an array
   If TypeName(varArray) <> "Variant()" Then
      Transpose = Null
      Exit Function
   End If
   
   'make sure the array only has 2 dimmension
   If GetDims(varArray) <> 2 Then
      Transpose = Null
      Exit Function
   End If
   
   'declare a new array switching the dimmensions
   ReDim varTranposed(LBound(varArray, 2) To UBound(varArray, 2) _
         , LBound(varArray, 1) To UBound(varArray, 1))
         
   For dCol = LBound(varArray, 1) To UBound(varArray, 1)
      For dRow = LBound(varArray, 2) To UBound(varArray, 2)
         varTranposed(dRow, dCol) = varArray(dCol, dRow)
      Next dRow
   Next dCol
   
   varResult = varTranposed
   Transpose = varResult
      
   
End Function

Public Function RemoveDuplicates(ByVal varArray As Variant) As Variant
                                    
   Dim objUnique As New Dictionary
   Dim varResult As Variant
   Dim varNoDup As Variant
   Dim varType As Variant
   Dim varConcate As Variant
   Dim varTemp As Variant
   Dim dRowInd As Variant
   Dim lDims As Long
   Dim dCol As Double
   Dim dRow As Double
   Dim dNoDupInd As Double
   
   
   'make sure the imput is an array
   If TypeName(varArray) <> "Variant()" Then
      RemoveDuplicates = Null
      Exit Function
   End If
   
   'check the number of dimmensions
   lDims = GetDims(varArray)
   
   'remove distinct based on the type of dimmension
   Select Case lDims
   
      Case 1
      
         'add items that exist to a dictionary
         For dCol = LBound(varArray, 1) To UBound(varArray, 1)
            If objUnique.Exists(varArray(dCol)) = False Then
               objUnique.Add varArray(dCol), dCol
            End If
         Next dCol
         
         
         varResult = objUnique.Keys
         
      Case 2
         'convert all the columns in the array to a single
         'row then add to dictionary and see if they already exit
      
         'extract the vartype of each column and store it in
         'an array
         dNoDupInd = LBound(varArray, 2)
         ReDim varNoDup(LBound(varArray, 2) To UBound(varArray, 2), _
                        LBound(varArray, 1) To dNoDupInd)
         
         For dCol = LBound(varArray, 1) To UBound(varArray, 1)
            ReDim varTemp(LBound(varArray, 2) To UBound(varArray, 2))
            dRowInd = LBound(varArray, 2)
            
            For dRow = LBound(varArray, 2) To UBound(varArray, 2)
               varTemp(dRowInd) = varArray(dCol, dRow)
                  dRowInd = dRowInd + 1
            Next dRow
                  
            varConcate = VBA.Join(varTemp)
            'if it is a duplicate then do not add to the
            'dictionary if it not then add to the dictionary and
            'new array
            If objUnique.Exists(varConcate) = False Then
               objUnique.Add varConcate, dRow
               ReDim Preserve varNoDup(LBound(varArray, 2) To UBound(varArray, 2), _
                                       LBound(varArray, 1) To dNoDupInd)
               
               For dRow = LBound(varArray, 2) To UBound(varArray, 2)
                  varNoDup(dRow, dNoDupInd) = varArray(dCol, dRow)
               Next dRow
               dNoDupInd = dNoDupInd + 1
            End If
         Next dCol
         
         'tranpose the data back to the original format
         varResult = Transpose(varNoDup)
   
   End Select
   
   RemoveDuplicates = varResult
   
   
   
End Function

'===========================================================
' Append Public Function
' ----------------------------------------------------------
' Purpose: Merge two 1 or 2 dimmensional arrays into one array
'
' Author: Hadi Ali March 2019
'
'Parameter(s)
'-----------
'varArray1: An Array of 1 or 2 dimmensions
'varArray2: An Array of 1 or 2 dimmensions
'lngAppendDim: the dimmension for 2d array to append with
'-----------------------------------------------------------
'Returns: A merged array and null if an error occurs
'
'-----------------------------------------------------------
'Revision History
'-----------------------------------------------------------
'20Feb20 HA: Initial Version
'===========================================================

Public Function Append(ByVal varArray1 As Variant, _
                       ByVal varArray2 As Variant, _
                       Optional ByVal lngAppendDim As Long = 1) As Variant

   Dim varResult As Variant
   Dim varAllData As Variant
   Dim lngrow As Long
   Dim lngCol As Long
   Dim lngMaxDim As Long
   Dim lngIndex As Long
   Dim lngDims1 As Long
   Dim lngDims2 As Long
   
   'find the number of dimmensions for both arrays and only proced if
   'they are less then 2 or equal
   
   #If debugging Then
      Debug.Assert (lngAppendDim <= 2)
   #End If
   
   lngDims1 = GetDims(varArray1)
   lngDims2 = GetDims(varArray2)
   
   If lngDims1 = lngDims2 Then
      'append 1 dimmensional arrays
      If lngDims1 = 1 And lngDims2 = 1 Then
         ReDim varAllData(LBound(varArray1) To UBound(varArray1) + UBound(varArray1) + 1)
            lngIndex = LBound(varAllData)
            For lngrow = LBound(varArray1) To UBound(varArray1)
               varAllData(lngIndex) = varArray1(lngrow)
               lngIndex = lngIndex + 1
            Next lngrow
            For lngrow = LBound(varArray2) To UBound(varArray2)
               varAllData(lngIndex) = varArray2(lngrow)
               lngIndex = lngIndex + 1
            Next lngrow
            varResult = varAllData
            GoTo FunctionClose
      ElseIf lngDims1 = 2 And lngDims2 = 2 Then
         'append 2 dimmensional arrays
         
         If lngAppendDim = 1 Then
            'find the max dimmension
             lngMaxDim = Application.WorksheetFunction.Max(UBound(varArray1, 2), _
                     UBound(varArray2, 2))
                     
             ReDim varAllData(LBound(varArray1, 1) To UBound(varArray1, 1) + UBound(varArray1, 1) + 1, _
                              LBound(varArray1) To lngMaxDim)
             
             For lngrow = LBound(varArray1, 1) To UBound(varArray1, 1)
               For lngCol = LBound(varArray1, 2) To UBound(varArray1, 2)
                  varAllData(lngrow, lngCol) = varArray1(lngrow, lngCol)
               Next lngCol
             Next lngrow
             
             lngIndex = UBound(varArray1) + 1
             For lngrow = LBound(varArray2, 1) To UBound(varArray2, 1)
               For lngCol = LBound(varArray2, 2) To UBound(varArray2, 2)
                  varAllData(lngIndex, lngCol) = varArray2(lngrow, lngCol)
               Next lngCol
               lngIndex = lngIndex + 1
             Next lngrow
             
         ElseIf lngAppendDim = 2 Then
             lngMaxDim = Application.WorksheetFunction.Max(UBound(varArray1, 1), _
                     UBound(varArray2, 1))
             
             ReDim varAllData(LBound(varArray1) To lngMaxDim, _
             LBound(varArray1, 2) To UBound(varArray1, 2) + UBound(varArray1, 2) + 1)

             
             For lngCol = LBound(varArray1, 2) To UBound(varArray1, 2)
               For lngrow = LBound(varArray1, 1) To UBound(varArray1, 1)
                  varAllData(lngrow, lngCol) = varArray1(lngrow, lngCol)
               Next lngrow
             Next lngCol
             
             lngIndex = UBound(varArray1, 2) + 1
             For lngCol = LBound(varArray2, 2) To UBound(varArray2, 2)
               For lngrow = LBound(varArray2, 1) To UBound(varArray2, 1)
                  varAllData(lngrow, lngIndex) = varArray2(lngrow, lngCol)
               Next lngrow
               lngIndex = lngIndex + 1
             Next lngCol
             
         Else
            varResult = ""
            GoTo FunctionClose
         End If
         
      Else
         varResult = ""
         GoTo FunctionClose
      End If
   Else
      varResult = ""
      GoTo FunctionClose
   End If
   varResult = varAllData
FunctionClose:
   Append = varResult

End Function


''''''''''''''''''''''''''' Private Functions '''''''''''''''''''''''''''''''''''''''
'===========================================================
' GetDims Private Function
' ----------------------------------------------------------
' Purpose: Determines the number of dimmensions in an array
'
' Author: Hadi Ali November 2019
'
'Parameter(s)
'-----------
'varArray: An Array
'-----------------------------------------------------------
'Returns: The Number of dimmension in an array if an error
'         occurs then a null value is returned
'-----------------------------------------------------------
'Revision History
'-----------------------------------------------------------
'20Nov19 HA: Initial Version
'===========================================================

Private Function GetDims(ByVal varArray As Variant) As Long
   
   Dim LResult As Long
   Dim lDims As Long
   Dim Ltmp As Long
   
   On Error GoTo ErrorHandler:
   
   If TypeName(varArray) <> "Variant()" Then
      GetDims = 0
      Exit Function
   End If
   
   lDims = 0
   Do While True
      lDims = lDims + 1
      Ltmp = UBound(varArray, lDims)
   Loop
   Exit Function
ErrorHandler:
   GetDims = lDims - 1
End Function

'this function coverts a given value into a vba type based on
'the type integer provided

Private Function TypeConvert(varValue As Variant, intType As Integer) As Variant
   
   Dim varResult As Variant
   
   On Error Resume Next
   Select Case intType
      Case 2
         varResult = CInt(varValue)
         If Err.Number = 0 Then
            TypeConvert = varResult
         Else
            TypeConvert = "Error#"
         End If
      Case 3
         varResult = CLng(varValue)
         If Err.Number = 0 Then
            TypeConvert = varResult
         Else
            TypeConvert = "Error#"
         End If
      Case 4
         varResult = CSng(varValue)
         If Err.Number = 0 Then
            TypeConvert = varResult
         Else
            TypeConvert = "Error#"
         End If
      Case 5
         varResult = CDbl(varValue)
         If Err.Number = 0 Then
            TypeConvert = varResult
         Else
            TypeConvert = "Error#"
         End If
      Case 6
         varResult = CCur(varValue)
         If Err.Number = 0 Then
            TypeConvert = varResult
         Else
            TypeConvert = "Error#"
         End If
      Case 7
         varResult = CDate(varValue)
         If Err.Number = 0 Then
            TypeConvert = varResult
         Else
            TypeConvert = "Error#"
         End If
      Case 8
         varResult = CStr(varValue)
         If Err.Number = 0 Then
            TypeConvert = varResult
         Else
            TypeConvert = "Error#"
         End If
      Case 11
         varResult = CBool(varValue)
         If Err.Number = 0 Then
            TypeConvert = varResult
         Else
            TypeConvert = "Error#"
         End If
      Case 12
         varResult = CVar(varValue)
         If Err.Number = 0 Then
            TypeConvert = varResult
         Else
            TypeConvert = "Error#"
         End If
      Case 14
         varResult = CDec(varValue)
         If Err.Number = 0 Then
            TypeConvert = varResult
         Else
            TypeConvert = "Error#"
         End If
      Case 17
         varResult = CByte(varValue)
         If Err.Number = 0 Then
            TypeConvert = varResult
         Else
            TypeConvert = "Error#"
         End If
      Case 20
         #If Win64 Then
            varResult = Clnglng(varValue)
            If Err.Number = 0 Then
               TypeConvert = varResult
            Else
               TypeConvert = "Error#"
            End If
         #Else
            TypeConvert = "Error#32Bit"
         #End If
   End Select

End Function
