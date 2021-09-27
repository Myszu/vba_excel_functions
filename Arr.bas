Attribute VB_Name = "Arr"
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                                       '''
    ''   Functions:                                                                            ''
    ''                                                                                         ''
    ''   Adds JavaScript-like arrays functions simplyfying working with arrays and modifying   ''
    '' their structure by adding or removing first or last element of given array.             ''
    ''                                                                                         ''
    ''      Arr.Push(Arr, Addon) - adds new entry in the end of an array                       ''
    ''      Arr.Unshift(Arr, Addon) - adds new entry at the beginning of an array              ''
    ''      Arr.Pop(Arr, Addon) - removes last entry of an array                               ''
    ''      Arr.Shift(Arr, Addon) - removes first entry of an array                            ''
    ''                                                                                         ''
    '' Author: Marcin Nowacki <nowacki0508@gmail.com>                                          ''
    '' Copyright: GNU General Public License                                                   ''
    '' Version: 1.0                                                                            ''
    '' Since: 16.08.2021                                                                       ''
    ''                                                                                         ''
    ''                                                                                         ''
    ''   This program is free software; you can redistribute it and/or modify                  ''
    '' it under the terms of the GNU General Public License as published by                    ''
    '' the Free Software Foundation; either version 2 of the License, or                       ''
    '' (at your option) any later version.                                                     ''
    ''                                                                                         ''
    ''   This program is distributed in the hope that it will be useful,                       ''
    '' but WITHOUT ANY WARRANTY; without even the implied warranty of                          ''
    '' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the                           ''
    '' GNU General Public License for more details.                                            ''
    ''                                                                                         ''
    '''                                                                                       '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function Push(Arr As Variant, Addon As Variant) As Variant
Dim NewArr As Variant
    
    On Error Resume Next
    ArrType = Arr(1, 1)
    
    If Err.Number <> 0 Then
    
        ReDim NewArr(UBound(Arr) + 1)
        
    Else
    
        ReDim NewArr(UBound(Arr))
    
    End If
    
    For i = LBound(Arr) To UBound(Arr) + 1

        If Err.Number <> 0 Then
    
            If i <= UBound(Arr) Then
            
                NewArr(i) = Arr(i)
                
            Else
            
                NewArr(i) = Addon
                
            End If
            
        Else
        
            If i <= UBound(Arr) Then
        
                NewArr(i - 1) = Arr(i, 1)
                
            Else
            
                NewArr(i - 1) = Addon
                
            End If
        
        End If
        
    Next i
    
    Push = NewArr
    
End Function

Function Unshift(Arr As Variant, Addon As Variant) As Variant
Dim NewArr As Variant
    
    On Error Resume Next
    ArrType = Arr(1, 1)
    
    If Err.Number <> 0 Then
    
        ReDim NewArr(UBound(Arr) + 1)
        
    Else
    
        ReDim NewArr(UBound(Arr))
    
    End If
    
    For i = LBound(Arr) To UBound(Arr) + 1

        If Err.Number <> 0 Then
    
            If i = LBound(Arr) Then
            
                NewArr(i) = Addon
                
            Else
            
                NewArr(i) = Arr(i - 1)
                
            End If
            
        Else
        
            If i = LBound(Arr) Then
        
                NewArr(i - 1) = Addon
                
            Else
            
                NewArr(i - 1) = Arr(i - 1, 1)
                
            End If
        
        End If
        
    Next i
    
    Unshift = NewArr
    
End Function

Function Pop(Arr As Variant) As Variant
Dim NewArr As Variant
    
    On Error Resume Next
    ArrType = Arr(1, 1)

    If Err.Number <> 0 Then
    
        ReDim NewArr(UBound(Arr) - 1)
        
        For i = LBound(Arr) To UBound(Arr)
        
            NewArr(i) = Arr(i)
                
        Next i
        
    Else
    
        ReDim NewArr(UBound(Arr) - 2)
        
        For i = LBound(Arr) To UBound(Arr) - 1
        
            NewArr(i - 1) = Arr(i, 1)
                
        Next i
    
    End If
    
    Pop = NewArr
    
End Function

Function Shift(Arr As Variant) As Variant
Dim NewArr As Variant
    
    On Error Resume Next
    ArrType = Arr(1, 1)

    If Err.Number <> 0 Then
    
        ReDim NewArr(UBound(Arr) - 1)
        
        For i = LBound(Arr) To UBound(Arr)
        
            NewArr(i) = Arr(i + 1)
                
        Next i
        
    Else
    
        ReDim NewArr(UBound(Arr) - 2)
        
        For i = LBound(Arr) To UBound(Arr) - 1
            
            NewArr(i - 1) = Arr(i + 1, 1)
                
        Next i
    
    End If
    
    Shift = NewArr
    
End Function

Function Rotate(Arr As Variant) As Variant
Dim NewArr As Variant
    
    On Error Resume Next
    ArrType = Arr(1, 1)

    If Err.Number <> 0 Then
    
        ReDim NewArr(UBound(Arr))
        
        For i = LBound(Arr) To UBound(Arr)

            NewArr(i) = Arr((UBound(Arr)) - i)
        
        Next i
        
    Else
    
        ReDim NewArr(UBound(Arr) - 1)
        
        For i = LBound(Arr) To UBound(Arr)

            NewArr(i - 1) = Arr((UBound(Arr) + 1) - i, 1)
        
        Next i
        
    End If
    
    Rotate = NewArr
    
End Function
