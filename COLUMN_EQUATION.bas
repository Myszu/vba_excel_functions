Attribute VB_Name = "COLUMN_EQUATION"
Function COL_EQU(Range1 As Range, Range2 As Range, equType As String) As Variant
Dim currCell As Range
Dim colDifference, rowDifference As Long
Dim sum As Double

    
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                                       '''
    ''   Functions:                                                                            ''
    ''                                                                                         ''
    ''   Function multiplies, divides or subtracts two values from each columns row and then   ''
    '' sums them. Personally I find it useful to calculate sum of products value, while one    ''
    '' column contains units and second contains prices. Correct syntax of function should     ''
    '' look as below:                                                                          ''
    ''                                                                                         ''
    ''      COL_EQU(RANGE1,RANGE2,EQUATION)                                                       ''
    ''      COL_EQU(A1:A4,B1:B4,"MULTI")                                                       ''
    ''                                                                                         ''
    '' Given example returns value of MULTI(plication) but you can also use DIV(ision) or      ''
    '' SUB(traction). Also be noted, that delimiting sign may vary and in stead of coma you    ''
    '' may have to use semicolon.                                                              ''
    ''                                                                                         ''
    ''                                                                                         ''
    '' Author: Marcin Nowacki <nowacki0508@gmail.com>                                          ''
    '' Copyright: GNU General Public License                                                   ''
    '' Version: 1.0                                                                            ''
    '' Since: 18.07.2020                                                                       ''
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
    
    sum = 0
    
    colDifference = Range2.Column - Range1.Column
    rowDifference = Range2.Row - Range1.Row
    
    If equType <> "" Then
    
        If Range1.Count = Range2.Count Then
        
            For Each currCell In Range1
            
                If equType = "MULTI" Then
                
                    sum = sum + (currCell.Value * currCell.Offset(rowDifference, colDifference).Value)
                
                End If
            
                If equType = "DIV" Then
                
                    sum = sum + (currCell.Value / currCell.Offset(rowDifference, colDifference).Value)
                
                End If
                
                If equType = "SUB" Then
                
                    sum = sum + (currCell.Value - currCell.Offset(rowDifference, colDifference).Value)
                
                End If
            
            Next currCell
            
            COL_EQU = sum
            
        Else
        
            COL_EQU = "NOT EQUAL COLUMNS"
        
        End If
        
    Else
        
        COL_EQU = "EQUATION TYPE NOT SET"
        
    End If

End Function
