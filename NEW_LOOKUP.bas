Attribute VB_Name = "NEW_LOOKUP"
Function XLOOKUP(lookup_value As Variant, lookup_range As Range, offset_column As Long, value_number As Long, no_value As Variant) As Variant
Dim currCell As Range
Dim currValue As Long
Dim found As Variant


     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                                       '''
    ''   Functions:                                                                            ''
    ''                                                                                         ''
    ''   Function works similar to VLOOKUP, but have some tweaks, that VLOOKUP lacks. It       ''
    '' searches every cell in given range, not only first column. If it finds looked up value, ''
    '' it may give back value from offset cell both to the right or to the left. Also if       ''
    '' multiple cells contains looked up value, you may choose which instance do you want to   ''
    '' return. Finally - you may also set the message to return if there is non values found.  ''
    '' Correct syntax of function should look as below:                                        ''
    ''                                                                                         ''
    ''      =XLOOKUP(LOOKUP VALUE,LOOKUP RANGE,OFFSET COLUMN,VALUE NUMBER,NO VALUE)            ''
    ''      =XLOOKUP("Anna",$A$1:$E$10,-1,1,"NOTHING FOUND")                                   ''
    ''                                                                                         ''
    '' Also be noted, that delimiting sign may vary and in stead of coma you may have to use   ''
    '' semicolon.                                                                              ''
    ''                                                                                         ''
    ''                                                                                         ''
    '' Author: Marcin Nowacki <nowacki0508@gmail.com>                                          ''
    '' Copyright: GNU General Public License                                                   ''
    '' Version: 1.0                                                                            ''
    '' Since: 21.09.2020                                                                       ''
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

    currValue = 0

    For Each currCell In lookup_range
    
        If currCell.Value = lookup_value Then
        
            currValue = currValue + 1
            
            If value_number = currValue Then
            
                found = currCell.Offset(0, offset_column)
                
                Exit For
                
            End If
        
        End If
    
    Next currCell
    
    If found <> "" Then
        
        XLOOKUP = found
    
    Else
    
        XLOOKUP = no_value
    
    End If

End Function
