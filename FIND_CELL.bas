Attribute VB_Name = "FIND_CELL"
Function CELL_ADDRESS(lookup_value As Variant, lookup_range As Range, value_number As Long, no_value As Variant, value_type As Long) As Variant
Dim currCell As Range
Dim currVal As Long
Dim found As Variant


     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                                       '''
    ''   Functions:                                                                            ''
    ''                                                                                         ''
    ''   Function looks up for a given value and returns address of cell matching it. You can  ''
    '' set a number of returned value in case there are multiple matches. You can also write a ''
    '' custom message if there is no value found. Finally you may also change the type of      ''
    '' returned address from plain to fully absolute. Correct syntax of function should        ''
    '' look as below:                                                                          ''
    ''                                                                                         ''
    ''      =CELL_ADDRESS(LOOKUP VALUE,LOOKUP RANGE,VALUE NUMBER,NO VALUE,VALUE TYPE)          ''
    ''      =CELL_ADDRESS("Anna",$A$1:$E$10,1,"NOTHING FOUND",0)                               ''
    ''                                                                                         ''
    '' Also be noted, that delimiting sign may vary and in stead of coma you may have to use   ''
    '' semicolon.                                                                              ''
    ''                                                                                         ''
    ''                                                                                         ''
    '' Author: Marcin Nowacki <nowacki0508@gmail.com>                                          ''
    '' Copyright: GNU General Public License                                                   ''
    '' Version: 1.0                                                                            ''
    '' Since: 11.08.2020                                                                       ''
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

    currVal = 0

    For Each currCell In lookup_range
    
        If currCell.Value = lookup_value Then
        
            currVal = currVal + 1
            
            If value_number = currVal Then
                
                Select Case value_type
                
                    Case 0
                    
                        found = currCell.Address(0, 0)
                        Exit For
                    
                    Case 1
                    
                        found = currCell.Address(1, 0)
                        Exit For
                        
                    Case 2
                    
                        found = currCell.Address(0, 1)
                        Exit For
                        
                    Case 3
                    
                        found = currCell.Address(1, 1)
                        Exit For
                        
                    Case ""
                
                End Select
            
            End If
        
        End If
    
    Next currCell
    
    If found <> "" Then
    
        CELL_ADDRESS = found
        
    Else
    
        CELL_ADDRESS = no_value
        
    End If

End Function
