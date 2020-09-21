Attribute VB_Name = "STREAK_COUNT"
Function STREAK(rng As Range) As Long
Dim currCell As Range
Dim hit, maxHit As Long


     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                                       '''
    ''   Functions:                                                                            ''
    ''                                                                                         ''
    ''   Function returns maximal streak of not empty cells in given range. Personally I find  ''
    '' it useful in training calendars to rise the difficulty of my training the more often    ''
    '' I fulfill my target. Correct syntax of function should look as below:                   ''
    ''                                                                                         ''
    ''      STREAK(RANGE)                                                                     ''
    ''      STREAK(A1:A10)                                                                     ''
    ''                                                                                         ''
    '' Also be noted, that delimiting sign may vary and in stead of coma you may have to use   ''
    '' semicolon.                                                                              ''
    ''                                                                                         ''
    '' Author: Marcin Nowacki <nowacki0508@gmail.com>                                          ''
    '' Copyright: GNU General Public License                                                   ''
    '' Version: 1.0                                                                            ''
    '' Since: 15.04.2020                                                                       ''
    ''                                                                                         ''
    ''                                                                                         ''
    ''   This program is free software; you can redistribute it and/or modify                  ''
    '' it under the terms of the GNU General Public License as published by                    ''
    '' the Free Software Foundation; either version 2 of the License, or                       ''
    '' (at your option) any later version.                                                     ''
    ''                                                                                         ''
    ''   This program is distributed in the hope that it will be useful,                       ''
    '' but WITHOUT ANY WARRANTY; without even the implied warranty of                          ''
    '' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the                            ''
    '' GNU General Public License for more details.                                            ''
    ''                                                                                         ''
    '''                                                                                       '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    hit = 0
    maxHit = 0
    
    For Each currCell In rng
    
        If currCell.Value <> "" Then
        
            hit = hit + 1
            
            If hit > maxHit Then
            
                maxHit = hit
                
            End If
            
        Else
        
            hit = 0
            
        End If
    
    Next currCell
    
    STREAK = maxHit

End Function
