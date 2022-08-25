Attribute VB_Name = "Utils"
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                                       '''
    ''   Functions:                                                                            ''
    ''                                                                                         ''
    ''   Function generates progress bar visible in status bar. Prefferably use from           ''
    '' your VBA code by calling function wiith CURRENT as current iteration number,            ''
    '' MAX as number of steps/iterations and optionally STEP as text discription of current    ''
    '' iteration. After all iterations will pass put below line in the end of your code:       ''
    ''                                                                                         ''
    ''      Application.Statusbar = False                                                      ''
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
    

Function Progress(current, max, Optional step)
Dim done, all, task As String
    
    If step <> "" Then
        
        task = current & "/" & max & " || " & step
    
    Else
    
        task = current & "/" & max
        
    End If
    
    If max <= 25 Then
    
        For i = 1 To current
            
            done = done & "|"
            
        Next i
        
        For i = 1 To (max - current)
        
            all = all & "-"
            
        Next i
        
        Progress = "[ " & done & all & " ]" & task
        
    Else
        
        current = CInt((current / max) * 25)
        max = CInt((max / max) * 25)
        
        For i = 1 To current
            
            done = done & "|"
            
        Next i
        
        For i = 1 To (max - current)
        
            all = all & "-"
            
        Next i
        
        Progress = "[ " & done & all & " ]" & task
    
    End If
    
    Application.StatusBar = Progress

End Function

