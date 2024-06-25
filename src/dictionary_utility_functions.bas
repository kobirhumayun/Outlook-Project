Attribute VB_Name = "dictionary_utility_functions"
Option Explicit

Private Function AddKeysAndValueSame(dictionary As Object, keysArray As Variant) As Object
        
    Dim removedAllInvalidChrFromKeys As Variant
        
    Dim i As Long
    ' Add same keys and value
    For i = LBound(keysArray) To UBound(keysArray)
        removedAllInvalidChrFromKeys = Application.Run("outlook_utility_functions.RemoveInvalidChars", keysArray(i))   'remove all invalid characters for use dic keys
        dictionary(removedAllInvalidChrFromKeys) = keysArray(i)
    Next i

    ' Return the modified dictionary
    Set AddKeysAndValueSame = dictionary
End Function

 Private Function PutDictionaryValuesIntoWorksheet(wsRange As Range, dict As Object, keysPrint As Boolean, itemsPrint As Boolean, printOnColumn As Boolean)
    ' wsRange is just starting one cell address, the function dynamically resizes the range
    
    If dict.Count > 0 Then
    
        If (keysPrint And itemsPrint And printOnColumn) Then
    
            wsRange.Resize(dict.Count, 1).Value = Application.Run("outlook_utility_functions.oneDArrayConvertToTwoDArray", dict.Keys)
            wsRange.Offset(0, 1).Resize(dict.Count, 1).Value = Application.Run("outlook_utility_functions.oneDArrayConvertToTwoDArray", dict.Items)
            
        ElseIf (keysPrint And printOnColumn) Then
    
            wsRange.Resize(dict.Count, 1).Value = Application.Run("outlook_utility_functions.oneDArrayConvertToTwoDArray", dict.Keys)
    
        ElseIf (itemsPrint And printOnColumn) Then
    
            wsRange.Resize(dict.Count, 1).Value = Application.Run("outlook_utility_functions.oneDArrayConvertToTwoDArray", dict.Items)
    
        ElseIf (keysPrint And itemsPrint) Then
    
            wsRange.Resize(1, dict.Count).Value = dict.Keys
            wsRange.Offset(1, 0).Resize(1, dict.Count).Value = dict.Items
    
        ElseIf (keysPrint) Then
    
            wsRange.Resize(1, dict.Count).Value = dict.Keys
    
        ElseIf (itemsPrint) Then
    
            wsRange.Resize(1, dict.Count).Value = dict.Items
    
        End If
    
    End If


End Function

