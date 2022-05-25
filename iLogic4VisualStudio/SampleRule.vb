''' <summary>
'''     Test iLogic rule implementation
''' </summary>
''' <remarks></remarks>
Public Class SampleRule
    Inherits RuleBase 'Basic class which implements basic communication with iLogic

    ''' <summary>
    '''     Rule entry point.
    '''     For iLogic rule you need to copy the Main method content or
    '''     content of whole class. In this case you need to remove
    '''     keywords Public Overrides from method Main()
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides _
        Sub Main()
        'Main iLogic code
        If (ThisDoc.Document.FileSaveCounter > 0) Then
            MsgBox(ThisDoc.FileName(False))
        Else
            MsgBox("File is not saved")
        End If
    End Sub
End Class