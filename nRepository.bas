Attribute VB_Name = "nRepository"
'*************************************************************************************************************************************************************************************************************************************************
'            COPYRIGHT NOTICE
'
' Copyright (C) David Briant 2009 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************

Option Explicit

Private myRepository As New Dictionary

Function Repository() As Dictionary
    Set Repository = myRepository
End Function

Sub ReleaseRepository()
    Set myRepository = Nothing
End Sub
