Attribute VB_Name = "mdlConversion"

Function ConvertCoordinates(latlon As Variant, Optional ByVal OutputFormat As Integer = 0) As Variant
'
' Take a Latitude or Longitude coordinate and convert it to a differnt format
' © 2019 Johan Rodskog
'


    Dim degrees As Double
    Dim minutes As Double
    Dim seconds As Double
    Dim space_location As Integer

    Dim RE As Object, REMatches As Object
    Set RE = CreateObject("vbscript.regexp")
    RE.Global = True
    RE.IgnoreCase = True
    valid = False
    decimal_degrees = 0
    degrees = 0
    minutes = 0
    seconds = 0
    DirectionSign = 0
    DirectionText = 0

    On Error Resume Next
    
    latlon = Trim(latlon)
    'Check if there is a negative sign at the beginnning
    If Left(latlon, 1) = "-" Then DirectionSign = -1
    
    'Get the first NSEW character and use this for direction
    RE.Pattern = "[NSEWnsew]"
    RE.Global = False
    Set allMatches = RE.Execute(latlon)
    If allMatches.Count <> 0 Then
        Select Case UCase(allMatches.Item(0))
            Case "N"
                DirectionText = 1
                NSEW = "Lat"
            Case "S"
                DirectionText = -1
                NSEW = "Lat"
            Case "E"
                DirectionText = 1
                NSEW = "Lon"
            Case "W"
                DirectionText = -1
                NSEW = "Lon"
            Case Else
                DirectionText = 0
        End Select
    End If
    RE.Pattern = "[A-Za-z]"
    RE.Global = True
    latlon = Trim(RE.Replace(latlon, ""))

    ' Replace DMS characters with letters
    latlon = Replace(latlon, Chr(39) & Chr(39), "S") ' ''
    latlon = Replace(latlon, Chr(34), "S")   ' " - Double Prime. This is the most common way to denote Seconds
    latlon = Replace(latlon, Chr(147), "S")  ' “ - Reversed Double Prime
    latlon = Replace(latlon, Chr(148), "S")  ' ”
    latlon = Replace(latlon, Chr(39), "M")   ' ' - Prime. This is the most common way to denote Minutes
    latlon = Replace(latlon, Chr(145), "M")  ' ‘ - Reversed Prime
    latlon = Replace(latlon, Chr(146), "M")  ' ’
    latlon = Replace(latlon, Chr(176), "D")  ' ° - Degrees. This is the most common way to denote Degrees
    latlon = Replace(latlon, Chr(186), "D")  ' º - Masculine Ordinal Indicator. Sometime mistakenly used instead of Unicode 176
    latlon = Trim(latlon)
    
    ' Remove all other characters from string, except '.' (handled below)
    RE.Pattern = "[^0-9. DMS]"
    RE.Global = True
    latlon = Replace(latlon, "-", " ", 1)
    latlon = Trim(Replace(latlon, "  ", " "))
    latlon = Trim(Replace(latlon, "  ", " "))
        
    num_periods = Len(latlon) - Len(Replace(latlon, ".", ""))
    If num_periods > 1 Then
        Temp = Replace(latlon, ".", " ", 1, num_periods - 1)        ' Replace periods with ' ', except for the last one
        Temp = Replace(Temp, "  ", " ")                             ' Replace double spaces with a single space
        Temp = Replace(Temp, "  ", " ")                             ' Replace double spaces with a single space
        num_spaces = Len(Temp) - Len(Replace(Temp, " ", ""))        ' Count number of spaces in the sring
        If num_spaces > 1 Then
            latlon = Replace(latlon, ".", " ", 1, num_periods - 1)
        Else
            latlon = Replace(latlon, ".", " ", 1)
        End If
    End If
    latlon = RE.Replace(latlon, " ")
    latlon = Trim(Replace(latlon, "  ", " "))
    latlon = Trim(Replace(latlon, "  ", " "))
    
    hasD = InStr(1, latlon, "D")
    hasM = InStr(1, latlon, "M")
    hasS = InStr(1, latlon, "S")
    num_spaces = Len(latlon) - Len(Replace(latlon, " ", ""))
    
    If hasD = 0 And num_spaces = 2 Then
        hasD = InStr(1, latlon, " ")
        latlon = Left(latlon, hasD - 1) & "D" & Mid(latlon, hasD + 1, 99)
        If hasM = 0 Then
            hasM = InStr(1, latlon, " ")
            latlon = Left(latlon, hasM - 1) & "M" & Mid(latlon, hasM + 1, 99)
        End If
    ElseIf hasD = 0 And num_spaces = 1 Then
        hasD = InStr(1, latlon, " ")
        latlon = Left(latlon, hasD - 1) & "D" & Mid(latlon, hasD + 1, 99)
    ElseIf hasD > 0 And hasM = 0 And hasS > 0 Then
        latlon = Left(latlon, hasD - 1) & "D00M" & Trim(Mid(latlon, hasD + 1, 99))
        hasD = InStr(1, latlon, "D")
        hasM = InStr(1, latlon, "M")
        hasS = InStr(1, latlon, "S")
    End If
    
    If latlon <> "" Then
        If IsNumeric(latlon) = True Then
            degreesD = latlon
            degrees = Int(latlon)
            minutes = Int((latlon - Int(latlon)) * 60)
            seconds = (((latlon - Int(latlon)) * 60) - (Int((latlon - Int(latlon)) * 60))) * 60
            valid = True
        Else
            degrees = Abs(Val(Left(latlon, hasD)))
            minutes = IIf(hasM > 0, Val(Mid(latlon, hasD + 1, hasM - hasD - 1)), 0)
            seconds = Val(Mid(latlon, hasM + 1, 99))
            latlon = (degrees + (minutes / 60) + (seconds / 3600))
            degreesD = degrees + (minutes / 60) + (seconds / 3600)
            degrees = Int(latlon)
            minutes = Int((latlon - Int(latlon)) * 60)
            seconds = (((latlon - Int(latlon)) * 60) - (Int((latlon - Int(latlon)) * 60))) * 60
            valid = True
        End If
    End If
            
    If valid = True Then
        If DirectionSign = 0 And DirectionText = 0 Then
            DirectionSign = 1
        ElseIf DirectionSign <> 0 And DirectionText = 0 Then
            DirectionSign = DirectionSign
        Else
            DirectionSign = DirectionText
        End If
            
        If NSEW = "Lat" And DirectionSign = 1 Then
            NSEWtext = "N"
        ElseIf NSEW = "Lat" And DirectionSign = -1 Then
            NSEWtext = "S"
        ElseIf NSEW = "Lon" And DirectionSign = 1 Then
            NSEWtext = "E"
        ElseIf NSEW = "Lon" And DirectionSign = -1 Then
            NSEWtext = "W"
        Else
            NSEWtext = ""
        End If
        
        If NSEWtext = "" And DirectionSign = -1 Then
            degrees = -degrees
            degreesD = -degreesD
        ElseIf DirectionSign = -1 Then
            degreesD = -degreesD
        End If
        
        If OutputFormat = 0 Then
            ConvertCoordinates = Round(degreesD, 7)
        ElseIf OutputFormat = 1 And NSEWtext <> "" Then
            ConvertCoordinates = Round(Abs(degreesD), 7) & NSEWtext
        ElseIf OutputFormat = 1 Then
            ConvertCoordinates = Round(degreesD, 7) & NSEWtext
        Else
            ConvertCoordinates = degrees & Chr(176) & " " & Format(minutes, "00") & Chr(39) & " " & Format(seconds, "00.00") & Chr(34) & NSEWtext
        End If
    Else
        ConvertCoordinates = CVErr(xlErrValue)
    End If

End Function




Function DMS2DEC(latlon As Variant, Optional ByVal OutputFormat As Integer = 0) As Variant
'
' Wrapper function to handle both Latitude and Longitude being passed to the conversion function at the same time
' © 2019 Johan Rodskog
'

    On Error GoTo ErrorHandler
    
    DMS2DEC = 0
    
    If InStr(latlon, ",") = 0 Then                  ' Only a single coordinate
        DMS2DEC = ConvertCoordinates(latlon, OutputFormat)
    Else                                            ' Coordinate pair
        lat1 = Left(latlon, InStr(latlon, ",") - 1)
        lon1 = Mid(latlon, InStr(latlon, ",") + 1, 99)
        
        'Convert Latitude and correct for missing Cardinal Directions
            myLat = ConvertCoordinates(lat1, OutputFormat)
            If OutputFormat = 1 Or OutputFormat = -1 Then
                If Not (Right(myLat, 1) = "N" Or Right(myLat, 1) = "S") Then
                    If Left(myLat, 1) = "-" Then
                        myLat = Mid(myLat, 2, Len(myLat) - 1) & "S"
                    Else
                        myLat = myLat & "N"
                    End If
                End If
            End If
            
        'Convert Longitude and correct for missing Cardinal Directions
            myLon = ConvertCoordinates(lon1, OutputFormat)
            If OutputFormat = 1 Or OutputFormat = -1 Then
                If Not (Right(myLon, 1) = "E" Or Right(myLat, 1) = "W") Then
                    If Left(myLon, 1) = "-" Then
                        myLon = Mid(myLon, 2, Len(myLon) - 1) & "W"
                    Else
                        myLon = myLon & "E"
                    End If
                End If
            End If
        
        'Combine the results
            DMS2DEC = myLat & ", " & myLon
    
    End If
    
    Exit Function
    
ErrorHandler:
    DMS2DEC = CVErr(xlErrValue)

End Function



