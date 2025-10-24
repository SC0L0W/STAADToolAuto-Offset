Option Explicit

Sub Main()
    Call ApplyMemberOffsets
End Sub

Sub ApplyMemberOffsets()
    ' Clear information
    Debug.Clear
    Debug.Print "========================================="
    Debug.Print "MEMBER OFFSET APPLICATION TOOL"
    Debug.Print "========================================="
    Debug.Print ""
    
    ' Define OpenStaad Reference with error handling
    Dim Func As Object
    On Error GoTo ConnectionError
    
    Debug.Print "Connecting to STAAD.Pro..."
    Set Func = GetObject(, "StaadPro.OpenSTAAD")
    Debug.Print "Connection successful!"
    Debug.Print ""
    
    ' Test connection
    Debug.Print "Reading member count..."
    Dim testCount As Variant
    testCount = Func.Geometry.GetMemberCount()
    
    On Error GoTo 0
    
    Dim lMemberCount As Long
    lMemberCount = testCount
    
    Debug.Print "Total Members in Model: "; lMemberCount
    Debug.Print ""
    
    If lMemberCount = 0 Then
        Debug.Print "ERROR: No members found in the model."
        MsgBox "No members found in the model.", vbInformation
        Exit Sub
    End If
    
    ' Get list of all member numbers
    Debug.Print "Getting member list..."
    Dim memberList() As Long
    ReDim memberList(lMemberCount - 1)
    
    On Error Resume Next
    Dim result As Variant
    result = Func.Geometry.GetBeamList(memberList)
    If Err.Number <> 0 Then
        Debug.Print "ERROR: Cannot get member list - "; Err.Description
        Debug.Print "Trying alternative approach - using sequential numbering..."
        Err.Clear
        Dim j As Long
        For j = 0 To lMemberCount - 1
            memberList(j) = j + 1
        Next j
    Else
        Debug.Print "Member list retrieved successfully"
    End If
    On Error GoTo 0
    
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "RETRIEVING MEMBER DATA"
    Debug.Print "========================================="
    Debug.Print ""
    
    Dim i As Long
    Dim memberNum As Long
    Dim startNode As Long, endNode As Long
    Dim startX As Double, startY As Double, startZ As Double
    Dim endX As Double, endY As Double, endZ As Double
    Dim sectionName As String
    Dim length As Double
    
    ' Tolerance for coordinate comparison
    Const coordTolerance As Double = 0.001
    
    Dim beamCount As Long, columnCount As Long, bracingCount As Long
    Dim offsetAppliedCount As Long, offsetFailedCount As Long
    beamCount = 0
    columnCount = 0
    bracingCount = 0
    offsetAppliedCount = 0
    offsetFailedCount = 0
    
    For i = 0 To lMemberCount - 1
        On Error Resume Next
        
        memberNum = memberList(i)
        
        Debug.Print "==================================="
        Debug.Print "Array Index: "; i; " of "; (lMemberCount - 1)
        Debug.Print "Member Number: "; memberNum
        Debug.Print "==================================="
        
        ' Get nodes using GetMemberIncidence
        result = Func.Geometry.GetMemberIncidence(memberNum, startNode, endNode)
        If Err.Number <> 0 Then
            Debug.Print "ERROR: Cannot get member incidence - "; Err.Description
            Err.Clear
            GoTo NextMember
        End If
        
        Debug.Print "Start Node: "; startNode
        Debug.Print "End Node: "; endNode
        
        ' Get node coordinates using GetNodeCoordinates
        Func.Geometry.GetNodeCoordinates startNode, startX, startY, startZ
        If Err.Number <> 0 Then
            Debug.Print "ERROR: Cannot get start node coordinates - "; Err.Description
            Err.Clear
            GoTo NextMember
        End If
        
        Func.Geometry.GetNodeCoordinates endNode, endX, endY, endZ
        If Err.Number <> 0 Then
            Debug.Print "ERROR: Cannot get end node coordinates - "; Err.Description
            Err.Clear
            GoTo NextMember
        End If
        
        Debug.Print ""
        Debug.Print "COORDINATES:"
        Debug.Print "  Start: X="; startX; " Y="; startY; " Z="; startZ
        Debug.Print "  End:   X="; endX; " Y="; endY; " Z="; endZ
        
        ' Get section info
        sectionName = Func.Property.GetMemberSectionLabel(memberNum)
        length = Func.Geometry.GetMemberLength(memberNum)
        If Err.Number <> 0 Then
            Err.Clear
            sectionName = "UNKNOWN"
            length = 0
        End If
        Debug.Print "Section: "; sectionName
        Debug.Print "Length: "; length
        
        ' Get section properties with improved detection
        Dim depth As Double, width As Double
        Dim sectionProps As SectionProperties
        sectionProps = GetSectionProperties(Func, memberNum, sectionName)
        
        depth = sectionProps.depth
        width = sectionProps.width
        
        Debug.Print "Depth: "; depth; " | Width: "; width
        Debug.Print "Section Type: "; sectionProps.sectionType
        
        On Error GoTo 0
        
        ' Calculate ABSOLUTE coordinate changes
        Dim deltaX As Double, deltaY As Double, deltaZ As Double
        deltaX = Abs(endX - startX)
        deltaY = Abs(endY - startY)
        deltaZ = Abs(endZ - startZ)
        
        Debug.Print ""
        Debug.Print "COORDINATE CHANGES (Absolute):"
        Debug.Print "  |DeltaX| = "; deltaX
        Debug.Print "  |DeltaY| = "; deltaY
        Debug.Print "  |DeltaZ| = "; deltaZ
        Debug.Print "  Tolerance = "; coordTolerance
        
        ' Check which coordinates are changing
        Dim isXChanging As Boolean, isYChanging As Boolean, isZChanging As Boolean
        isXChanging = (deltaX > coordTolerance)
        isYChanging = (deltaY > coordTolerance)
        isZChanging = (deltaZ > coordTolerance)
        
        Debug.Print ""
        Debug.Print "COORDINATE CHANGE DETECTION:"
        Debug.Print "  X changing? "; isXChanging; " ("; deltaX; " > "; coordTolerance; ")"
        Debug.Print "  Y changing? "; isYChanging; " ("; deltaY; " > "; coordTolerance; ")"
        Debug.Print "  Z changing? "; isZChanging; " ("; deltaZ; " > "; coordTolerance; ")"
        
        ' Count how many coordinates are changing
        Dim changingCount As Integer
        changingCount = 0
        If isXChanging Then changingCount = changingCount + 1
        If isYChanging Then changingCount = changingCount + 1
        If isZChanging Then changingCount = changingCount + 1
        
        Debug.Print "  Total coordinates changing: "; changingCount
        
        ' Classify member based on coordinate changes
        Dim memberType As String
        Dim offsetStart_X As Double, offsetStart_Y As Double, offsetStart_Z As Double
        Dim offsetEnd_X As Double, offsetEnd_Y As Double, offsetEnd_Z As Double
        
        Debug.Print ""
        Debug.Print "CLASSIFICATION:"
        
        ' Check for COLUMN first (only Y changes)
        If isYChanging And Not isXChanging And Not isZChanging Then
            memberType = "COLUMN"
            columnCount = columnCount + 1
            Debug.Print "  TYPE: COLUMN (Only Y coordinate changes)"
            Debug.Print "  REASON: Y changing, X and Z constant"
            Debug.Print "  ACTION: No offset applied to columns"
            
        ' Check for BEAM (only X or only Z changes, Y constant)
        ElseIf Not isYChanging And (isXChanging Or isZChanging) And Not (isXChanging And isZChanging) Then
            memberType = "BEAM"
            beamCount = beamCount + 1
            
            ' Determine beam orientation (parallel to X or Z axis)
            Dim beamOrientation As String
            If isXChanging Then
                beamOrientation = "X"
                Debug.Print "  TYPE: BEAM (parallel to X-axis)"
            Else
                beamOrientation = "Z"
                Debug.Print "  TYPE: BEAM (parallel to Z-axis)"
            End If
            
            ' Get supporting member information at start and end nodes
            Dim startSupportData As SupportMemberData
            Dim endSupportData As SupportMemberData
            startSupportData = GetSupportingMemberAtNode(Func, startNode, memberNum, memberList, lMemberCount)
            endSupportData = GetSupportingMemberAtNode(Func, endNode, memberNum, memberList, lMemberCount)
            
            Debug.Print "  Start Support: Type="; startSupportData.memberType; " Width="; startSupportData.width; " Tf="; startSupportData.flangeThickness; " SecType="; startSupportData.sectionType
            Debug.Print "  End Support:   Type="; endSupportData.memberType; " Width="; endSupportData.width; " Tf="; endSupportData.flangeThickness; " SecType="; endSupportData.sectionType
            
            ' VERTICAL OFFSET (Y-axis): Always -depth/2 for both start and end
            Dim verticalOffset As Double
            verticalOffset = -depth / 2
            
            ' HORIZONTAL OFFSET (Local X-axis): Based on supporting member type and beam orientation
            Dim horizontalOffsetStart As Double, horizontalOffsetEnd As Double
            
            ' === START NODE OFFSET CALCULATION ===
            If startSupportData.hasSupport Then
                If startSupportData.memberType = "COLUMN" Then
                    ' Connected to COLUMN
                    If beamOrientation = "Z" Then
                        ' Beam parallel to Z-axis: always use depth/2 for any column type
                        horizontalOffsetStart = startSupportData.depth / 2
                        If startSupportData.sectionType = "RECT" Or startSupportData.sectionType = "CIRCLE" Then
                            Debug.Print "  Start: Using depth/2 (RECT/CIR column, beam parallel to Z)"
                        ElseIf IsWideFlangeSteelSection(startSupportData.sectionType) Then
                            Debug.Print "  Start: Using depth/2 (W/S/M/HP/B column, beam parallel to Z)"
                        Else
                            Debug.Print "  Start: Using depth/2 (column, beam parallel to Z)"
                        End If
                    Else
                        ' Beam parallel to X-axis: use width/2
                        horizontalOffsetStart = startSupportData.width / 2
                        Debug.Print "  Start: Using width/2 (column, beam parallel to X)"
                    End If
                ElseIf startSupportData.memberType = "BEAM" Then
                    ' Connected to BEAM
                    ' Check if other end is also connected to beam (beam-to-beam on both ends)
                    If endSupportData.hasSupport And endSupportData.memberType = "BEAM" Then
                        ' Both ends connected to beams: use width/2
                        horizontalOffsetStart = startSupportData.width / 2
                        Debug.Print "  Start: Using width/2 (beam-to-beam, both ends to beams)"
                    Else
                        ' One end to beam, other end to column: NO offset at beam side
                        horizontalOffsetStart = 0
                        Debug.Print "  Start: No offset (connected to beam, other end to column)"
                    End If
                Else
                    ' Default
                    horizontalOffsetStart = startSupportData.width / 2
                    Debug.Print "  Start: Using width/2 (default)"
                End If
            Else
                horizontalOffsetStart = 0
                Debug.Print "  Start: No support - offset = 0"
            End If
            
            ' === END NODE OFFSET CALCULATION ===
            If endSupportData.hasSupport Then
                If endSupportData.memberType = "COLUMN" Then
                    ' Connected to COLUMN
                    If beamOrientation = "Z" Then
                        ' Beam parallel to Z-axis: always use depth/2 for any column type
                        horizontalOffsetEnd = -endSupportData.depth / 2
                        If endSupportData.sectionType = "RECT" Or endSupportData.sectionType = "CIRCLE" Then
                            Debug.Print "  End: Using depth/2 (RECT/CIR column, beam parallel to Z)"
                        ElseIf IsWideFlangeSteelSection(endSupportData.sectionType) Then
                            Debug.Print "  End: Using depth/2 (W/S/M/HP/B column, beam parallel to Z)"
                        Else
                            Debug.Print "  End: Using depth/2 (column, beam parallel to Z)"
                        End If
                    Else
                        ' Beam parallel to X-axis: use width/2
                        horizontalOffsetEnd = -endSupportData.width / 2
                        Debug.Print "  End: Using width/2 (column, beam parallel to X)"
                    End If
                ElseIf endSupportData.memberType = "BEAM" Then
                    ' Connected to BEAM
                    ' Check if other end is also connected to beam (beam-to-beam on both ends)
                    If startSupportData.hasSupport And startSupportData.memberType = "BEAM" Then
                        ' Both ends connected to beams: use width/2
                        horizontalOffsetEnd = -endSupportData.width / 2
                        Debug.Print "  End: Using width/2 (beam-to-beam, both ends to beams)"
                    Else
                        ' One end to beam, other end to column: NO offset at beam side
                        horizontalOffsetEnd = 0
                        Debug.Print "  End: No offset (connected to beam, other end to column)"
                    End If
                Else
                    ' Default
                    horizontalOffsetEnd = -endSupportData.width / 2
                    Debug.Print "  End: Using width/2 (default)"
                End If
            Else
                horizontalOffsetEnd = 0
                Debug.Print "  End: No support - offset = 0"
            End If
            
            ' All offsets in LOCAL coordinates
            offsetStart_X = horizontalOffsetStart
            offsetStart_Y = verticalOffset
            offsetStart_Z = 0
            
            offsetEnd_X = horizontalOffsetEnd
            offsetEnd_Y = verticalOffset
            offsetEnd_Z = 0
            
            Debug.Print ""
            Debug.Print "  OFFSET CALCULATION:"
            Debug.Print "    Horizontal (Local X): Start = "; horizontalOffsetStart
            Debug.Print "                         End = "; horizontalOffsetEnd
            Debug.Print "    Vertical (Local Y):   Both = -"; depth; "/2 = "; verticalOffset
            Debug.Print "    Local Z:              Both = 0"
            
            Debug.Print ""
            Debug.Print "  OFFSETS TO APPLY (Local Coordinates):"
            Debug.Print "    Start: X="; offsetStart_X; " Y="; offsetStart_Y; " Z="; offsetStart_Z
            Debug.Print "    End:   X="; offsetEnd_X; " Y="; offsetEnd_Y; " Z="; offsetEnd_Z
            
            ' Apply offset
            Dim offsetResult As Boolean
            offsetResult = ApplyMemberOffset(Func, memberNum, _
                                  offsetStart_X, offsetStart_Y, offsetStart_Z, _
                                  offsetEnd_X, offsetEnd_Y, offsetEnd_Z)
            
            If offsetResult Then
                offsetAppliedCount = offsetAppliedCount + 1
                Debug.Print "  RESULT: Offset successfully applied!"
            Else
                offsetFailedCount = offsetFailedCount + 1
                Debug.Print "  RESULT: Offset application FAILED!"
            End If
            
        ' Everything else is BRACING
        Else
            memberType = "BRACING"
            bracingCount = bracingCount + 1
            Debug.Print "  TYPE: BRACING (Diagonal - multiple coordinates change)"
            
            If changingCount = 0 Then
                Debug.Print "  REASON: No coordinate changes (zero length?)"
            ElseIf changingCount = 2 Then
                Debug.Print "  REASON: Two coordinates change"
                If isXChanging And isYChanging Then Debug.Print "          (X and Y changing)"
                If isXChanging And isZChanging Then Debug.Print "          (X and Z changing)"
                If isYChanging And isZChanging Then Debug.Print "          (Y and Z changing)"
            ElseIf changingCount = 3 Then
                Debug.Print "  REASON: All three coordinates change"
            End If
            Debug.Print "  ACTION: No offset applied to bracing"
        End If
        
        Debug.Print ""
        
NextMember:
    Next i
    
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "FINAL SUMMARY"
    Debug.Print "========================================="
    Debug.Print "Total Members Processed: "; lMemberCount
    Debug.Print "  BEAMS (Horizontal):    "; beamCount
    Debug.Print "  COLUMNS (Vertical):    "; columnCount
    Debug.Print "  BRACING (Diagonal):    "; bracingCount
    Debug.Print "Offsets Applied:         "; offsetAppliedCount
    Debug.Print "Offsets Failed:          "; offsetFailedCount
    Debug.Print "========================================="
    Debug.Print ""
    
    Dim summaryMsg As String
    summaryMsg = "Member offset processing complete!" & vbCrLf & vbCrLf & _
           "Total Members: " & lMemberCount & vbCrLf & _
           "Beams: " & beamCount & vbCrLf & _
           "Columns: " & columnCount & vbCrLf & _
           "Bracing: " & bracingCount & vbCrLf & vbCrLf & _
           "Offsets Applied: " & offsetAppliedCount
    
    If offsetFailedCount > 0 Then
        summaryMsg = summaryMsg & vbCrLf & "Offsets Failed: " & offsetFailedCount & vbCrLf & vbCrLf & _
                    "Check Debug window for details."
        MsgBox summaryMsg, vbExclamation, "Complete with Warnings"
    Else
        MsgBox summaryMsg, vbInformation, "Complete"
    End If
    
    Exit Sub
    
ConnectionError:
    Debug.Print "========================================="
    Debug.Print "CONNECTION ERROR"
    Debug.Print "========================================="
    Debug.Print "Error Number: "; Err.Number
    Debug.Print "Error Description: "; Err.Description
    Debug.Print "========================================="
    MsgBox "Cannot connect to STAAD.Pro!" & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "Connection Error"
End Sub

' Type to hold section properties
Type SectionProperties
    depth As Double
    width As Double
    sectionType As String
    sectionLongerSideIsX As Boolean
End Type

' Type to hold supporting member data (replaces ColumnData)
Type SupportMemberData
    hasSupport As Boolean
    memberType As String ' "COLUMN" or "BEAM"
    width As Double
    depth As Double
    flangeThickness As Double
    sectionType As String
End Type

Function GetSectionProperties(Func As Object, memberNum As Long, sectionName As String) As SectionProperties
    On Error Resume Next
    
    Dim props As SectionProperties
    props.depth = 0
    props.width = 0

        ' Determine which side is longer
    If props.width > props.depth Then
        ' Longer side is width
        props.sectionLongerSideIsX = True
    Else
        ' Longer side is depth
        props.sectionLongerSideIsX = False
    End If

    props.sectionType = "UNKNOWN"
    
    Debug.Print ""
    Debug.Print "SECTION PROPERTY DETECTION:"
    Debug.Print "  Section Label: "; sectionName
    
    ' Get the full section definition string from STAAD
    Dim sectionDefString As String
    sectionDefString = ""
    Err.Clear
    sectionDefString = Func.Property.GetBeamSectionName(memberNum)
    
    If Err.Number = 0 And Len(sectionDefString) > 0 Then
        Debug.Print "  Section Definition String: "; sectionDefString
    Else
        Debug.Print "  Section Definition String: Not available"
        Err.Clear
    End If
    
    ' Parse section definition to determine type and extract properties
    Dim sectionDefUpper As String
    sectionDefUpper = UCase(Trim(sectionDefString))
    
    ' Check for RECT section (e.g., "Rect 0.40x0.25")
    If InStr(sectionDefUpper, "RECT") > 0 Then
        props.sectionType = "RECT"
        Debug.Print "  Detected as: RECTANGULAR section"
        
        ' Extract dimensions from "Rect 0.40x0.25"
        Dim rectDims As String
        rectDims = ExtractRectDimensions(sectionDefString)
        If Len(rectDims) > 0 Then
            Dim xPos As Long
            xPos = InStr(rectDims, "x")
            If xPos = 0 Then xPos = InStr(rectDims, "X")
            
            If xPos > 0 Then
                Dim depthStr As String, widthStr As String
                depthStr = Trim(Left(rectDims, xPos - 1))
                widthStr = Trim(Mid(rectDims, xPos + 1))
                
                If IsNumeric(depthStr) Then
                    props.depth = CDbl(depthStr)
                    Debug.Print "    Extracted depth: "; props.depth
                End If
                
                If IsNumeric(widthStr) Then
                    props.width = CDbl(widthStr)
                    Debug.Print "    Extracted width: "; props.width
                End If
            End If
        End If
    
    ' Check for CIR section (e.g., "Cir 0.5")
    ElseIf InStr(sectionDefUpper, "CIR") > 0 Then
        props.sectionType = "CIRCLE"
        Debug.Print "  Detected as: CIRCULAR section"
        
        ' Extract diameter from "Cir 0.5"
        Dim cirDim As String
        cirDim = ExtractCircleDiameter(sectionDefString)
        If IsNumeric(cirDim) Then
            props.depth = CDbl(cirDim)
            props.width = CDbl(cirDim)
            Debug.Print "    Extracted diameter: "; props.depth
        End If
    
    ' Detect section type from definition string
    ElseIf InStr(sectionDefUpper, "TABLE ST") > 0 Or InStr(sectionDefUpper, "TABLE D") > 0 Or InStr(sectionDefUpper, "TABLE SD") > 0 Then
        ' Standard Steel Section from Database
        props.sectionType = "STEEL_TABLE"
        Debug.Print "  Detected as: STEEL (from database table)"
        
    ElseIf InStr(sectionDefUpper, "PIPE") > 0 Then
        ' Pipe section
        props.sectionType = "PIPE"
        Debug.Print "  Detected as: PIPE section"
        
    ElseIf InStr(sectionDefUpper, "TUBE") > 0 Then
        ' Tube section
        props.sectionType = "TUBE"
        Debug.Print "  Detected as: TUBE section"
        
    ElseIf InStr(sectionDefUpper, "PRIS") > 0 Then
        ' Prismatic section - extract YD and ZD from definition
        props.sectionType = "PRISMATIC"
        Debug.Print "  Detected as: PRISMATIC section"
        
        ' Try to extract YD and ZD from string like "PRIS YD 1.5 ZD 1.0"
        Dim ydPos As Long, zdPos As Long
        ydPos = InStr(sectionDefUpper, "YD")
        zdPos = InStr(sectionDefUpper, "ZD")
        
        If ydPos > 0 Then
            Dim ydValue As String
            ydValue = ExtractNumberAfterKeyword(sectionDefString, ydPos + 2)
            If IsNumeric(ydValue) Then
                props.depth = CDbl(ydValue)
                Debug.Print "    Extracted YD from definition: "; props.depth
            End If
        End If
        
        If zdPos > 0 Then
            Dim zdValue As String
            zdValue = ExtractNumberAfterKeyword(sectionDefString, zdPos + 2)
            If IsNumeric(zdValue) Then
                props.width = CDbl(zdValue)
                Debug.Print "    Extracted ZD from definition: "; props.width
            End If
        End If
        
    ElseIf InStr(sectionDefUpper, "TAPERED") > 0 Then
        ' Tapered section
        props.sectionType = "TAPERED"
        Debug.Print "  Detected as: TAPERED section"
        
    ElseIf InStr(sectionDefUpper, "UPT") > 0 Then
        ' User Provided Table
        props.sectionType = "USER_TABLE"
        Debug.Print "  Detected as: USER PROVIDED TABLE"
        
    Else
        ' Check if it's a steel section name (e.g., W12X30, C10X20)
        If Len(sectionDefString) > 0 Then
            Dim firstChar As String
            firstChar = UCase(Left(Trim(sectionDefString), 1))
            
            ' Common steel section prefixes
            If firstChar = "W" Or firstChar = "C" Or firstChar = "L" Or firstChar = "M" Or firstChar = "S" Or firstChar = "HP" Or InStr(sectionDefUpper, "HSS") > 0 Then
                props.sectionType = "STEEL_NAME"
                Debug.Print "  Detected as: STEEL section (by name pattern)"
                
                ' Try to lookup in AISC database
                Dim steelProps As SteelSectionData
                steelProps = LookupSteelSection(sectionDefString)
                
                If steelProps.found Then
                    ' Convert from inches to meters (1 inch = 0.0254 meters)
                    props.depth = steelProps.depth * 0.0254
                    props.width = steelProps.width * 0.0254
                    Debug.Print "    Found in AISC database:"
                    Debug.Print "      Depth: "; steelProps.depth; " inches = "; props.depth; " meters"
                    Debug.Print "      Width: "; steelProps.width; " inches = "; props.width; " meters"
                End If
            Else
                props.sectionType = "CONCRETE"
                Debug.Print "  Detected as: CONCRETE/GENERIC section"
            End If
        Else
            ' Try to determine from section label
            Dim nameUpper As String
            nameUpper = UCase(Trim(sectionName))
            
            If InStr(nameUpper, "W") = 1 Or InStr(nameUpper, "I") = 1 Or _
               InStr(nameUpper, "C") = 1 Or InStr(nameUpper, "L") = 1 Or _
               InStr(nameUpper, "HSS") > 0 Then
                props.sectionType = "STEEL"
                Debug.Print "  Detected as: STEEL section (from label)"
            Else
                props.sectionType = "CONCRETE"
                Debug.Print "  Detected as: CONCRETE/GENERIC section"
            End If
        End If
    End If
    
    ' Now try to get properties using GetSectionPropertyValue if not already extracted
    If props.depth <= 0 Or props.width <= 0 Then
        Debug.Print "  Retrieving section properties from STAAD..."
    End If
    
    ' Try to get DEPTH if not already extracted
    If props.depth <= 0 Then
        Debug.Print "  Trying to get DEPTH..."
        
        ' Try "YD" first (most common for depth)
        props.depth = Func.Property.GetSectionPropertyValue(memberNum, "YD")
        If Err.Number = 0 And props.depth > 0 Then
            Debug.Print "    Found YD = "; props.depth
        Else
            Err.Clear
            ' Try "D"
            props.depth = Func.Property.GetSectionPropertyValue(memberNum, "D")
            If Err.Number = 0 And props.depth > 0 Then
                Debug.Print "    Found D = "; props.depth
            Else
                Err.Clear
                ' Try "DEPTH"
                props.depth = Func.Property.GetSectionPropertyValue(memberNum, "DEPTH")
                If Err.Number = 0 And props.depth > 0 Then
                    Debug.Print "    Found DEPTH = "; props.depth
                Else
                    Err.Clear
                    Debug.Print "    Could not retrieve depth property"
                End If
            End If
        End If
    End If
    
    ' Try to get WIDTH if not already extracted
    If props.width <= 0 Then
        Debug.Print "  Trying to get WIDTH..."
        
        ' Try "ZD" first (most common for width)
        props.width = Func.Property.GetSectionPropertyValue(memberNum, "ZD")
        If Err.Number = 0 And props.width > 0 Then
            Debug.Print "    Found ZD = "; props.width
        Else
            Err.Clear
            ' Try "B"
            props.width = Func.Property.GetSectionPropertyValue(memberNum, "B")
            If Err.Number = 0 And props.width > 0 Then
                Debug.Print "    Found B = "; props.width
            Else
                Err.Clear
                ' Try "WIDTH"
                props.width = Func.Property.GetSectionPropertyValue(memberNum, "WIDTH")
                If Err.Number = 0 And props.width > 0 Then
                    Debug.Print "    Found WIDTH = "; props.width
                Else
                    Err.Clear
                    ' Try "BF" (flange width for steel)
                    props.width = Func.Property.GetSectionPropertyValue(memberNum, "BF")
                    If Err.Number = 0 And props.width > 0 Then
                        Debug.Print "    Found BF = "; props.width
                    Else
                        Err.Clear
                        Debug.Print "    Could not retrieve width property"
                    End If
                End If
            End If
        End If
    End If
    
    ' After assigning final depth and width
    If props.width > props.depth Then
        props.sectionLongerSideIsX = True
    Else
        props.sectionLongerSideIsX = False
    End If

    ' Apply default values if still not found
    If props.depth <= 0 Then
        props.depth = 0.5
        Debug.Print "  Using default depth = "; props.depth
    End If
    
    If props.width <= 0 Then
        props.width = 0.3
        Debug.Print "  Using default width = "; props.width
    End If
    
    Debug.Print "  Final Properties: Depth="; props.depth; " Width="; props.width
    
    GetSectionProperties = props
    
    On Error GoTo 0
End Function

' Function to get supporting member (column or beam) at a node
Function GetSupportingMemberAtNode(Func As Object, nodeNum As Long, excludeMember As Long, memberList() As Long, memberCount As Long) As SupportMemberData
    On Error Resume Next
    
    Dim i As Long
    Dim memNum As Long
    Dim startNode As Long, endNode As Long
    Dim startX As Double, startY As Double, startZ As Double
    Dim endX As Double, endY As Double, endZ As Double
    Dim deltaX As Double, deltaY As Double, deltaZ As Double
    Dim sectionName As String
    Dim sectionProps As SectionProperties
    Dim result As SupportMemberData
    
    ' Initialize default values
    result.hasSupport = False
    result.memberType = ""
    result.width = 0.3
    result.flangeThickness = 0
    result.sectionType = ""
    
    Debug.Print "    [Searching for supporting member at node "; nodeNum; "]"
    
    Const coordTolerance As Double = 0.001
    
    For i = 0 To memberCount - 1
        memNum = memberList(i)
        If memNum = excludeMember Then GoTo NextMem
        
        Func.Geometry.GetMemberIncidence memNum, startNode, endNode
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextMem
        End If
        
        If startNode = nodeNum Or endNode = nodeNum Then
            Func.Geometry.GetNodeCoordinates startNode, startX, startY, startZ
            If Err.Number <> 0 Then
                Err.Clear
                GoTo NextMem
            End If
            
            Func.Geometry.GetNodeCoordinates endNode, endX, endY, endZ
            If Err.Number <> 0 Then
                Err.Clear
                GoTo NextMem
            End If
            
            deltaX = Abs(endX - startX)
            deltaY = Abs(endY - startY)
            deltaZ = Abs(endZ - startZ)
            
            ' Check if this is a COLUMN (only Y changes)
            If deltaY > coordTolerance And deltaX < coordTolerance And deltaZ < coordTolerance Then
                Debug.Print "    [Found COLUMN member "; memNum; " - deltaY="; deltaY; "]"
                
                result.memberType = "COLUMN"
                result.hasSupport = True
                
                ' Get section name
                sectionName = Func.Property.GetMemberSectionLabel(memNum)
                If Err.Number <> 0 Then
                    Err.Clear
                    sectionName = "UNKNOWN"
                End If
                
                ' Get section properties
                sectionProps = GetSectionProperties(Func, memNum, sectionName)
                
                result.width = sectionProps.width
                result.depth = sectionProps.depth
                result.sectionType = sectionProps.sectionType
                
                ' Get flange thickness for steel sections
                If sectionProps.sectionType = "STEEL_NAME" Then
                    Dim sectionDefString As String
                    sectionDefString = Func.Property.GetBeamSectionName(memNum)
                    If Err.Number <> 0 Then
                        Err.Clear
                        sectionDefString = sectionName
                    End If
                    
                    Dim tfData As Double
                    tfData = GetFlangThicknessFromDatabase(sectionDefString)
                    result.flangeThickness = tfData
                    Debug.Print "    [Column: Width="; result.width; " Tf="; result.flangeThickness; " Type="; result.sectionType; "]"
                Else
                    Debug.Print "    [Column: Width="; result.width; " Type="; result.sectionType; "]"
                End If
                
                GetSupportingMemberAtNode = result
                Exit Function
                
            ' Check if this is a BEAM (only X or only Z changes, Y constant)
            ElseIf Not (deltaY > coordTolerance) And ((deltaX > coordTolerance) Or (deltaZ > coordTolerance)) And Not ((deltaX > coordTolerance) And (deltaZ > coordTolerance)) Then
                Debug.Print "    [Found BEAM member "; memNum; " - deltaX="; deltaX; " deltaZ="; deltaZ; "]"
                
                result.memberType = "BEAM"
                result.hasSupport = True
                
                ' Get section name
                sectionName = Func.Property.GetMemberSectionLabel(memNum)
                If Err.Number <> 0 Then
                    Err.Clear
                    sectionName = "UNKNOWN"
                End If
                
                ' Get section properties
                sectionProps = GetSectionProperties(Func, memNum, sectionName)
                
                result.width = sectionProps.width
                result.depth = sectionProps.depth
                result.sectionType = sectionProps.sectionType
                result.flangeThickness = 0 ' Beams don't use flange thickness for offsets
                
                Debug.Print "    [Beam: Width="; result.width; " Type="; result.sectionType; "]"
                
                GetSupportingMemberAtNode = result
                Exit Function
            End If
        End If
NextMem:
    Next i
    
    Debug.Print "    [No supporting member found - using defaults]"
    GetSupportingMemberAtNode = result
    
End Function

' Function to check if section is W, S, M, HP, or B shape
Function IsWideFlangeSteelSection(sectionType As String) As Boolean
    Dim typeUpper As String
    typeUpper = UCase(Trim(sectionType))
    
    IsWideFlangeSteelSection = (typeUpper = "STEEL_NAME")
End Function

' Function to get flange thickness from database
Function GetFlangThicknessFromDatabase(sectionName As String) As Double
    On Error Resume Next
    
    GetFlangThicknessFromDatabase = 0
    
    Dim scriptPath As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    scriptPath = CurDir
    Dim dbPath As String
    dbPath = scriptPath & "\AISC_STEEL_DATABASE.xlsx"
    
    If Not fso.FileExists(dbPath) Then
        Set fso = Nothing
        Exit Function
    End If
    
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    
    Set xlBook = xlApp.Workbooks.Open(dbPath)
    Set xlSheet = xlBook.Worksheets(1)
    
    Dim lastRow As Long
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 2).End(-4162).Row
    
    Dim i As Long
    Dim cellValue As String
    Dim searchName As String
    searchName = Trim(UCase(sectionName))
    
    For i = 1 To lastRow
        cellValue = Trim(UCase(CStr(xlSheet.Cells(i, 2).Value)))
        
        If cellValue = searchName Then
            ' Get flange thickness from column 8, convert from inches to meters
            Dim tfValue As Variant
            tfValue = xlSheet.Cells(i, 8).Value
            
            If Not IsEmpty(tfValue) And Not IsNull(tfValue) And IsNumeric(tfValue) Then
                GetFlangThicknessFromDatabase = CDbl(tfValue) * 0.0254 ' Convert inches to meters
                Debug.Print "      [Flange thickness from DB: "; CDbl(tfValue); " inches = "; GetFlangThicknessFromDatabase; " meters]"
            End If
            Exit For
        End If
    Next i
    
    xlBook.Close False
    xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    Set fso = Nothing
    
    On Error GoTo 0
End Function

Function ApplyMemberOffset(Func As Object, memberNum As Long, _
                          startX As Double, startY As Double, startZ As Double, _
                          endX As Double, endY As Double, endZ As Double) As Boolean
    On Error Resume Next
    
    ApplyMemberOffset = False
    
    Debug.Print ""
    Debug.Print "  APPLYING MEMBER OFFSET:"
    Debug.Print "    Member: "; memberNum
    Debug.Print "    Start: X="; startX; " Y="; startY; " Z="; startZ
    Debug.Print "    End:   X="; endX; " Y="; endY; " Z="; endZ
    
    ' Clear any existing errors
    Err.Clear
    
    ' Create START offset specification (Location=0, Local=1)
    Dim startSpecID As Long
    startSpecID = Func.Property.CreateMemberOffsetSpec(0, 1, startX, startY, startZ)
    
    If Err.Number <> 0 Then
        Debug.Print "    ERROR creating START spec: "; Err.Description; " ("; Err.Number; ")"
        Err.Clear
        Exit Function
    End If
    
    If startSpecID <= 0 Then
        Debug.Print "    ERROR: Invalid START spec ID: "; startSpecID
        Exit Function
    End If
    
    Debug.Print "    START Spec ID created: "; startSpecID
    
    ' Create END offset specification (Location=1, Local=1)
    Dim endSpecID As Long
    endSpecID = Func.Property.CreateMemberOffsetSpec(1, 1, endX, endY, endZ)
    
    If Err.Number <> 0 Then
        Debug.Print "    ERROR creating END spec: "; Err.Description; " ("; Err.Number; ")"
        Err.Clear
        Exit Function
    End If
    
    If endSpecID <= 0 Then
        Debug.Print "    ERROR: Invalid END spec ID: "; endSpecID
        Exit Function
    End If
    
    Debug.Print "    END Spec ID created: "; endSpecID
    
    ' Create array with single member number
    Dim beamNumbers(0 To 0) As Long
    beamNumbers(0) = memberNum
    
    ' Assign START specification
    Debug.Print "    Assigning START spec to member..."
    Err.Clear
    
    Dim retVal As Long
    retVal = Func.Property.AssignMemberSpecToBeam(beamNumbers, startSpecID)
    
    If Err.Number <> 0 Then
        Debug.Print "    ERROR assigning START: "; Err.Description; " (Code: "; Err.Number; ")"
        Debug.Print "    Return value: "; retVal
        Err.Clear
        Exit Function
    End If
    
    Debug.Print "    START assigned (Return: "; retVal; ")"
    
    ' Assign END specification
    Debug.Print "    Assigning END spec to member..."
    Err.Clear
    
    retVal = Func.Property.AssignMemberSpecToBeam(beamNumbers, endSpecID)
    
    If Err.Number <> 0 Then
        Debug.Print "    ERROR assigning END: "; Err.Description; " (Code: "; Err.Number; ")"
        Debug.Print "    Return value: "; retVal
        Err.Clear
        Exit Function
    End If
    
    Debug.Print "    END assigned (Return: "; retVal; ")"
    
    ApplyMemberOffset = True
    Debug.Print "    SUCCESS: Offsets applied to member "; memberNum
    
    On Error GoTo 0
End Function

' Helper function to extract numeric value after a keyword in a string
Function ExtractNumberAfterKeyword(sourceString As String, startPos As Long) As String
    Dim i As Long
    Dim char As String
    Dim numStr As String
    Dim foundDigit As Boolean
    
    numStr = ""
    foundDigit = False
    
    ' Skip spaces after keyword
    For i = startPos To Len(sourceString)
        char = Mid(sourceString, i, 1)
        If char <> " " Then
            Exit For
        End If
    Next i
    
    ' Extract number (including decimal point)
    For i = i To Len(sourceString)
        char = Mid(sourceString, i, 1)
        If (char >= "0" And char <= "9") Or char = "." Or char = "-" Then
            numStr = numStr & char
            foundDigit = True
        ElseIf foundDigit Then
            ' Stop at first non-numeric character after finding digits
            Exit For
        End If
    Next i
    
    ExtractNumberAfterKeyword = numStr
End Function

' Helper function to extract dimensions from Rect definition
Function ExtractRectDimensions(sectionString As String) As String
    Dim startPos As Long
    Dim i As Long
    Dim char As String
    Dim dimStr As String
    
    ' Find "Rect" and skip to dimensions
    startPos = InStr(UCase(sectionString), "RECT")
    If startPos = 0 Then
        ExtractRectDimensions = ""
        Exit Function
    End If
    
    startPos = startPos + 4 ' Skip "RECT"
    
    ' Skip spaces
    For i = startPos To Len(sectionString)
        char = Mid(sectionString, i, 1)
        If char <> " " Then
            startPos = i
            Exit For
        End If
    Next i
    
    ' Extract dimension string (e.g., "0.40x0.25")
    dimStr = ""
    For i = startPos To Len(sectionString)
        char = Mid(sectionString, i, 1)
        If (char >= "0" And char <= "9") Or char = "." Or char = "x" Or char = "X" Then
            dimStr = dimStr & char
        ElseIf Len(dimStr) > 0 Then
            Exit For
        End If
    Next i
    
    ExtractRectDimensions = dimStr
End Function

' Helper function to extract diameter from Cir definition
Function ExtractCircleDiameter(sectionString As String) As String
    Dim startPos As Long
    Dim i As Long
    Dim char As String
    Dim dimStr As String
    
    ' Find "Cir" and skip to diameter
    startPos = InStr(UCase(sectionString), "CIR")
    If startPos = 0 Then
        ExtractCircleDiameter = ""
        Exit Function
    End If
    
    startPos = startPos + 3 ' Skip "CIR"
    
    ' Skip spaces
    For i = startPos To Len(sectionString)
        char = Mid(sectionString, i, 1)
        If char <> " " Then
            startPos = i
            Exit For
        End If
    Next i
    
    ' Extract diameter value
    dimStr = ""
    For i = startPos To Len(sectionString)
        char = Mid(sectionString, i, 1)
        If (char >= "0" And char <= "9") Or char = "." Then
            dimStr = dimStr & char
        ElseIf Len(dimStr) > 0 Then
            Exit For
        End If
    Next i
    
    ExtractCircleDiameter = dimStr
End Function

' Type to hold steel section data from database
Type SteelSectionData
    found As Boolean
    depth As Double
    width As Double
End Type

' Function to lookup steel section in AISC database
Function LookupSteelSection(sectionName As String) As SteelSectionData
    On Error Resume Next
    
    Dim result As SteelSectionData
    result.found = False
    result.depth = 0
    result.width = 0
    
    ' Get the directory where the VBS script is located
    Dim scriptPath As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Use current directory as script path
    scriptPath = CurDir
    
    Dim dbPath As String
    dbPath = scriptPath & "\AISC_STEEL_DATABASE.xlsx"
    
    Debug.Print "    Script directory: "; scriptPath
    Debug.Print "    Looking up section in database: "; dbPath
    
    On Error Resume Next
    
    ' Check if file exists
    If Not fso.FileExists(dbPath) Then
        Debug.Print "    Database file not found: "; dbPath
        Set fso = Nothing
        Exit Function
    End If
    
    ' Open Excel file
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    
    Set xlBook = xlApp.Workbooks.Open(dbPath)
    Set xlSheet = xlBook.Worksheets(1) ' First worksheet
    
    ' Find the section name in column 2
    Dim lastRow As Long
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 2).End(-4162).Row ' xlUp = -4162
    
    Dim i As Long
    Dim cellValue As String
    Dim searchName As String
    Dim widthValue As Variant
    searchName = Trim(UCase(sectionName))
    
    Debug.Print "    Searching for: "; searchName
    Debug.Print "    Rows to search: 1 to "; lastRow
    
    For i = 1 To lastRow
        cellValue = Trim(UCase(CStr(xlSheet.Cells(i, 2).Value)))
        
        If cellValue = searchName Then
            ' Found! Get depth from column 5
            result.depth = CDbl(xlSheet.Cells(i, 5).Value)
            
            ' Get width/base from column 6
            widthValue = xlSheet.Cells(i, 6).Value
            
            ' Check if column 6 is empty or zero
            If IsEmpty(widthValue) Or IsNull(widthValue) Or Trim(CStr(widthValue)) = "" Or CDbl(widthValue) = 0 Then
                ' If empty, use depth as width
                result.width = result.depth
                Debug.Print "    Found at row "; i; ": Depth="; result.depth; " inches, Width="; result.width; " inches (width empty, using depth)"
            Else
                result.width = CDbl(widthValue)
                Debug.Print "    Found at row "; i; ": Depth="; result.depth; " inches, Width="; result.width; " inches"
            End If
            
            result.found = True
            Exit For
        End If
    Next i
    
    If Not result.found Then
        Debug.Print "    Section not found in database"
    End If
    
    ' Clean up
    xlBook.Close False
    xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    Set fso = Nothing
    
    LookupSteelSection = result
    
    On Error GoTo 0
End Function