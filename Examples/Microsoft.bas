Public Sub PowerPoint()

'On Error GoTo fehler
If Not checkFile(getPath & "test.pptx") Then

    MsgBox ("PowerPoint not found. Please put the template into " & getPath & "test.pptx")
    Exit Sub
End If

Dim ppObj As PowerPoint.Application
Dim ppPres As PowerPoint.Presentation
Dim ppSlide As PowerPoint.Slide
Dim ppTextframe As PowerPoint.Shape
Dim oPPTShape As PowerPoint.Shape
Dim assets As Variant
Dim Location, termin As String
Dim Tag, Monat, zeilen As Integer

Dim AssetName, assetpicture As String
Dim iSlide, iAmount, i, iAsset As Integer

Set ppObj = CreateObject("PowerPoint.Application")
ppObj.Visible = msoTrue

'----------------------------

'open File
Set ppPres = ppObj.Presentations.Open(getPath & "test.pptx")

'Get Asset Amount
assets = sql_to_array("Select * from Assets where FK_AssetGroup = " & Forms("Cover").Controls("listAssetGroups").column(0))
iAmount = UBound(assets, 2) 'zukünftig variabel

'Get AssetType Categories and Amount
Dim assetInformation As Variant
assetInformation = sql_to_array("Select Type, count(Number) from Assets where FK_AssetGroup=" & Forms("Cover").Controls("listAssetGroups").column(0) & " group by Type")

'Get assetGroup information (Location, SaleDue, Adress)
Dim assetGroup As Variant
assetGroup = sql_to_array("select AssetGroups.location, SaleDue, Adress from AssetGroups left join location on AssetGroups.location=Location.location where Number=" & Forms("Cover").Controls("listAssetGroups").column(0))


'Month entwickeln (Tag im Monat <15; Vormonat sonst tatsächlicher Monat)

If Not IsNull(assetGroup(1, 0)) Then
    If Day(assetGroup(1, 0)) < 15 Then
            Monat = Month(assetGroup(1, 0))
            Else: Monat = Month(assetGroup(1, 0)) + 1
    End If
    
    termin = Format(DateSerial(Year(assetGroup(1, 0)), Monat, 1), "MMMM") & " " & Year(assetGroup(1, 0))
End If
'Location entwickeln
Location = assetGroup(0, 0)

pic = sql_to_array("Select pictures from location where Location ='" & Location & "'")

If InStr(1, Location, "_", vbTextCompare) > 0 Then
    Location = Left(loaction, InStr(1, Location, "_", vbTextCompare) - 1)
End If
    

'AssetAdress
adress = assetGroup(2, 0)

'Zeilen der Tabelle
    zeilen = UBound(assetInformation, 2) + 1

'Startseite
        'Header
        iSlide = 1
        
        With ppPres.Slides(iSlide)
                .Shapes("Textfeld 9").TextEffect.text = "Ford Plant " & Location              '''can be different retangle
                .Shapes("Rectangle 4").TextEffect.text = "Upcoming Private Sale " & termin      '''can be different retangle
                
                .Shapes.AddTable zeilen, 2, 70, 500, 400, 200
                .Shapes("Table 5").table.Rows(1).Height = 40
                .Shapes("Table 5").table.Columns(1).Width = 300
                .Shapes("Table 5").table.Columns(2).Width = 100
                .Shapes("Table 5").table.Cell(1, 1).Shape.TextFrame2.TextRange.text = "Asset Type"
                .Shapes("Table 5").table.Cell(1, 1).Shape.TextFrame2.VerticalAnchor = msoAnchorMiddle
                .Shapes("Table 5").table.Cell(1, 2).Shape.TextFrame2.TextRange.text = "Amount"
                .Shapes("Table 5").table.Cell(1, 2).Shape.TextFrame2.VerticalAnchor = msoAnchorMiddle
                
                
                For y = 1 To UBound(assetInformation, 2)
                
                    .Shapes("Table 5").table.Cell(y + 1, 1).Shape.TextFrame2.TextRange.text = assetInformation(0, y)
                    .Shapes("Table 5").table.Cell(y + 1, 2).Shape.TextFrame2.TextRange.text = assetInformation(1, y)
                    .Shapes("Table 5").table.Rows(y + 1).Height = 20
                    .Shapes("Table 5").table.Cell(y + 1, 1).Shape.TextFrame2.VerticalAnchor = msoAnchorMiddle
                    .Shapes("Table 5").table.Cell(y + 1, 2).Shape.TextFrame2.VerticalAnchor = msoAnchorMiddle
                
                Next y
                
                
              'Adress eintragen
              
               .Shapes("Footer Placeholder 2").TextFrame2.TextRange.text = "Assets Location:" & vbCrLf & assetGroup(2, 0)
               .Shapes("Footer Placeholder 2").TextFrame2.HorizontalAnchor = msoAnchorCenter
                        
                        If Not IsEmpty(pic) And Not IsNull(pic(0, 0)) Then
                            Call ppPres.Slides(iSlide).Shapes.AddPicture(pic(0, 0), msoFalse, msoCTrue, 120, 200, 300, 200)
                        End If
        End With

  


    

    

'Start Asset

iAsset = 0
Do Until iAsset > iAmount
        iSlide = iAsset + 2
        'If new slide is needed added a new slide
        If iSlide > 1 Then
            Call ppPres.Slides.AddSlide(iSlide, ppPres.Designs(1).SlideMaster.CustomLayouts(1))
        End If
        
        'Define which Slide is the current assset
       
        
        
        'Header
        ppPres.Slides(iSlide).Shapes.AddTextbox msoTextOrientationHorizontal, Left:=120, Top:=50, Width:=300, Height:=50
        ppPres.Slides(iSlide).Shapes(4).TextEffect.text = "Asset " & iAsset + 1 & " - " & assets(1, iAsset)
        ppPres.Slides(iSlide).Shapes(4).TextEffect.Alignment = msoTextEffectAlignmentCentered
        ppPres.Slides(iSlide).Shapes(4).TextEffect.FontSize = 36

        'Table
        ppPres.Slides(iSlide).Shapes.AddTable 5, 2, 70, 520, 400, 200
        With ppPres.Slides(iSlide).Shapes(5)
                Debug.Print (.Name)
                '.Fill.ForeColor.RGB = RGB(0,0,0)  'Hintergrundfarbe
                .table.Columns(1).Width = 100
                .table.Columns(2).Width = 300
                .table.Rows(1).Height = 25
                .table.Rows(2).Height = 25
                .table.Rows(3).Height = 25
                .table.Rows(4).Height = 25
                .table.Rows(5).Height = 25
                .table.Cell(1, 1).Shape.TextFrame2.TextRange.text = "Asset Name"
                .table.Cell(2, 1).Shape.TextFrame2.TextRange.text = "Type"
                .table.Cell(3, 1).Shape.TextFrame2.TextRange.text = "Age"
                .table.Cell(4, 1).Shape.TextFrame2.TextRange.text = "FK Number"
                .table.Cell(5, 1).Shape.TextFrame2.TextRange.text = "Notes"
                
                .table.Cell(1, 2).Shape.TextFrame2.TextRange.text = assets(1, iAsset)
                .table.Cell(2, 2).Shape.TextFrame2.TextRange.text = assets(2, iAsset)
                .table.Cell(3, 2).Shape.TextFrame2.TextRange.text = assets(4, iAsset)
                .table.Cell(4, 2).Shape.TextFrame2.TextRange.text = assets(6, iAsset)
                .table.Cell(5, 2).Shape.TextFrame2.TextRange.text = assets(5, iAsset)
        End With
        
    
        'picture
        Dim pfad As String
        Dim anzahlArray As Variant
        Dim fk As Integer
        Dim anzahl, offset_H, offset_V As Integer
        Dim pictures As Variant
        
        fk = assets(0, iAsset)
        pictures = sql_to_array("Select * from pictures where FK_Asset=" & fk)
        anzahlArray = sql_to_array("Select count(*) from pictures where FK_Asset=" & fk)
        anzahl = anzahlArray(0, 0)
        
        If anzahl = 1 Then
            pfad = pictures(2, 0)
            Call ppPres.Slides(iSlide).Shapes.AddPicture(pfad, msoFalse, msoCTrue, 70, 150, 400, 300)
        End If
        
        If anzahl = 2 Then
            For y = 0 To UBound(pictures, 2)
                pfad = pictures(2, y)
                Call ppPres.Slides(iSlide).Shapes.AddPicture(pfad, msoFalse, msoCTrue, 170, 120 + y * 170, 200, 150)
            Next
        End If
        
        If anzahl = 3 Or anzahl = 4 Then
            For y = 0 To UBound(pictures, 2)
                If y = 1 Or y = 3 Then
                    offset_H = 1
                    Else: offset_H = 0
                End If
                If y = 2 Or y = 3 Then
                    offset_V = 1
                    Else: offset_V = 0
                End If
                
                pfad = pictures(2, y)
                Call ppPres.Slides(iSlide).Shapes.AddPicture(pfad, msoFalse, msoCTrue, 70 + offset_H * 200, 120 + offset_V * 170, 180, 130)
            Next
        End If
        
            

    

        iAsset = iAsset + 1
Loop


'save
ppPres.SaveAs generateFolder(Forms("Cover").Controls("listAssetGroups").column(0)) & "brochure.pptx"
Exit Sub
fehler:
    MsgBox ("An Error occured:" & vbTab & Err.Description)

End Sub


Public Sub Word()

On Error GoTo fehler:
If Forms("Cover").Controls("listAssetGroups").ListIndex <> -1 Then  'check if tender can be created


    Dim wdApp As Word.Application       'create word application

    Set wdApp = New Word.Application     'set new word
    Dim table() As Variant
    Dim arr As Variant
    Dim list() As Variant
    
    On Error Resume Next
    
    
    
    'list of all information
    list = Array("Name", "Age", "Notes")
    wdApp.Visible = True
    wdApp.Activate      'show word file while creation
    wdApp.Documents.Add 'new document
    remark = sql_to_array("Select Remarks from AssetGroups WHERE AssetGroups.Number=" & Forms("Cover").Controls("listAssetGroups").column(0))
    arr = sql_to_array("SELECT Assets.name, Assets.age, Assets.Notes FROM Assets LEFT JOIN AssetGroups ON Assets.FK_AssetGroup=AssetGroups.Number WHERE AssetGroups.Number=" & Forms("Cover").Controls("listAssetGroups").column(0))
      
    sql = "Select * from AssetGroups left join AssetGroupStatus on AssetGroups.Number=AssetGroupStatus.FK_AssetGroup WHERE AssetGroups.Number=" & Forms("Cover").Controls("listAssetGroups").column(0)
    group = sql_to_array(sql)
    
    With wdApp.Selection
    
    .Font.Size = 16  'font
    .ParagraphFormat.Alignment = wdAlignParagraphCenter 'position center
    .Font.Bold = True
    .TypeText ("Tender Sheet")   'Header
    .Font.Bold = False
    .TypeParagraph
    .Font.Size = 12  'font
    .ParagraphFormat.Alignment = wdAlignParagraphLeft 'position center
    .TypeParagraph
        
    .TypeText ("Project No:" & vbTab & vbTab & vbTab & vbTab & "AS_2018/005_CGN")
    .TypeParagraph
        
    .TypeText ("Sales object:" & vbTab & vbTab & vbTab & vbTab & group(1, 0))
    .TypeParagraph
        
    .TypeText ("Department:" & vbTab & vbTab & vbTab & vbTab & group(2, 0))
        '  new line
    .TypeParagraph

    .Font.Bold = True
    .TypeText ("General")
    .Font.Bold = False
    'first table with general information
    generalTable = wdApp.Selection.Tables.Add(wdApp.Selection.Range, 4, 2, wdWord9TableBehavior, wdAutoFitFixed)
    generalTable.Cells(1) = "Department"
    generalTable.Cells(2) = group(map("location"), 0)
    generalTable.Cells(3) = "Business Owner(commercial)"
    generalTable.Cells(5) = "Business Owner(optional)"
      
    controllerCount = sql_to_array("Select count(*) from location where location ='" & group(map("location"), 0) & "'")
    controller = sql_to_array("Select * from location where location ='" & group(map("location"), 0) & "'")
      'get controller
    generalTable.Cells(7) = "Controlling"
    generalTable.Cells(8) = IIf(controllerCount(0, 0) > 0, controller(2, 0), "")
    generalTable.Select
    .Collapse WdCollapseDirection.wdCollapseEnd
    .TypeParagraph
        'second table
    .Font.Bold = True
    .TypeText ("Timeline")
    .Font.Bold = False
    timelineTable = wdApp.Selection.Tables.Add(wdApp.Selection.Range, 7, 2, wdWord9TableBehavior, wdAutoFitFixed)
          
    timelineTable.Cells(1) = "Start Date"
    timelineTable.Cells(2) = group(sqlMap(sql, "creationDate"), 0)
        
    timelineTable.Cells(3) = "RFQ Sent Out Date"
    timelineTable.Cells(4) = group(sqlMap(sql, "brochure"), 0)
        
    timelineTable.Cells(5) = "DUE Date Bidding"
    timelineTable.Cells(6) = group(sqlMap(sql, "InfoDeadline"), 0)
         
    timelineTable.Cells(7) = "Preview of Assets / Site visit"
    timelineTable.Cells(8) = group(sqlMap(sql, "extVisit"), 0)
         
         
    timelineTable.Cells(9) = "Disposal request Nr"
    timelineTable.Cells(10) = group(sqlMap(sql, "DisposalNr"), 0)
         
    timelineTable.Cells(11) = "SA created"
    timelineTable.Cells(12) = group(sqlMap(sql, "signed"), 0)
         
    timelineTable.Cells(13) = "SA signed"
    timelineTable.Cells(14) = group(sqlMap(sql, "SalesAgreement"), 0)
          
    timelineTable.Select
    .Collapse WdCollapseDirection.wdCollapseEnd 'this command is used for geting a new section below the table
    
    .TypeParagraph
    
    .Font.Bold = True
    .TypeText ("Customer/ Purchaser")
    .Font.Bold = False
    brokerSql = "Select BrokerOffer.Id, Broker.Company, BrokerOffer.Offer from BrokerOffer left join Broker on BrokerOffer.FK_Broker=Broker.ID  where FK_AssetGroup=" & Forms("Cover").Controls("listAssetGroups").column(0)
    brokerArray = sql_to_array(brokerSql)
    brokertable = wdApp.Selection.Tables.Add(wdApp.Selection.Range, UBound(brokerArray, 2) + 1 + 1, 2, wdWord9TableBehavior, wdAutoFitFixed)
    
    'brokertable
    brokertable.Cells(1) = "Customer / Purchaser"
    brokertable.Cells(2) = "Offer"
    brokertable.Cells(1).Range.Font.Bold = True
    brokertable.Cells(2).Range.Font.Bold = True
        
    Dim highestOffer As Integer
    highestOffer = 0
    'for each broker get offer and write into table
    For i = 0 To UBound(brokerArray, 2)
        brokertable.Cells(1 + i * 2 + 2) = brokerArray(1, i)
        brokertable.Cells(2 + i * 2 + 2) = brokerArray(2, i) & "€"
        If brokerArray(2, i) > highestOffer Then
            highestOffer = brokerArray(2, i)
        End If
    Next
    brokertable.Select

    wdApp.Selection.Collapse WdCollapseDirection.wdCollapseEnd
    
    wdApp.Selection.TypeParagraph
    
    .Font.Bold = True
    .TypeText ("Recommendation")
    .Font.Bold = False
    reccTable = wdApp.Selection.Tables.Add(wdApp.Selection.Range, 1, 1, wdWord9TableBehavior, wdAutoFitFixed)
    reccTable.Select
    .Collapse WdCollapseDirection.wdCollapseEnd
    .TypeParagraph

    .Font.Bold = True
    .TypeText ("Remarks")
    .Font.Bold = False
    
    'get remark without html tags
    remarkTable = wdApp.Selection.Tables.Add(wdApp.Selection.Range, 1, 1, wdWord9TableBehavior, wdAutoFitFixed)
    remarkTable.Cells(1) = Replace(Replace(Replace(Replace(remark(0, 0), "<div>", ""), "</div>", ""), "&amp", ""), "&nbsp", "")
    remarkTable.Select
    .Collapse WdCollapseDirection.wdCollapseEnd
    .TypeParagraph

    .Font.Bold = True
    .TypeText ("Approval")
    .Font.Bold = False
    'new approve table
    approvalTable = wdApp.Selection.Tables.Add(wdApp.Selection.Range, 1, 1, wdWord9TableBehavior, wdAutoFitFixed)
    If highestOffer >= 5000 Then
        approvalTable.Cells(1) = "Bernd Götz"
    Else
        approvalTable.Cells(1) = "Jens Unterhansberg"
    End If
    approvalTable.Select
    .Collapse WdCollapseDirection.wdCollapseEnd
    
    .TypeParagraph
    
    .InsertBreak (wdPageBreak)
    
    .ParagraphFormat.Alignment = wdAlignParagraphCenter 'position center
    
    .Font.Bold = True
    .TypeText ("Assets")
    .Font.Bold = False
    .TypeParagraph

    
    'get information from asset gtoup out of database
    
    ReDim table(UBound(arr, 2)) As Variant  'redim table with the count of assets as index
    
    
    
    .ParagraphFormat.Alignment = wdAlignParagraphCenter 'position center
    
    
    'for all assets
    For Z = 0 To UBound(arr, 2)
        wdApp.Selection.TypeText (arr(0, Z))
        table(Z) = wdApp.Selection.Tables.Add(wdApp.Selection.Range, UBound(arr, 1) + 1, 2, wdWord9TableBehavior, wdAutoFitFixed)
        'new table
        
        For i = 0 To UBound(list)   'fill table with information header from list
            table(Z).Cells(1 + (i * 2)) = list(i)
        Next
    
        For i = 0 To UBound(arr, 1)
        'fill in asset information
            If (IsNull(arr(i, 0)) = False) Then
                table(Z).Cells(2 + (i) * 2) = arr(i, Z)
            End If
        Next
    
    
    
        'get table selected and move to place below the table. Important for writing outside the table again
        table(Z).Select
        .Tables(1).Rows.Alignment = wdAlignRowCenter
        .Tables(1).Columns(1).PreferredWidth = 75
        .Collapse WdCollapseDirection.wdCollapseEnd
        .TypeParagraph
    

    Next

    'generate folder just to check if existing
    


        Call Update("AssetGroupStatus", "FK_AssetGroup", Forms("Cover").Controls("listAssetGroups").column(0), "Tender", Date)
        Call wdApp.ActiveDocument.SaveAs2(generateFolder(Forms("Cover").Controls("listAssetGroups").column(0)) & "TenderSheet.docx") 'save tender in folder
        'wdApp.Documents.Close (wdDoNotSaveChanges)
        'wdApp.Quit
        Set wdApp = Nothing
        
        
    End With
Else
    MsgBox ("No Asset Group selected or no Assets for Asset Group") 'reson why no tender can be created
    
End If


Exit Sub
fehler:
    MsgBox ("An Error occured:" & vbTab & Err.Description)



End Sub





