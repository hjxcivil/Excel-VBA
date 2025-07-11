# Scraping Websites with Internet Explorer or XML HTTP Requests

[TOC]

## Controlling Internet Explorer with VBA

- Referencing the Required Object Libraries

  - *Microsoft Internet Controls*

  - *Microsoft HTML Object Library*

    `Dim IE As New SHDocVw.InternetExplorer`
    `Dim HTMLDoc As MSHTML.HTMLDocument`

- Navigating to a Web Page
  `IE.navigate "https://www.oddschecker/com/golf/memorial-tournament/winner"`

- Checking that IE is Ready

- Referencing an HTML Document

- Getting a Reference to an Element by ...

- Testing the Code

  ^Sub *ScrapeOddsUsingIE*()

      Dim IE As New SHDocVw.InternetExplorer
      Dim HTMLDoc As MSHTML.HTMLDocument
      Dim HTMLDiv As MSHTML.IHTMLElement
      Dim HTMLTable As MSHTML.IHTMLElement
          
      IE.Visible = True
      IE.navigate "https://www.oddschecker.com/golf/memorial-tournament/winner"
      
      Do While IE.readyState <> READYSTATE_COMPLETE Or IE.Busy
      Loop
      
      Set HTMLDoc = IE.document
      Set HTMLDiv = HTMLDoc.getElementById("oddsTableContainer")
      Set HTMLTable = HTMLDiv.getElementsByTagName("table")(0)
      
      Debug.Print HTMLTable.className

  End Sub

## The XML HTTP Request Approach

- Referencing the XML Library : *Microsoft XML, v6.0*

  `Dim XMLRequest As New MSXML2.XMLHTTP60`

- Opening a Request

  `XMLRequest.Open "GET", "https://www.oddschecker.com/golf/memorial-tournament/winner", False`

- Sending a Request and Checking the Response

- Testing the Code

  ^Sub *ScrapeOddsUsingXMLHTTP*()

      Dim XMLRequest As New MSXML2.XMLHTTP60
      Dim HTMLDoc As New MSHTML.HTMLDocument
      Dim HTMLDiv As MSHTML.IHTMLElement
      Dim HTMLTable As MSHTML.IHTMLElement
      
      XMLRequest.Open "GET", "https://www.oddschecker.com/golf/memorial-tournament/winner", False
      XMLRequest.send
      
      If XMLRequest.Status <> 200 Then
          MsgBox XMLRequest.Status & " - " & XMLRequest.statusText
          Exit Sub
      End If
      
      HTMLDoc.body.innerHTML = XMLRequest.responseText
      
      Set HTMLDiv = HTMLDoc.getElementById("oddsTableContainer")
      Set HTMLTable = HTMLDiv.getElementsByTagName("table")(0)
      
      Debug.Print HTMLTable.className

  End Sub

## Creating a Method to Process the Table

- *WriteTableToWorksheet* Procedure

  - *Declaring the Required Variables*

  - *Looping Through Table Elements*

  - *Setting the Row and Column Numbers*

  - *Writing the Information to a Worksheet*

  - *Getting Extra Information from the Web Page*

    Sub *WriteTableToWorksheet*(*TableToProcess* As MSHTML.IHTMLElement)

        Dim TableSection As MSHTML.IHTMLElement
        Dim TableRow As MSHTML.IHTMLElement
        Dim TableCell As MSHTML.IHTMLElement
        Dim BookieLink As MSHTML.IHTMLElement
        Dim RowNum As Long, ColNum As Long
        Dim OutputSheet As Worksheet
        
        RowNum = 0
        ColNum = 0
        
        Set OutputSheet = ThisWorkbook.Worksheets.Add
        
        For Each TableSection In TableToProcess.Children
        
            For Each TableRow In TableSection.Children
            
                RowNum = RowNum + 1
                
                For Each TableCell In TableRow.Children
                
                    ColNum = ColNum + 1
                    
                    If TableRow.className = "eventTableHeader" Then
                        
                        Set BookieLink = TableCell.getElementsByTagName("a")(0)
                        
                        If Not BookieLink Is Nothing Then
                            OutputSheet.Cells(RowNum, ColNum).Value = BookieLink.Title
                        End If
                        
                        Set BookieLink = Nothing
                    Else
                        OutputSheet.Cells(RowNum, ColNum).Value = TableCell.innerText
                    End If
                    
                Next TableCell
                
                ColNum = 0
                
            Next TableRow
        
        Next TableSection

    End Sub

- Call the Subroutine

  - With Internet Controls

    *Sub *ScrapeOddsUsingIE*()

        ...
        Debug.Print HTMLTable.className
        WriteTableToWorksheet HTMLTable
        
        IE.Quit

    End Sub

  - With XML Request

    *Sub *ScrapeOddsUsingXMLHTTP*()

        ...
        Debug.Print HTMLTable.className
        
        WriteTableToWorksheet HTMLTable

    End Sub



## Tbd