Sub alphTest()
	Dim ws                 As Worksheet
	Dim labelNames         As Variant
	'References to Columns
	Dim tickerColID        As Integer
	Dim yearOpenColID      As Integer
	Dim yearCloseColID     As Integer
	Dim volColID           As Integer
	Dim tickerOUTColID     As Integer
	Dim yearChangeColID    As Integer
	Dim percentOUTColID    As Integer
	Dim volOUTColID        As Integer	
	
	labelNames = Array("<ticker name>", "<yearly change>", "<percent change>", "<total stock volume>")
	
	tickerColID = 1
	yearOpenColID = 3
	yearCloseColID = 6
	volColID = 7
	tickerOUTColID = 9
	yearlyOUTColID = 10
	percentOUTColID = 11
	volOUTColID = 12

	
	For Each ws In ThisWorkbook.Worksheets
		'These loops apply labels to all the sheets for the data analysis outputs
		counter = 0
		For c = tickerOUTColID To volOUTColID
			ws.Cells(1, c).Value = labelNames(counter)
			counter = counter + 1
		Next c
		   
		   
		'Dynamic arrays that will store values to be outputed
		Dim tickerNames()     As String
		Dim yearOpens()       As Double
		Dim yearCloses()      As Double
		Dim yearlyChanges()   As Double
		Dim totalVolumes()    As Double
		'Index used to move through the aforementioned arrays
		'Looks at unique ticker names; only changes when we see a new ticker
		Dim startTicker       As Long
		Dim currentTicker     As Long
		'The row of data being examined/compared
		Dim currentRow        As Long
		'rowCountA defines the length of Column A
		Dim rowCountA         As Long
	

		rowCountA = ws.Cells(Rows.Count, tickerColID).End(xlUp).Row
		
		'These arrays are started with index 2 to match the row number that they will ultimately refer to
		currentTicker = 2
		startTicker = 2
		
		ReDim tickerNames(startTicker To currentTicker)
		ReDim yearOpens(startTicker To currentTicker)
		ReDim totalVolumes(startTicker To currentTicker)
		
		'The following four lines initialize the first values in given arrays
		'Note: yearCloses() does not occur until the last row of each unique tickerName
		'therefore it is not initialized here
		currentRow = 2
		tickerNames(currentTicker) = ws.Cells(currentRow, tickerColID).Value
		yearOpens(currentTicker) = ws.Cells(currentRow, yearOpenColID).Value
		totalVolumes(currentTicker) = ws.Cells(currentRow, volColID).Value
	
		'This loop is comparing the previous row's entry
		For currentRow = 3 To rowCountA
			'This code grows the parallel arrays by 1 to accommodate the new values being added
			If ws.Cells(currentRow, tickerColID).Value <> ws.Cells(currentRow - 1, tickerColID).Value Then
				currentTicker = currentTicker + 1
				ReDim Preserve yearCloses(startTicker To currentTicker)
				ReDim Preserve tickerNames(startTicker To currentTicker)
				ReDim Preserve yearOpens(startTicker To currentTicker)
				ReDim Preserve totalVolumes(startTicker To currentTicker)

				'The yearClose value is on the previous row where the "Else" condition was still true
				'hence the "currentTicker - 1" and "currentRow - 1"
				yearCloses(currentTicker - 1) = ws.Cells(currentRow - 1, yearCloseColID).Value
				tickerNames(currentTicker) = ws.Cells(currentRow, tickerColID).Value
				yearOpens(currentTicker) = ws.Cells(currentRow, yearOpenColID).Value
				totalVolumes(currentTicker) = ws.Cells(currentRow, volColID).Value
			
			'Accumulating the new volume into the currentTicker's total volume
			Else
				totalVolumes(currentTicker) = totalVolumes(currentTicker) + ws.Cells(currentRow, volColID).Value
			End If
		Next currentRow
		
		'Because the conditions of the previous loop will not include the final value for the yearCloses(),
		'the final value is added outside of the loop afterwards, much like the initial values for
		'tickerNames(), yearOpens(), and totalVolumes() were added before the loop.
		yearCloses(currentTicker) = ws.Cells(rowCountA, yearCloseColID).Value
		
		'This loop writes out the values of the arrays to the appropriate columns.
		'Using an array to store the data while being calculated instead of writing it to the cell each iteration
		'saves a huge amount of time and computing power.
		For currentRow = startTicker To currentTicker
			ws.Cells(currentRow, tickerOUTColID).Value = tickerNames(currentRow)
			ws.Cells(currentRow, yearlyOUTColID).Value = yearCloses(currentRow) - yearOpens(currentRow)
			ws.Cells(currentRow, volOUTColID).Value = totalVolumes(currentRow)
		Next currentRow
		
		'This loop writes out the percent change column.  Because there is no condition established for 
		'cases where we divide by 0, the "If" statement assigns a sentinel number to the cell.
		'Note:  Do you know why 12509 is an interesting number?
		For currentRow = startTicker To currentTicker
			If yearOpens(currentRow) = 0 Then
				ws.Cells(currentRow, percentOUTColID).Value = 12509	
			Else
				ws.Cells(currentRow, percentOUTColID).Value = ((ws.Cells(currentRow, yearlyOUTColID).Value) / yearOpens(currentRow))
			End If
		Next currentRow
		

		'This assigns green to values greater than or equal to 0 and red to all others (i.e. negative numbers)
		For currentRow = startTicker To currentTicker
			If ws.Cells(currentRow, yearlyOUTColID).Value >= 0 Then
				ws.Cells(currentRow, yearlyOUTColID).Interior.ColorIndex = 4
			Else
				ws.Cells(currentRow, yearlyOUTColID).Interior.ColorIndex = 3
			End If
		Next currentRow
		
		ws.Columns(percentOUTColID).NumberFormat = "0.00%"
		
		
	Next ws
	
End Sub
