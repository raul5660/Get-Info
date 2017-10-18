## TODO ##

function Get-Info
{
	Param([String]$Path)

	#$RegularExpression = "\d{3}-\d{2}-\d{4}|\d{4}-\d{4}-\d{4}-\d{4}"
	$files = (Get-ChildItem -Force $Path -Recurse | Select-Object Directory, Name)

	function TextDocuments
	{
		Param([String]$CompleteFilePath)
		$RegularExpressionMatches = (Select-String -Path $CompleteFilePath -Pattern $RegularExpression -AllMatches)
		if ($RegularExpressionMatches.Count -gt 0)
		{
			Write-Host $filePath
		}
	}

	function SpreadSheets
	{
		Param([String]$CompleteFilePath)
		$excel = New-Object -ComObject Excel.Application
		$excel.Visible = $false
		$workbook = $excel.WorkBooks.Open($CompleteFilePath, $false, $true)
		$done = $false
		foreach($spreadsheet in $workbook.Sheets)
		{
			$columnCount = 5000
			$rowCount = 5000
			if($done){break}
			for($x = 1; $x -lt $rowCount; $x++)
			{
				if($done){break}
				for($y = 1; $y -lt $columnCount; $y++)
				{
					$data = $spreadsheet.Cells.Item($y,$x).Text
					$RegularExpressionMatches = (Select-String -InputObject $data -Pattern $RegularExpression -AllMatches)
					if ($RegularExpressionMatches.Count -gt 0)
					{
						Write-Host $filePath
						$done = $true
						break
					}
				}
			}
		}
		$done = $false
	}

	function WordDouments
	{
		Param([String]$CompleteFilePath)
		$word = New-Object -ComObject Word.Application
		$word.Visible = $false
		$doc = $word.Documents.Open($CompleteFilePath, $false, $true)
		$data = $doc.WordOpenXML.toString()
		$RegularExpressionMatches = (Select-String -InputObject $data -Pattern $RegularExpression -AllMatches)
		if ($RegularExpressionMatches.Count -gt 0)
		{
			Write-Host $filePath
		}
	}

	function MessageFile
	{
		Param([String]$CompleteFilePath)
		$outlook = New-Object -ComObject Outlook.Application
		$msg = $outlook.CreateItemFromTemplate($CompleteFilePath)
		$data = -join ($msg.Body,$msg.Subject)
		$RegularExpressionMatches = (Select-String -InputObject $data -Pattern $RegularExpression -AllMatches)
		if ($RegularExpressionMatches.Count -gt 0)
		{
			Write-Host $filePath
		}
	}

	function PSTFile
	{
		Param([String]$CompleteFilePath)
		$oProc = ( Get-Process | where { $_.Name -eq "OUTLOOK" } )
		if ( $oProc -eq $null ) 
		{ 
			Start-Process outlook -WindowStyle Hidden
			Start-Sleep -Seconds 5 
		}
		$outlook = New-Object -ComObject Outlook.Application
		$namespace = $outlook.GetNamespace("MAPI")
		$namespace.AddStoreEx($CompleteFilePath, 1)
		$pstStore = ( $nameSpace.Stores | where { $_.FilePath -eq $CompleteFilePath } )
		$pstRootFolder = $pstStore.GetRootFolder()
		$inboxFolder = $pstRootFolder.Folders|? { $_.Name -eq 'Inbox' }
		foreach($email in $inboxFolder.Items)
		{
			$data = -join ($email.Body,$email.Subject)
			$RegularExpressionMatches = (Select-String -InputObject $data -Pattern $RegularExpression -AllMatches)
			if ($RegularExpressionMatches.Count -gt 0)
			{
				Write-Host $filePath,":",$email.SenderEmailAddress,":",$email.Subject,":",$email.ReceivedTime
			}
		}
	}

	function PDF
	{
		Param([String]$CompleteFilePath)
		Add-Type -Path .\itextsharp-dll-core\itextsharp.dll

		$CompleteFilePath;

		$reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $CompleteFilePath

		for ($page = 1; $page -le $reader.NumberOfPages; $page++)
		{
			$text = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader,$page).Split([char]0x000A)

			foreach ($line in $text)
			{
				$RegularExpressionMatches = (Select-String -InputObject $line -Pattern $RegularExpression -AllMatches)
				if ($RegularExpressionMatches.Count -gt 0)
				{
						Write-Host $filePath,": Page ",$page,"; Matches: ",$RegularExpressionMatches
				}
			}
		}

		$reader.Close()
	}

	foreach($file in $files)
	{
		if ($file.Directory -ne $null)
		{
			$filePath =  (-join ($file.Directory, "\", $file.Name))
			$fileType = ($file.Name.Split('.')[-1]).toLower()
			if($filePath.Contains('~'))
			{
				continue
			}
			switch($fileType)
			{
				"txt"
				{
					TextDocument -CompleteFilePath $filePath
				}
				"rtf"
				{
					TextDocument -CompleteFilePath $filePath
				}
				"xlsx"
				{
					SpreadSheets -CompleteFilePath $filePath
				}
				"accdb"
				{
					Write-Host "Access Database File"
				}
				"docx"
				{
					WordDocument -CompleteFilePath $filePath
				}
				"msg"
				{
					MessageFile -CompleteFilePath $filePath
				}
				"pst"
				{
					PSTFile -CompleteFilePath $filePath	
				}
				"pdf"
				{
					Write-Host "PDF File"
					PDF -CompleteFilePath $filePath
				}
			}
		}
	}
}