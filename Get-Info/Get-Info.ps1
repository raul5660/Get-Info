## TODO ##
Param([String]$Path)

$RegularExpression = "\d{3}-\d{2}-\d{4}|(?:\d[ -]*?){13,16}"
$files = (Get-ChildItem -Force $Path -Recurse | Select-Object Directory, Name)

foreach($file in $files)
{
	if ($file.Directory -ne $null)
	{
		$filePath =  (-join ($file.Directory, "\", $file.Name))
		$fileType = $file.Name.Split('.')[-1]
		if($filePath.Contains('~'))
		{
			continue
		}
		switch($fileType)
		{
			"txt"
			{
				$RegularExpressionMatches = (Select-String -Path $filePath -Pattern $RegularExpression -AllMatches)
				if ($RegularExpressionMatches.Count -gt 0)
				{
					Write-Host $filePath
				}
			}
			"rtf"
			{
				$RegularExpressionMatches = (Select-String -Path $filePath -Pattern $RegularExpression -AllMatches)
				if ($RegularExpressionMatches.Count -gt 0)
				{
					Write-Host $filePath
				}
			}
			"xlsx"
			{
				$excel = New-Object -ComObject Excel.Application
				$excel.Visible = $false
				$workbook = $excel.WorkBooks.Open($filePath, $false, $true)
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
			"accdb"
			{
				Write-Host "Access Database File"
			}
			"docx"
			{
				$word = New-Object -ComObject Word.Application
				$word.Visible = $false
				$doc = $word.Documents.Open($filePath, $false, $true)
				$data = $doc.WordOpenXML.toString()
				$RegularExpressionMatches = (Select-String -InputObject $data -Pattern $RegularExpression -AllMatches)
				if ($RegularExpressionMatches.Count -gt 0)
				{
					Write-Host $filePath
				}
			}
			"msg"
			{
				$outlook = New-Object -ComObject Outlook.Application
				$msg = $outlook.CreateItemFromTemplate($filePath)
				$data = -join ($msg.Body,$msg.Subject)
				$RegularExpressionMatches = (Select-String -InputObject $data -Pattern $RegularExpression -AllMatches)
				if ($RegularExpressionMatches.Count -gt 0)
				{
					Write-Host $filePath
				}
			}
		}
	}
}