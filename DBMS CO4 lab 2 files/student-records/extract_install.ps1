Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead('c:\Users\sowse\Downloads\1.Student_Records_MongoDB_VSCode_Installation_Guide.docx')
$entry = $zip.GetEntry('word/document.xml')
$stream = $entry.Open()
$reader = New-Object System.IO.StreamReader($stream)
[xml]$xml = $reader.ReadToEnd()
$reader.Close()
$zip.Dispose()
$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
$paras = $xml.SelectNodes('//w:p', $ns) | ForEach-Object { $_.InnerText }
$paras | Out-File -FilePath 'c:\Users\sowse\Downloads\student-records\install_doc.txt' -Encoding utf8
