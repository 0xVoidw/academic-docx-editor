param([string]$Python = 'python', [switch]$ProbeWord)
$arguments = @('-m', 'academic_docx_editor', 'doctor')
if ($ProbeWord) { $arguments += '--probe-word' }
& $Python @arguments
exit $LASTEXITCODE
