# Hide PowerShell Console
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '320,270'
$Form.text                       = "Remote Text to Speech"
$Form.TopMost                    = $false
$Form.MaximizeBox                = $false
$Form.FormBorderStyle            = 'Fixed3D'

$Computer                        = New-Object system.Windows.Forms.TextBox
$Computer.multiline              = $false
$Computer.width                  = 150
$Computer.height                 = 20
$Computer.location               = New-Object System.Drawing.Point(10,16)
$Computer.Font                   = 'Microsoft Sans Serif,10'

$ComboBox1                       = New-Object system.Windows.Forms.ComboBox
$ComboBox1.text                  = "Computers"
$ComboBox1.width                 = 150
$ComboBox1.height                = 20
$ComboBox1.location              = New-Object System.Drawing.Point(10,16)
$ComboBox1.Font                  = 'Microsoft Sans Serif,10'

$text                            = New-Object system.Windows.Forms.TextBox
$text.multiline                  = $false
$text.width                      = 300
$text.height                     = 20
$text.location                   = New-Object System.Drawing.Point(10,70)
$text.Font                       = 'Microsoft Sans Serif,10'
$text.text                       = ""

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Computer Name"
$Label1.AutoSize                 = $true
$Label1.location                 = New-Object System.Drawing.Point(161,21)
$Label1.Font                     = 'Microsoft Sans Serif,10'

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Text:"
$Label2.AutoSize                 = $true
$Label2.location                 = New-Object System.Drawing.Point(10,50)
$Label2.Font                     = 'Microsoft Sans Serif,10'

$SubmitButton                    = New-Object system.Windows.Forms.Button
$SubmitButton.text               = "Submit"
$SubmitButton.width              = 300
$SubmitButton.height             = 30
$SubmitButton.location           = New-Object System.Drawing.Point(10,230)
$SubmitButton.Font               = 'Microsoft Sans Serif,10'

$Label4                          = New-Object system.Windows.Forms.Label
$Label4.text                     = ""
$Label4.AutoSize                 = $true
$Label4.location                 = New-Object System.Drawing.Point(10,240)
$Label4.Font                     = 'Microsoft Sans Serif,10'

$Male                            = New-Object system.Windows.Forms.RadioButton
$Male.text                       = "Male Voice"
$Male.AutoSize                   = $true
$Male.width                      = 104
$Male.height                     = 20
$Male.location                   = New-Object System.Drawing.Point(10,120)
$Male.Font                       = 'Microsoft Sans Serif,10'
$Male.checked                    = $true

$Female                          = New-Object system.Windows.Forms.RadioButton
$Female.text                     = "Female Voice"
$Female.AutoSize                 = $true
$Female.width                    = 104
$Female.height                   = 20
$Female.location                 = New-Object System.Drawing.Point(10,140)
$Female.Font                     = 'Microsoft Sans Serif,10'
$female.checked                  = $false

$Test                            = New-Object System.Windows.Forms.CheckBox
$Test.text                       = "Test on local computer"
$Test.AutoSize                   = $true
$Test.width                      = 104
$Test.height                     = 20
$Test.location                   = New-Object System.Drawing.Point(10,100)
$Test.Font                       = 'Microsoft Sans Serif,10'

$speed                           = New-Object system.Windows.Forms.Trackbar
$speed.width                     = 300
$speed.height                    = 20
$speed.location                  = New-Object System.Drawing.Point(10,190)
$speed.Font                      = 'Microsoft Sans Serif,10'
$speed.value                     = 0
$speed.SetRange(-10,10)

$Label5                          = New-Object system.Windows.Forms.Label
$Label5.text                     = "Voice Speed (0)"
$Label5.AutoSize                 = $true
$Label5.location                 = New-Object System.Drawing.Point(10,170)
$Label5.Font                     = 'Microsoft Sans Serif,10'

$computers = (get-adcomputer -filter *).name | Sort-Object

$cred = Get-Credential

$speed.add_valuechanged({
    $speedvalue = $speed.value
    $label5.text = "Voice Speed ($speedvalue)"
})

if (-not ([string]::IsNullOrEmpty($computers))) { 
    $Form.controls.AddRange(@($ComboBox1,$text,$Label1,$SubmitButton,$Label2,$Label4,$speed,$male,$female,$label5,$test))
    foreach ($line in $computers) { [void]$ComboBox1.Items.Add($line) }
} else { 
    $Form.controls.AddRange(@($Computer,$text,$Label1,$SubmitButton,$Label2,$Label4,$male,$speed,$female,$label5,$test))
}

$SubmitButton.Add_Click({
    $SubmitButton.enabled = $false
    $label4.text = ""
    $speech = $text.text
    $speedvalue = $speed.value
    if ($male.checked -eq $true) {$voice = 'Microsoft David Desktop'}
    if ($female.checked -eq $true) {$voice = 'Microsoft Zira Desktop'}
    if (-not ([string]::IsNullOrEmpty($computers))) { $comp = $ComboBox1.SelectedItem } else { $comp = $Computer.Text }
    if ($test.checked -eq $false) {
        if ( (test-connection $comp -count 1 -quiet) -eq $true ) {
            $SubmitButton.enabled = $false
            invoke-command -computername $comp -scriptblock {
                [Reflection.Assembly]::LoadWithPartialName('System.Speech') | Out-Null   
                $object = New-Object System.Speech.Synthesis.SpeechSynthesizer 
                $object.rate = ($using:speedvalue)
                $object.SelectVoice($using:voice)
                $object.Speak($using:speech) 
            } -Credential $cred
        } else { $label4.text = "Computer unavailable" }
    }
    if ($test.checked -eq $true) {
        [Reflection.Assembly]::LoadWithPartialName('System.Speech') | Out-Null   
        $object = New-Object System.Speech.Synthesis.SpeechSynthesizer 
        $object.rate = ($speedvalue)
        $object.SelectVoice($voice)
        $object.Speak($speech) 
    }
    $SubmitButton.enabled = $true
})

[void]$Form.ShowDialog()
