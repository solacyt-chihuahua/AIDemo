

function Out-Speech( [string]$text, [string]$Sex)
{
	
	$speechy = New-Object –ComObject SAPI.SPVoice;
	
	$voices = $speechy.GetVoices();
	
	IF ($Sex -eq "L")
	{
		$voiceType = '*david*'
	}
	ELSEIF ($Sex -eq "R")
	{
		$voiceType = 'Microsoft Raul - Spanish (Mexico)'
	}
	ELSE
	{
		$voiceType = 'Microsoft Sabina - Spanish (Mexico)'
	}
	
	
	foreach ($voice in $voices)
	{
		# without if statement both zira and david would talk
		IF ($voice.GetDescription() -like $voiceType)
		{
			$voice.GetDescription() | Out-Null
			$speechy.Voice = $voice;
			$speechy.Speak($text) | Out-Null
                        break
		}
		
	}
}

$picturespath="C:\AIDemo\Pictures\"
$videospath="C:\AIDemo\Videos\"

$wait = 1
[void][reflection.assembly]::LoadWithPartialName("System.Windows.Forms")
$form = new-object Windows.Forms.Form
$form.Text = "Image Viewer"
$form.WindowState= "Maximized"
$form.controlbox = $false
$form.formborderstyle = "0"
$form.BackColor = [System.Drawing.Color]::black

$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.dock = "fill"
$pictureBox.sizemode = 4
$form.controls.add($pictureBox)
$form.Add_Shown( { $form.Activate()} )
$form.Show()

$file = $picturespath + "Interface1.png"
$pictureBox.Image = [System.Drawing.Image]::Fromfile($file)
Start-Sleep -Seconds $wait
$form.Refresh()
$form.bringtofront()


Add-Type -AssemblyName System.speech
$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
$speak.Rate   = 10  # -10 to 10; -10 is slowest, 10 is fastest
$speak.Speak('Hello...')
$speak.Speak('Testing, Testing...1 2 3 4 5 6 7')
$speak.Rate   = -10  # -10 to 10; -10 is slowest, 10 is fastest
$speak.Speak('1 2 3 4 5...')
$speak.Rate   = -5  # -10 to 10; -10 is slowest, 10 is fastest
$speak.Speak('Testing')
$speak.Rate   = 0  # -10 to 10; -10 is slowest, 10 is fastest
$speak.Speak('Probando, Probando...')
$speak.Speak('I think it is ready...')
$speak.Speak("I'll try to speak spanish now...")
$speak.Speak('Hola...')
$speak.Speak('Buenos Dias...')
$speak.Speak('Are you ready?')
$speak.Speak('Estan Listos?')

Out-Speech "Si, estamos listos" M
Out-Speech "Yo, tambien estoy listo" R
Out-Speech "Estan ustedes listos?" M
Out-Speech "Laura, Jacobo, Nohemi, Jose y Fernando, Estan listos?" R

Out-Speech "Excelente.., Muchas Gracias" M
Out-Speech "All right!!! hehe... bye!" L

$form.close()