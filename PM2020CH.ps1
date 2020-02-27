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

$picturespath="C:\AIDemo\Pictures\PM2020ChihSlides\"
$textpath="C:\AIDemo\Text"


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
$file = $picturespath + "Slide1.png"
$pictureBox.Image = [System.Drawing.Image]::Fromfile($file)

Start-Sleep -Seconds $wait
$form.Refresh()
$form.bringtofront()

$Slides = import-csv “.\text\Texto.txt” –header Id,Texto,Voz
$OldSlide = ""

ForEach ($Slide in $Slides){
  if ($Slide.id -ne $oldSlide){
      $OldSlide = $Slide.id  
      $file = $picturespath + $OldSlide +".png"
      $pictureBox.Image = [System.Drawing.Image]::Fromfile($file)
   }
  Start-Sleep -Seconds $wait
  $form.Refresh()
  $form.bringtofront()

  Out-Speech $Slide.Texto $Slide.Voz
}

