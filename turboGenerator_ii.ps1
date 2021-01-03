
# start of with the right language
Set-Culture -CultureInfo se-SV
# set root for all needed resource documents
$PARTY_ROOT = "C:\Users\...\Documents\turbofest\stage\"
# make sure that there is a mall.pptx, bilder and sounds in the folder

Add-type -AssemblyName office
$Application = New-Object -ComObject powerpoint.application
$application.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$slideType = “microsoft.office.interop.powerpoint.ppSlideLayout” -as [type]

# presentationen som vi använder, behåll en första sida som mall
$template = $PARTY_ROOT + "mall.pptx"
$presentation = $application.Presentations.open($template)
$sld = $presentation.Slides[1]

# get data from csv
$csvpath = $PARTY_ROOT + "data.csv"
$csv = Import-Csv -Encoding Default -Delimiter ";" $csvpath

# prepare soundfiles
$soundHorn = $PARTY_ROOT + "ljud\horn.mp3"
$soundTrumpet = $PARTY_ROOT + "ljud\trumpet.wav"

$targetYear = 2021 #set this
$date = Get-Date -Year $targetYear -Month 1 -Day 1
$numberOfDays = (Get-Date -Year $targetYear -Month 12 -Day 31).DayOfYear # 365 or 366 days

# debug keep to 5
$numberOfDays = 200

# make sure there are slides enough
While($presentation.Slides.Count -lt $numberOfDays) {
    $presentation.Slides.item(1).duplicate()
}

# order of shapes in the slide
$datum = 2
$veckodag = 3
$attGora = 4
$img = 1

# iterate over all days and set the appropriate date and information
For ($dayOfYear = 1; $dayOfYear -le $numberOfDays; $dayOfYear ++) {
    $sld = $presentation.Slides.item($dayOfYear)

    $month = (Get-Culture).DateTimeFormat.GetMonthName($date.Month)
    $weekDay = (Get-Culture).DateTimeFormat.GetDayName($date.DayOfWeek)
    $dayOfMOnth = $date.Day

    # capital letter on weekday (fredag > Fredag)
    $TextInfo = (Get-Culture).TextInfo
    $weekDay = $TextInfo.ToTitleCase($weekDay)

    $sld.Shapes[$datum].TextFrame.Textrange.Text = "$dayOfMonth $month"
    $sld.Shapes[$veckodag].TextFrame.Textrange.Text = $weekDay

    # default på dagar
    if ($weekDay -eq "Måndag") {
        $sld.Shapes[$attGora].TextFrame.Textrange.Text = "Bara att köra"
    }
    if ($weekDay -eq "Tisdag") {
        $sld.Shapes[$attGora].TextFrame.Textrange.Text = "Läsa läxor "
    }
    if ($weekDay -eq "Onsdag") {
        $sld.Shapes[$attGora].TextFrame.Textrange.Text = "Taco-tisdag, fast på en onsdag `nLoka, öl och chips"
    }
    if ($weekDay -eq "Torsdag") {
        $sld.Shapes[$attGora].TextFrame.Textrange.Text = "Chilla"
    }
    if ($weekDay -eq "Fredag") {
        $sld.Shapes[$attGora].TextFrame.Textrange.Text = "Fredag! `n16:00 Städröj upp runt tallriken.`n17:00 Baren öppnar, Loka, lättdryck, kall öl och chirre"
    }
    if ($weekDay -eq "Lördag") {
        $sld.Shapes[$attGora].TextFrame.Textrange.Text = "Shhh .... Melodikrysset"
    }
    if ($weekDay -eq "Söndag") {
        $sld.Shapes[$attGora].TextFrame.Textrange.Text = "Helgmålsbön"
    }

    # set background to the image named as [month].jpg in resource\bilder folder
    $backgroundImage = $PARTY_ROOT + "bilder\" + $month + ".jpg"
    $sld.Shapes[$img].Fill.UserPicture($backgroundImage)
    $sld.Shapes[$img].sendtoback

    # check for special dates
    $dateStr = $date.ToString("yyyyMMdd")
    foreach($item in $csv){
        if ($($item.Datum).ToString() -eq $dateStr) {
            # only fill if something to fill with
            if($($item.Aktivitet) -ne "") {
                $sld.Shapes[$attGora].TextFrame.Textrange.Text = $($item.Aktivitet)
            }
            # add temadags name på the date
            $sld.Shapes[$datum].TextFrame.Textrange.Text = "$dayOfMonth $month" + "  -  " + $($item.Namn)

            # sound the trumpet if food
            if($($item.Mat) -ne "") {
                $sld.Shapes.AddMediaObject2($soundTrumpet, 10, 10, 10, 10)
            } else {
                $sld.Shapes.AddMediaObject2($soundHorn, 10, 10, 10, 10)
            }
        }
    }
    # goto next date
    $date = $date.addDays(1)
}

$presentation.Save()
$presentation.Close()
“Modifying $template”

$application.quit()
$application = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()



