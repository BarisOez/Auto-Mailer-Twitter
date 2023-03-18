#************************** A U T O - M A I L E R / A U T O - T W I T T E R ****************************#
#*******************************************************************************************************#
#**************************  A U T H O R -  B A R I S  O E Z C A N *************************************#
#*******************************************************************************************************#
#*******************************************************************************************************#
#*******************************************************************************************************#
                     #********************************************************#
                     #********************************************************#





###########################################      Funktionen Start   ################################################################
####################################################################################################################################
####################################################################################################################################

# Hier wird getestet ob beim XML.Datei die entsprechende Werte vorhanden sind, ansonsten werden sie initialisiert.
#################################################################################################################

function VariabelWert_Testen {

#Hier wird geprüft ob die Variablen vorhanden sind.
if(Test-Path $Emailpfad) {                    
$vars = Import-Clixml -Path $Emailpfad
} 
#Wenn nicht vorhanden, werden die Variablen initialisiert
else {                                       
$vars = @{
        Var1 = ""
        Var2 = ""
        Var3 = ""
        Var4 = ""
        Var5 = ""
    }
    }
    }



# Hier werden die nur Daten von der XML Datei importiet und den Werten zugeordnet.
##################################################################################

function Variabeln_Importieren {
$vars = Import-Clixml -Path $Emailpfad  #Pfad im gleichen Ordner

#Importierte Werte werden den entsprechenden Variablen zugeordnet
$to = $vars.Var1 
$subject = $vars.Var2
$body = $vars.Var3
$cc= $vars.Var4
$bcc= $vars.Var5

}


# Eine Funktion welche die Variabelwerte der E-Maildaten als Label anzeigt
##########################################################################

function Maildatenanzeige {
$TextMaildaten.Text = "Hier können Sie die E-Mail Vorlage anpassen.

Aktuelle Werte:

E-Mail-Empfänger:
$to 

CC:
$cc

BCC:
$bcc

Betreff: 
$subject

Inhalt:
$body" 

}


#########################################      FUNKTIONEN ENDE      ################################################################
####################################################################################################################################
####################################################################################################################################


#Pfadbestimmung der Stammdaten für Mail und Twitter
$Emailpfad = Join-Path -Path $PSScriptRoot -ChildPath "Maildaten.xml"
$Twitterpfad = Join-Path -Path $PSScriptRoot -ChildPath "Twitterdaten.xml"

VariabelWert_Testen  #Variablenzuordnung
Variabeln_Importieren

#########################################      BEGRÜSSUNGSMENÜ      ################################################################
####################################################################################################################################

# Begrüssungsfenster wird erstellt
$Startmenu = New-Object System.Windows.Forms.Form
$Startmenu.Text = "Auto-Mailer / Auto-Tweeter"
$Startmenu.Size = New-Object System.Drawing.Size(500,500)
$Startmenu.StartPosition = "CenterScreen"

# Hier wird ein Bild hinzugefügt.
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.Width = 180
$pictureBox.Height = 80
$pictureBox.Location = New-Object System.Drawing.Point(160,110)
$imagePath = Join-Path $PSScriptRoot "image.png"
$image = [System.Drawing.Image]::FromFile($imagePath)
$pictureBox.Image = $image

$Startmenu.Controls.Add($pictureBox)


# Die Tabelle am oberen Rand des Fensters wird erstellt
$TabelleStartmenu = New-Object System.Windows.Forms.TableLayoutPanel #Erstellung der Tabelle
$TabelleStartmenu.RowCount = 1 #Spaltenazahl
$TabelleStartmenu.BackColor = "#6B8E23" #Hintergrundfarbe
$TabelleStartmenu.Size = New-Object System.Drawing.Size(400, 90)  #Grösse 400 in Breite, 90 in Höhe
$TabelleStartmenu.Dock = [System.Windows.Forms.DockStyle]::Top  #Position Oberhalb des Fensters
$Startmenu.Controls.Add($TabelleStartmenu) #Einfügen der Tabelle beim Fenster "Startmenü


# Der Tabelle wird ein Text eingefügt.
$TitelStartmenu = New-Object System.Windows.Forms.Label #Es wird ein Label(Textfeld) eingefügt.
$TitelStartmenu.Text = "Auto-Mailer/ Auto-Tweeter" #Der Text welche angezeigt werden soll.
$TitelStartmenu.Font = New-Object System.Drawing.Font("Arial", 20) # Definition der Schriftgrösse und Schriftart
$TitelStartmenu.ForeColor = "White" #Schriftfarbe 
$TitelStartmenu.TextAlign = "MiddleCenter" #Text ist zentriert
$TitelStartmenu.AutoSize = $true #Automatische Anpassung des Textes
$TitelStartmenu.Padding = New-Object System.Windows.Forms.Padding(70,25,0,0) # Padding zum Parentobjekt angepasst, 70 von links 25 von oben
$TabelleStartmenu.Controls.Add($TitelStartmenu, 0, 0) #Einfügen des Textes zur Tabelle, definierte Positionierung


#Text, Label wird eingefügt und positioniert
$TextStartmenu = New-Object System.Windows.Forms.Label
$TextStartmenu.Text = "Dieses Programm kann eine gespeicherte E-Mail-Vorlage
an eine beliebige E-Mail-Adresse senden oder eine
gespeicherte Twitter-Vorlage auf Twitter posten."
$TextStartmenu.Font = New-Object System.Drawing.Font("Arial", 12)
$TextStartmenu.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
$TextStartmenu.AutoSize = $true
$TextStartmenu.Top = 200 #Position/Abstand in Pixels von Oben nach unten
$TextStartmenu.Left = 20 #Position/Abstand in Pixels von links nach rechts
$Startmenu.Controls.Add($TextStartmenu)


# Erstellung der Taste "Programm starten"
$Starttaste = New-Object System.Windows.Forms.Button
$Starttaste.Location = New-Object System.Drawing.Point(170, 280)
$Starttaste.Size = New-Object System.Drawing.Size(150, 50)
$Starttaste.Text = "Programm starten"
$Starttaste.Add_Click({ #Definiert was passiert, wenn die Taste gedrückt wird.

$Startmenu.Hide() #Der aktuelle Fenster wird ausgeblendet

#########################################      HAUPTMENÜ      ######################################################################
####################################################################################################################################
   
# Erstellung eines neuen Fensters
$Hauptmenu = New-Object System.Windows.Forms.Form
$Hauptmenu.Text = "Auto-Mailer / Auto-Tweeter" #Der Titel des Fensters
$Hauptmenu.Size = New-Object System.Drawing.Size(500,500) #Fenstergrösse wird fix definiert 500 x 500
$Hauptmenu.StartPosition = "CenterScreen" #Beim öffnen wird es im Zentrum des Bildschirms angezeigt.


# Die Tabelle am oberen Rand des Fensters wird erstellt
$TabelleHauptmenu = New-Object System.Windows.Forms.TableLayoutPanel
$TabelleHauptmenu.RowCount = 1
$TabelleHauptmenu.BackColor = "#6B8E23"
$TabelleHauptmenu.Size = New-Object System.Drawing.Size(400, 90)
$TabelleHauptmenu.Dock = [System.Windows.Forms.DockStyle]::Top
$Hauptmenu.Controls.Add($TabelleHauptmenu)

# Der Tabelle wird ein Titel eingefügt.
$TitelHauptmenu = New-Object System.Windows.Forms.Label
$TitelHauptmenu.Text = "Auto-Mailer/ Auto-Tweeter"
$TitelHauptmenu.Font = New-Object System.Drawing.Font("Arial", 20)
$TitelHauptmenu.ForeColor = "White"
$TitelHauptmenu.TextAlign = "MiddleCenter"
$TitelHauptmenu.AutoSize = $true
$TitelHauptmenu.Padding = New-Object System.Windows.Forms.Padding(70,25,0,0) 
$TabelleHauptmenu.Controls.Add($TitelHauptmenu, 0, 0)

#Text, Label wird eingefügt und positioniert
$TextHauptmenu = New-Object System.Windows.Forms.Label
$TextHauptmenu.Location = New-Object System.Drawing.Point(25, 160)
$TextHauptmenu.Size = New-Object System.Drawing.Size(400, 25)
$TextHauptmenu.Text = "Bitte wählen Sie eine Funktion aus:"
$TextHauptmenu.Font = New-Object System.Drawing.Font("Arial", 10)
$TextHauptmenu.ForeColor = "black"
$Hauptmenu.Controls.Add($TextHauptmenu)
$TextHauptmenu.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
$TextHauptmenu.AutoSize = $true
$TextHauptmenu.Top = 170
$TextHauptmenu.Left = 150

# Erstellung der Taste "Eine Email senden"
$Mailsendentaste = New-Object System.Windows.Forms.Button
$Mailsendentaste.Location = New-Object System.Drawing.Point(175, 220)
$Mailsendentaste.Size = New-Object System.Drawing.Size(150, 50)
$Mailsendentaste.Text = "Eine E-Mail senden"
$Mailsendentaste.Add_Click({

$Hauptmenu.Hide()      

#########################################      MAILMENÜ      #######################################################################
####################################################################################################################################

# Erstellung eines neuen Menüfensters
$Mailsendenmenu = New-Object System.Windows.Forms.Form
$Mailsendenmenu.Text = "Auto-Mailer"
$Mailsendenmenu.Size = New-Object System.Drawing.Size(500,500)
$Mailsendenmenu.StartPosition = "CenterScreen"

# Hier wird ein Bild hinzugefügt.
$pictureBox.Location = New-Object System.Drawing.Point(60,260)
$Mailsendenmenu.Controls.Add($pictureBox)


# Die Tabelle am oberen Rand des Fensters wird erstellt
$TabelleMailsenden = New-Object System.Windows.Forms.TableLayoutPanel
$TabelleMailsenden.RowCount = 1
$TabelleMailsenden.BackColor = "#6B8E23"
$TabelleMailsenden.Size = New-Object System.Drawing.Size(400, 90)
$TabelleMailsenden.Dock = [System.Windows.Forms.DockStyle]::Top
$Mailsendenmenu.Controls.Add($TabelleMailsenden)

# Der Tabelle wird ein Titel eingefügt.
$TitelMailsenden = New-Object System.Windows.Forms.Label
$TitelMailsenden.Text = "Auto-Mailer"
$TitelMailsenden.Font = New-Object System.Drawing.Font("Arial", 20)
$TitelMailsenden.ForeColor = "White"
$TitelMailsenden.TextAlign = "MiddleCenter"
$TitelMailsenden.AutoSize = $true
$TitelMailsenden.Padding = New-Object System.Windows.Forms.Padding(170,25,0,0) 
$TabelleMailsenden.Controls.Add($TitelMailsenden, 0, 0)

#Text, Label wird eingefügt und positioniert
$TextMailsenden = New-Object System.Windows.Forms.Label
$TextMailsenden.Location = New-Object System.Drawing.Point(25, 160)
$TextMailsenden.Size = New-Object System.Drawing.Size(400, 25)
$TextMailsenden.Text = "Hier können Sie Ihre gespeicherte E-Mail-Vorlage an 
eine gewünschte E-Mail-Adresse senden."
$TextMailsenden.Font = New-Object System.Drawing.Font("Arial", 10)
$TextMailsenden.ForeColor = "black"
$Mailsendenmenu.Controls.Add($TextMailsenden)
$TextMailsenden.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
$TextMailsenden.AutoSize = $true
$TextMailsenden.Top = 170
$TextMailsenden.Left = 50

# Erstellung der Taste "Eine Email senden"
$MailSendentaste = New-Object System.Windows.Forms.Button
$MailSendentaste.Location = New-Object System.Drawing.Point(310, 270)
$MailSendentaste.Size = New-Object System.Drawing.Size(150, 50)
$MailSendentaste.Text = "Eine E-Mail Senden"
$MailSendentaste.Add_Click({

$vars = Import-Clixml -Path $Emailpfad 
$to = $vars.Var1
$subject = $vars.Var2
$body = $vars.Var3
$cc= $vars.Var4
$bcc= $vars.Var5
  
$body = $body -replace "'", "''"  # Spezielle Charaktere werden ersetzt.

#Hier wird die Verknüpfung zum E-Mail erstellt, Quelle (https://stackoverflow.com/questions/33736528/make-powershell-v2-open-a-new-mail-message-from-mailto-link)
$Verknüpfung = New-Object System.Uri("mailto:$to")
$Verknüpfungshersteller = New-Object System.UriBuilder($Verknüpfung)
$Verknüpfungshersteller.Port = -1 #Damit kein Port dem Empfänger eingefügt wird
$Verknüpfungshersteller.Query = "subject=$subject&cc=$cc&bcc=$bcc&body=$body&importance=$importance&Disposition-Notification-To=$to" #Zuordnung von Variablen dem Wert des Linkes
$url = $Verknüpfungshersteller.ToString() #Als String

Start-Process $url #Hier wird der Link ausgeführt und öffnet dann die Email Applikation auf Windows

     })
#Erstellung der Zurücktaste
$Zurücktaste = New-Object System.Windows.Forms.Button 
$Zurücktaste.Location = New-Object System.Drawing.Point(310, 420)
$Zurücktaste.Size = New-Object System.Drawing.Size(75, 25)
$Zurücktaste.Text = "Zurück"
$Zurücktaste.Add_Click({
$Mailsendenmenu.Hide()        
$Hauptmenu.Show()
     })

#Erstellung der Beendentaste
$Beendentaste = New-Object System.Windows.Forms.Button
$Beendentaste.Location = New-Object System.Drawing.Point(390, 420)
$Beendentaste.Size = New-Object System.Drawing.Size(75, 25)
$Beendentaste.Text = "Beenden"
$Beendentaste.Add_Click({
$Mailsendenmenu.Close();
    })

#Anzeige Tasten im Mailsendemenü
$Mailsendenmenu.Controls.Add($MailSendentaste)
$Mailsendenmenu.Controls.Add($Zurücktaste)
$Mailsendenmenu.Controls.Add($Beendentaste)
$Mailsendenmenu.ShowDialog()

      
})
$Hauptmenu.Controls.Add($Mailsendentaste) #Anzeige Taste beim Hauptmenü

#Erstellung der Taste
$Tweetsendentaste = New-Object System.Windows.Forms.Button
$Tweetsendentaste.Location = New-Object System.Drawing.Point(175, 280)
$Tweetsendentaste.Size = New-Object System.Drawing.Size(150, 50)
$Tweetsendentaste.Text = "Einen Tweet posten"

$Tweetsendentaste.Add_Click({ #Die Aktion welche bei Klicken des "Einen Tweet posten" gemacht wird.

$Hauptmenu.Hide()   #Ausblenden des Hauptmenüs

#########################################      TWITTERMENÜ      ####################################################################
####################################################################################################################################
    
# Erstellung eines neuen Fensters
$Tweetsendenmenu = New-Object System.Windows.Forms.Form
$Tweetsendenmenu.Text = "Auto-Tweeter"
$Tweetsendenmenu.Size = New-Object System.Drawing.Size(500,500)
$Tweetsendenmenu.StartPosition = "CenterScreen"

# Hier wird ein Bild  (Picture box) hinzugefügt.
$pictureBox.Location = New-Object System.Drawing.Point(45,300)
$Tweetsendenmenu.Controls.Add($pictureBox)

#Hier wird die Tabelle hinzugefügt.
$TabelleTweetsenden = New-Object System.Windows.Forms.TableLayoutPanel
$TabelleTweetsenden.RowCount = 1
$TabelleTweetsenden.BackColor = "#6B8E23"
$TabelleTweetsenden.Size = New-Object System.Drawing.Size(400, 90)
$TabelleTweetsenden.Dock = [System.Windows.Forms.DockStyle]::Top
$Tweetsendenmenu.Controls.Add($TabelleTweetsenden)

#Hier wird der Text zur Tabelle hinzugefügt
$TitelTweetsenden = New-Object System.Windows.Forms.Label
$TitelTweetsenden.Text = "Auto-Tweeter"
$TitelTweetsenden.Font = New-Object System.Drawing.Font("Arial", 20)
$TitelTweetsenden.ForeColor = "White"
$TitelTweetsenden.TextAlign = "MiddleCenter"
$TitelTweetsenden.AutoSize = $true
$TitelTweetsenden.Padding = New-Object System.Windows.Forms.Padding(170,25,0,0) 
$TabelleTweetsenden.Controls.Add($TitelTweetsenden, 0, 0)

#Hier wird Label/Text erstellt
$TextTweetsenden = New-Object System.Windows.Forms.Label
$TextTweetsenden.Location = New-Object System.Drawing.Point(25, 160)
$TextTweetsenden.MaximumSize = New-Object System.Drawing.Size(200, 0)
$Twitterpost = Import-Clixml -Path $Twitterpfad  #Hier wird die Twitterpost variable aktualisiert, damit es up-to-date anzeigen kann.
$TextTweetsenden.Text = "Hier können Sie einen Tweet posten.

Aktuelle Twitter-Vorlage: 

$Twitterpost" #Anzeige des aktuellen Wert des Variable $Twitterpost
$TextTweetsenden.Font = New-Object System.Drawing.Font("Arial", 10)
$TextTweetsenden.ForeColor = "black"
$Tweetsendenmenu.Controls.Add($TextTweetsenden)
$TextTweetsenden.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
$TextTweetsenden.AutoSize = $true
$TextTweetsenden.Top = 150
$TextTweetsenden.Left = 50

$Twitterposten = New-Object System.Windows.Forms.Button
$Twitterposten.Location = New-Object System.Drawing.Point(310, 270)
$Twitterposten.Size = New-Object System.Drawing.Size(150, 50)
$Twitterposten.Text = "Jetzt auf Twitter posten"
$Twitterposten.Add_Click({
$Tweetsendenmenu.Close();
# Erstellung des Linkes zum Twitter
$url = "https://twitter.com/intent/tweet?text=$Twitterpost"

#Öffnung des Linkes zum Twitter
Start-Process $url
    })

$ZurückTweetsenden = New-Object System.Windows.Forms.Button
$ZurückTweetsenden.Location = New-Object System.Drawing.Point(310, 420)
$ZurückTweetsenden.Size = New-Object System.Drawing.Size(75, 25)
$ZurückTweetsenden.Text = "Zurück"
$ZurückTweetsenden.Add_Click({
$Tweetsendenmenu.Hide()         #Fenster ausblenden
$Hauptmenu.Show()      # Anzeige Zielfenster
     })

$Beendentaste = New-Object System.Windows.Forms.Button
$Beendentaste.Location = New-Object System.Drawing.Point(390, 420)
$Beendentaste.Size = New-Object System.Drawing.Size(75, 25)
$Beendentaste.Text = "Beenden"
$Beendentaste.Add_Click({
$Tweetsendenmenu.Close();
    })


$Tweetsendenmenu.Controls.Add($Twitterposten)
$Tweetsendenmenu.Controls.Add($ZurückTweetsenden)
$Tweetsendenmenu.Controls.Add($Beendentaste)
$Tweetsendenmenu.ShowDialog()
})

$Hauptmenu.Controls.Add($Tweetsendentaste)

#Erstellung neuer Taste
$Datenanpassentaste = New-Object System.Windows.Forms.Button
$Datenanpassentaste.Location = New-Object System.Drawing.Point(175, 340)
$Datenanpassentaste.Size = New-Object System.Drawing.Size(150, 50)
$Datenanpassentaste.Text = "Daten anpassen"

$Datenanpassentaste.Add_Click({
$Hauptmenu.Hide()

#########################################      DATEN ANPASSEN      #################################################################
####################################################################################################################################

$Datenanpassenmenu = New-Object System.Windows.Forms.Form
$Datenanpassenmenu.Text = "Daten anpassen"
$Datenanpassenmenu.Size = New-Object System.Drawing.Size(500,500)
$Datenanpassenmenu.StartPosition = "CenterScreen"

#Hier wird die Tabelle hinzugefügt.
$TabelleDatenanpassen = New-Object System.Windows.Forms.TableLayoutPanel
$TabelleDatenanpassen.RowCount = 1
$TabelleDatenanpassen.BackColor = "#6B8E23"
$TabelleDatenanpassen.Size = New-Object System.Drawing.Size(400, 90)
$TabelleDatenanpassen.Dock = [System.Windows.Forms.DockStyle]::Top
$Datenanpassenmenu.Controls.Add($TabelleDatenanpassen)

#Hier wird der Text zur Tabelle hinzugefügt
$TitelDatenanpassen = New-Object System.Windows.Forms.Label
$TitelDatenanpassen.Text = "Daten anpassen"
$TitelDatenanpassen.Font = New-Object System.Drawing.Font("Arial", 20)
$TitelDatenanpassen.ForeColor = "White"
$TitelDatenanpassen.TextAlign = "MiddleCenter"
$TitelDatenanpassen.AutoSize = $true
$TitelDatenanpassen.Padding = New-Object System.Windows.Forms.Padding(140,25,0,0) 
$TabelleDatenanpassen.Controls.Add($TitelDatenanpassen, 0, 0)

$TextDatenanpassen = New-Object System.Windows.Forms.Label
$TextDatenanpassen.Location = New-Object System.Drawing.Point(25, 160)
$TextDatenanpassen.Size = New-Object System.Drawing.Size(400, 25)
$TextDatenanpassen.Text = "Hier können Sie nun die Daten anpassen.
Bitte wählen Sie nun welche Daten Sie anpassen wollen."
$TextDatenanpassen.Font = New-Object System.Drawing.Font("Arial", 10)
$TextDatenanpassen.ForeColor = "black"
$Datenanpassenmenu.Controls.Add($TextDatenanpassen)
$TextDatenanpassen.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
$TextDatenanpassen.AutoSize = $true
$TextDatenanpassen.Top = 130
$TextDatenanpassen.Left = 50

#Hier wird die Taste hinzugefügt.
$Twitteranpassentaste = New-Object System.Windows.Forms.Button
$Twitteranpassentaste.Location = New-Object System.Drawing.Point(170, 320)
$Twitteranpassentaste.Size = New-Object System.Drawing.Size(150, 50)
$Twitteranpassentaste.Text = "Twitter-Vorlage anpassen"
$Twitteranpassentaste.Add_Click({
$Datenanpassenmenu.Hide();
$Twitteranpassenmenu = New-Object System.Windows.Forms.Form
$Twitteranpassenmenu.Text = "Twitter Vorlage anpassen"
$Twitteranpassenmenu.AutoSize = $true
$Twitteranpassenmenu.Size = New-Object System.Drawing.Size(500,500)
$Twitteranpassenmenu.StartPosition = "CenterScreen"

#Hier wird die Tabelle hinzugefügt.
$TabelleTwittervorlage = New-Object System.Windows.Forms.TableLayoutPanel
$TabelleTwittervorlage.RowCount = 1
$TabelleTwittervorlage.BackColor = "#6B8E23"
$TabelleTwittervorlage.Size = New-Object System.Drawing.Size(400, 90)
$TabelleTwittervorlage.Dock = [System.Windows.Forms.DockStyle]::Top
$Twitteranpassenmenu.Controls.Add($TabelleTwittervorlage)

#Hier wird der Text zur Tabelle hinzugefügt
$TitelTwittervorlage = New-Object System.Windows.Forms.Label
$TitelTwittervorlage.Text = "Twitter Vorlage anpassen"
$TitelTwittervorlage.Font = New-Object System.Drawing.Font("Arial", 20)
$TitelTwittervorlage.ForeColor = "White"
$TitelTwittervorlage.TextAlign = "MiddleCenter"
$TitelTwittervorlage.AutoSize = $true
$TitelTwittervorlage.Padding = New-Object System.Windows.Forms.Padding(40,25,0,0) 
$TabelleTwittervorlage.Controls.Add($TitelTwittervorlage, 0, 0)

$Texttwittervorlage = New-Object System.Windows.Forms.Label
$Texttwittervorlage.Location = New-Object System.Drawing.Point(25, 160)
$Texttwittervorlage.AutoSize = $false
$Texttwittervorlage.MaximumSize = New-Object System.Drawing.Size(200, 0)
$Twitterpost = Import-Clixml -Path $Twitterpfad #abruf der aktuellen Datei
$Texttwittervorlage.Text = "Hier können Sie die Twitter Vorlage anpassen.

Aktueller Tweet:

$Twitterpost" #Anzeige aktuelle Wert der Variable
$Texttwittervorlage.Font = New-Object System.Drawing.Font("Arial", 10)
$Texttwittervorlage.ForeColor = "black"
$Twitteranpassenmenu.Controls.Add($Texttwittervorlage)
$Texttwittervorlage.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
$Texttwittervorlage.AutoSize = $true
$Texttwittervorlage.Top = 130
$Texttwittervorlage.Left = 50

Add-Type -AssemblyName System.Windows.Forms

$EingabefensterTweet = New-Object System.Windows.Forms.Form
$EingabefensterTweet.Text = "Tweet Mitteilung eingeben"
$EingabefensterTweet.StartPosition = "CenterScreen"

$Tweetanpassenbutton = New-Object System.Windows.Forms.Button
$Tweetanpassenbutton.Location = New-Object System.Drawing.Point(310, 200)
$Tweetanpassenbutton.Size = New-Object System.Drawing.Size(150, 50)
$Tweetanpassenbutton.Text = "Tweet-Vorlage ändern"
$Tweetanpassenbutton.Add_Click({

$Tweeteingabefeld = New-Object System.Windows.Forms.TextBox
$Tweeteingabefeld.Location = New-Object System.Drawing.Point(10, 40)
$Tweeteingabefeld.Size = New-Object System.Drawing.Size(100, 23)
$EingabefensterTweet.Controls.Add($Tweeteingabefeld)
    
$Tweetaenderntaste = New-Object System.Windows.Forms.Button
$Tweetaenderntaste.Text = "Speichern"
$Tweetaenderntaste.Location = New-Object System.Drawing.Point(120, 40)
$Tweetaenderntaste.Size = New-Object System.Drawing.Size(75, 23)
$Tweetaenderntaste.Add_Click({
$Twitterpost = $Tweeteingabefeld.Text

$Twitterpost | Export-Clixml -Path $Twitterpfad
$Twitterpost = Import-Clixml -Path $Twitterpfad


$Texttwittervorlage.Text = "Hier können Sie die Twitter Vorlage anpassen.

Aktueller Tweet:

$Twitterpost"
$EingabefensterTweet.DialogResult = [System.Windows.Forms.DialogResult]::OK
})
$EingabefensterTweet.Controls.Add($Tweetaenderntaste)
$EingabefensterTweet.ShowDialog() | Out-Null

})

$Beendentaste = New-Object System.Windows.Forms.Button
$Beendentaste.Location = New-Object System.Drawing.Point(390, 420)
$Beendentaste.Size = New-Object System.Drawing.Size(75, 25)
$Beendentaste.Text = "Beenden"
$Beendentaste.Add_Click({
$Twitteranpassenmenu.Close();
})


$ZurückTwitterAnpassen = New-Object System.Windows.Forms.Button
$ZurückTwitterAnpassen.Location = New-Object System.Drawing.Point(310, 420)
$ZurückTwitterAnpassen.Size = New-Object System.Drawing.Size(75, 25)
$ZurückTwitterAnpassen.Text = "Zurück"
$ZurückTwitterAnpassen.Add_Click({
$Twitteranpassenmenu.Hide()         # Twitteranpassenmenü Fenster wird nicht angezeigt
$Datenanpassenmenu.Show()      # Datenanpassenmenü Fenster wird angezeigt
})
 
$Twitteranpassenmenu.Controls.Add($Beendentaste)
$Twitteranpassenmenu.Controls.Add($ZurückTwitterAnpassen)
$Twitteranpassenmenu.Controls.Add($Tweetanpassenbutton)
$Twitteranpassenmenu.ShowDialog()
})

$Mailanpassentaste = New-Object System.Windows.Forms.Button
$Mailanpassentaste.Location = New-Object System.Drawing.Point(170, 250)
$Mailanpassentaste.Size = New-Object System.Drawing.Size(150, 50)
$Mailanpassentaste.Text = "E-Mail-Vorlage anpassen"
$Mailanpassentaste.Add_Click({
$Datenanpassenmenu.Hide();
  
$Mailanpassenmenu = New-Object System.Windows.Forms.Form
$Mailanpassenmenu.Text = "E-Mail-Vorlage Anpassen"
$Mailanpassenmenu.Size = New-Object System.Drawing.Size(500)
$Mailanpassenmenu.AutoSize = $true
$Mailanpassenmenu.StartPosition = "CenterScreen"

#Hier wird die Tabelle hinzugefügt.
$TabelleEmailvorlage = New-Object System.Windows.Forms.TableLayoutPanel
$TabelleEmailvorlage.RowCount = 1
$TabelleEmailvorlage.BackColor = "#6B8E23"
$TabelleEmailvorlage.Size = New-Object System.Drawing.Size(400, 90)
$TabelleEmailvorlage.Dock = [System.Windows.Forms.DockStyle]::Top
$Mailanpassenmenu.Controls.Add($TabelleEmailvorlage)

#Hier wird der Text zur Tabelle hinzugefügt
$TitelEmailvorlage = New-Object System.Windows.Forms.Label
$TitelEmailvorlage.Text = "E-Mail-Vorlage anpassen"
$TitelEmailvorlage.Font = New-Object System.Drawing.Font("Arial", 20)
$TitelEmailvorlage.ForeColor = "White"
$TitelEmailvorlage.TextAlign = "MiddleCenter"
$TitelEmailvorlage.AutoSize = $true
$TitelEmailvorlage.Padding = New-Object System.Windows.Forms.Padding(90,25,0,0) # Erhöhe den oberen Padding-Wert um 2 Pixel
$TabelleEmailvorlage.Controls.Add($TitelEmailvorlage, 0, 0)

$TextMaildaten = New-Object System.Windows.Forms.Label
$TextMaildaten.Location = New-Object System.Drawing.Point(25, 160)
$TextMaildaten.Size = New-Object System.Drawing.Size(100, 100)

VariabelWert_Testen
$vars = Import-Clixml -Path $Emailpfad

$to = $vars.Var1 #Zuweisung der Daten des Importierten Daten
$subject = $vars.Var2
$body = $vars.Var3
$cc= $vars.Var4
$bcc= $vars.Var5

Maildatenanzeige #Ausführung Funktion Maildatenanzeige

$TextMaildaten.Font = New-Object System.Drawing.Font("Arial", 10)
$TextMaildaten.ForeColor = "black"
$Mailanpassenmenu.Controls.Add($TextMaildaten)
$TextMaildaten.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
$TextMaildaten.AutoSize = $true
$TextMaildaten.Top = 130
$TextMaildaten.Left = 50
$TextMaildaten.MaximumSize = New-Object System.Drawing.Size(250, 0)

$ZurücktasteEmailVorlage = New-Object System.Windows.Forms.Button
$ZurücktasteEmailVorlage.Location = New-Object System.Drawing.Point(310, 420)
$ZurücktasteEmailVorlage.Size = New-Object System.Drawing.Size(75, 25)
$ZurücktasteEmailVorlage.Text = "Zurück"
$ZurücktasteEmailVorlage.Add_Click({
$Mailanpassenmenu.Hide()        
$Datenanpassenmenu.Show()      
})

$Beendentaste = New-Object System.Windows.Forms.Button
$Beendentaste.Location = New-Object System.Drawing.Point(390, 420)
$Beendentaste.Size = New-Object System.Drawing.Size(75, 25)
$Beendentaste.Text = "Beenden"
$Beendentaste.Add_Click({
$Mailanpassenmenu.Close();
})

#########################################      EINGABEFELDER      ###3##############################################################
####################################################################################################################################

Add-Type -AssemblyName System.Windows.Forms
$EingabeEmpfaenger = New-Object System.Windows.Forms.Form
$EingabeEmpfaenger.Text = "Empfänger eingeben"
$EingabeEmpfaenger.StartPosition = "CenterScreen"

$Empfaengertaste = New-Object System.Windows.Forms.Button
$Empfaengertaste.Location = New-Object System.Drawing.Point(345, 180)
$Empfaengertaste.Size = New-Object System.Drawing.Size(120, 40)
$Empfaengertaste.Text = "Empfänger ändern"
$Empfaengertaste.Add_Click({

$Eingabefeldempfaenger = New-Object System.Windows.Forms.TextBox
$Eingabefeldempfaenger.Location = New-Object System.Drawing.Point(10, 40)
$Eingabefeldempfaenger.Size = New-Object System.Drawing.Size(100, 23)
$EingabeEmpfaenger.Controls.Add($Eingabefeldempfaenger)
    
$Sendentasteempfaenger = New-Object System.Windows.Forms.Button
$Sendentasteempfaenger.Text = "Speichern"
$Sendentasteempfaenger.Location = New-Object System.Drawing.Point(120, 40)
$Sendentasteempfaenger.Size = New-Object System.Drawing.Size(75, 23)
$Sendentasteempfaenger.Add_Click({
$to = $Eingabefeldempfaenger.Text #Benutzer Eingabe wird der Variable zugewiesen

VariabelWert_Testen
$vars.Var1 = $to


$vars | Export-Clixml -Path $Emailpfad

$vars = Import-Clixml -Path $Emailpfad


$to = $vars.Var1
$subject = $vars.Var2
$body = $vars.Var3
$cc= $vars.Var4
$bcc= $vars.Var5

Maildatenanzeige

   
$EingabeEmpfaenger.DialogResult = [System.Windows.Forms.DialogResult]::OK
})
$EingabeEmpfaenger.Controls.Add($Sendentasteempfaenger)
$EingabeEmpfaenger.ShowDialog() | Out-Null

})

Add-Type -AssemblyName System.Windows.Forms

$EingabeCC = New-Object System.Windows.Forms.Form
$EingabeCC.Text = "CC eingeben"
$EingabeCC.StartPosition = "CenterScreen"

$CCTaste = New-Object System.Windows.Forms.Button
$CCTaste.Location = New-Object System.Drawing.Point(345, 225)
$CCTaste.Size = New-Object System.Drawing.Size(120, 40)
$CCTaste.Text = "CC ändern"
$CCTaste.Add_Click({

$EingabefeldCC = New-Object System.Windows.Forms.TextBox
$EingabefeldCC.Location = New-Object System.Drawing.Point(10, 40)
$EingabefeldCC.Size = New-Object System.Drawing.Size(100, 23)
$EingabeCC.Controls.Add($EingabefeldCC)
    
$SendentasteCC = New-Object System.Windows.Forms.Button
$SendentasteCC.Text = "Speichern"
$SendentasteCC.Location = New-Object System.Drawing.Point(120, 40)
$SendentasteCC.Size = New-Object System.Drawing.Size(75, 23)
$SendentasteCC.Add_Click({
$cc = $EingabefeldCC.Text #Benutzer Eingabe wird der Variable zugewiesen
   
VariabelWert_Testen
$vars.Var4 = $cc

$vars | Export-Clixml -Path $Emailpfad

$vars = Import-Clixml -Path $Emailpfad

$to = $vars.Var1
$subject = $vars.Var2
$body = $vars.Var3
$cc= $vars.Var4
$bcc= $vars.Var5

Maildatenanzeige

$EingabeCC.DialogResult = [System.Windows.Forms.DialogResult]::OK
})
$EingabeCC.Controls.Add($SendentasteCC)
$EingabeCC.ShowDialog() | Out-Null

})
Add-Type -AssemblyName System.Windows.Forms

$EingabeBCC = New-Object System.Windows.Forms.Form
$EingabeBCC.Text = "BCC eingeben"
$EingabeBCC.StartPosition = "CenterScreen"

$BCCTaste = New-Object System.Windows.Forms.Button
$BCCTaste.Location = New-Object System.Drawing.Point(345, 270)
$BCCTaste.Size = New-Object System.Drawing.Size(120, 40)
$BCCTaste.Text = "BCC ändern"
$BCCTaste.Add_Click({

$EingabefeldBCC = New-Object System.Windows.Forms.TextBox
$EingabefeldBCC.Location = New-Object System.Drawing.Point(10, 40)
$EingabefeldBCC.Size = New-Object System.Drawing.Size(100, 23)
$EingabeBCC.Controls.Add($EingabefeldBCC)
    
$SendentasteBCC = New-Object System.Windows.Forms.Button
$SendentasteBCC.Text = "Speichern"
$SendentasteBCC.Location = New-Object System.Drawing.Point(120, 40)
$SendentasteBCC.Size = New-Object System.Drawing.Size(75, 23)
$SendentasteBCC.Add_Click({
$bcc = $EingabefeldBCC.Text #Benutzer Eingabe wird der Variable zugewiesen
   
VariabelWert_Testen
$vars.Var5 = $bcc

$vars | Export-Clixml -Path $Emailpfad

$vars = Import-Clixml -Path $Emailpfad

$to = $vars.Var1
$subject = $vars.Var2
$body = $vars.Var3
$cc= $vars.Var4
$bcc= $vars.Var5

Maildatenanzeige
    
$EingabeBCC.DialogResult = [System.Windows.Forms.DialogResult]::OK
})
$EingabeBCC.Controls.Add($SendentasteBCC)
$EingabeBCC.ShowDialog() | Out-Null
})

Add-Type -AssemblyName System.Windows.Forms
$EingabeBetreff = New-Object System.Windows.Forms.Form
$EingabeBetreff.Text = "Betreff eingeben"
$EingabeBetreff.StartPosition = "CenterScreen"

$Betrefftaste = New-Object System.Windows.Forms.Button
$Betrefftaste.Location = New-Object System.Drawing.Point(345, 315)
$Betrefftaste.Size = New-Object System.Drawing.Size(120, 40)
$Betrefftaste.Text = "Betreff ändern"
$Betrefftaste.Add_Click({

$Eingabefeldbetreff = New-Object System.Windows.Forms.TextBox
$Eingabefeldbetreff.Location = New-Object System.Drawing.Point(10, 40)
$Eingabefeldbetreff.Size = New-Object System.Drawing.Size(100, 23)
$EingabeBetreff.Controls.Add($Eingabefeldbetreff)
    
$Sendentastebetreff = New-Object System.Windows.Forms.Button
$Sendentastebetreff.Text = "Speichern"
$Sendentastebetreff.Location = New-Object System.Drawing.Point(120, 40)
$Sendentastebetreff.Size = New-Object System.Drawing.Size(75, 23)
$Sendentastebetreff.Add_Click({
$subject = $Eingabefeldbetreff.Text
    
VariabelWert_Testen
$vars.Var2 = $subject

$vars | Export-Clixml -Path $Emailpfad
$vars = Import-Clixml -Path $Emailpfad

$to = $vars.Var1
$subject = $vars.Var2
$body = $vars.Var3
$cc= $vars.Var4
$bcc= $vars.Var5

Maildatenanzeige  
    
$EingabeBetreff.DialogResult = [System.Windows.Forms.DialogResult]::OK
})
$EingabeBetreff.Controls.Add($Sendentastebetreff)
$EingabeBetreff.ShowDialog() | Out-Null
})

Add-Type -AssemblyName System.Windows.Forms

$EingabeInhalt = New-Object System.Windows.Forms.Form
$EingabeInhalt.Text = "E-Mail Inhalt eingeben"
$EingabeInhalt.StartPosition = "CenterScreen"

$Inhaltstaste = New-Object System.Windows.Forms.Button
$Inhaltstaste.Location = New-Object System.Drawing.Point(345, 360)
$Inhaltstaste.Size = New-Object System.Drawing.Size(120, 40)
$Inhaltstaste.Text = "E-Mail Inhalt ändern"
$Inhaltstaste.Add_Click({

$Eingabefeldinhalt = New-Object System.Windows.Forms.TextBox
$Eingabefeldinhalt.Location = New-Object System.Drawing.Point(10, 40)
$Eingabefeldinhalt.Size = New-Object System.Drawing.Size(100, 23)
$EingabeInhalt.Controls.Add($Eingabefeldinhalt)
    
$Sendentasteinhalt = New-Object System.Windows.Forms.Button
$Sendentasteinhalt.Text = "Speichern"
$Sendentasteinhalt.Location = New-Object System.Drawing.Point(120, 40)
$Sendentasteinhalt.Size = New-Object System.Drawing.Size(75, 23)
$Sendentasteinhalt.Add_Click({
$body = $Eingabefeldinhalt.Text
  
VariabelWert_Testen #Hier wird getestet ob Werte zugewiesen, ansonsten werden die Daten initialisiert, dass ist nötig damit die aktuellen Daten nicht überschrieben werden.
$vars.Var3 = $body

$vars | Export-Clixml -Path $Emailpfad
$vars = Import-Clixml -Path $Emailpfad

$to = $vars.Var1
$subject = $vars.Var2
$body = $vars.Var3
$cc= $vars.Var4
$bcc= $vars.Var5

Maildatenanzeige  
    
$EingabeInhalt.DialogResult = [System.Windows.Forms.DialogResult]::OK
})

$EingabeInhalt.Controls.Add($Sendentasteinhalt)
$EingabeInhalt.ShowDialog() | Out-Null
})

$Mailanpassenmenu.Controls.Add($ZurücktasteEmailVorlage)
$Mailanpassenmenu.Controls.Add($Beendentaste)
$Mailanpassenmenu.Controls.Add($Empfaengertaste)
$Mailanpassenmenu.Controls.Add($CCTaste)
$Mailanpassenmenu.Controls.Add($BCCTaste)
$Mailanpassenmenu.Controls.Add($Betrefftaste)
$Mailanpassenmenu.Controls.Add($Inhaltstaste)  

$Mailanpassenmenu.ShowDialog()
})

$Zurücktaste = New-Object System.Windows.Forms.Button
$Zurücktaste.Location = New-Object System.Drawing.Point(310, 420)
$Zurücktaste.Size = New-Object System.Drawing.Size(75, 25)
$Zurücktaste.Text = "Zurück"
$Zurücktaste.Add_Click({
   
$Datenanpassenmenu.Hide()       
$Hauptmenu.Show()      
})

$Beendentaste = New-Object System.Windows.Forms.Button
$Beendentaste.Location = New-Object System.Drawing.Point(390, 420)
$Beendentaste.Size = New-Object System.Drawing.Size(75, 25)
$Beendentaste.Text = "Beenden"
$Beendentaste.Add_Click({
$Datenanpassenmenu.Close();
})
   
$Datenanpassenmenu.Controls.Add($Twitteranpassentaste)
$Datenanpassenmenu.Controls.Add($Mailanpassentaste)
$Datenanpassenmenu.Controls.Add($Zurücktaste)
$Datenanpassenmenu.Controls.Add($Beendentaste)

$Datenanpassenmenu.ShowDialog()
})

$Hauptmenu.Controls.Add($Datenanpassentaste)
$Hauptmenu.ShowDialog()  
})

$Startmenu.Controls.Add($Starttaste)

# Erstellt Button für "Beenden"
$Hauptbeendentaste = New-Object System.Windows.Forms.Button
$Hauptbeendentaste.Location = New-Object System.Drawing.Point(170, 350)
$Hauptbeendentaste.Size = New-Object System.Drawing.Size(150, 50)
$Hauptbeendentaste.Text = "Beenden"
$Hauptbeendentaste.Add_Click({
$Startmenu.Close()
})
$Startmenu.Controls.Add($Hauptbeendentaste)

# Zeigt das Fenster an
$Startmenu.ShowDialog() | Out-Null







                     #********************************************************#
                     #********************************************************#
#************************** P R O G R A M M     E N D E ************************************************#
#*******************************************************************************************************#
#*******************************************************************************************************#
#************************** A U T O - M A I L E R / A U T O - T W I T T E R ****************************#
#*******************************************************************************************************#
#**************************  A U T H O R -  B A R I S  O E Z C A N *************************************#
#*******************************************************************************************************#
#*******************************************************************************************************#
#*******************************************************************************************************#
               
