Set IE = CreateObject("InternetExplorer.Application")

IE.Visible = True

CreateObject("WScript.Shell").AppActivate "Internet Explorer"

IE.navigate "http://REDACTED.REDACTED.com:7777/jde/E1Menu.maf"

While IE.Busy

    WScript.Sleep 50

Wend

Set ipf = IE.document.all.User

ipf.Value = "MY_USERNAME"

Set ipf = IE.document.all.Password

ipf.Value = "MY_PASSWORD"

Set oInputs = IE.document.getElementsByTagName("input") 'SIGN ON BUTTON

        For Each elm In oInputs

            If elm.value = "Sign In" Then

                elm.Click

                Exit For

            End If

Next

Set oShell = CreateObject("WScript.Shell")

oShell.SendKeys "% "

While IE.Busy

    WScript.Sleep 1500

Wend

oShell.SendKeys "x"

WScript.Sleep 2000

Set oInputs = IE.document.getElementsByTagName("img") 'FAVORITES

        For Each elm In oInputs

            If elm.alt = "Favorites" Then

                elm.Click

                Exit For

            End If

Next

Wscript.Sleep 1000

Do While IE.Busy or IE.ReadyState <> 4: Wscript.Sleep 100: Loop

Set oInputs = IE.document.getElementsByTagName("a") 'CUSTOMER ACTIVITY LOG

               For Each elm In oInputs

                   If elm.title = "Application: P03B31, Form: W03B31A" Then

                             elm.Click

                              Exit For

            End If

Next

Wscript.Sleep 100

Set oInputs = IE.document.getElementsByTagName("a") 'CUSTOMER SERVICE INQUIRY

               For Each elm In oInputs

                   If elm.title = "Application: P4210, Form: W4210E, Version: REDACTED" Then

                             elm.Click

                              Exit For

            End If

Next

Wscript.Sleep 100

Set oInputs = IE.document.getElementsByTagName("a") 'ADDRESS BOOK

               For Each elm In oInputs

                   If elm.title = "Application: P01012, Form: W01012B, Version: REDACTED" Then

                             elm.Click

                              Exit For

            End If

Next

Wscript.Sleep 100

Set oInputs = IE.document.getElementsByTagName("a") 'CUSTOMER LEDGER INQUIRY

               For Each elm In oInputs

                   If elm.title = "Application: P03B2002, Form: W03B2002A, Version: REDACTED" Then

                             elm.Click

                              Exit For

            End If

Next

Wscript.Sleep 100

Set oInputs = IE.document.getElementsByTagName("a") 'RELEASE ORDERS ON HOLD

               For Each elm In oInputs

                   If elm.title = "Application: P43070, Form: W43070A, Version: REDACTED" Then

                             elm.Click

                              Exit For

            End If

Next

Wscript.Sleep 150

Do While IE.Busy or IE.ReadyState <> 4: Wscript.Sleep 150: Loop

Set WshShell = WScript.CreateObject("WScript.Shell")                    

WshShell.SendKeys "{TAB}"

WScript.Sleep 100

WshShell.SendKeys "{TAB}"

WScript.Sleep 100

WshShell.SendKeys "{TAB}"

WScript.Sleep 100

WshShell.SendKeys "{TAB}"

WScript.Sleep 100

WshShell.SendKeys "{TAB}"

WScript.Sleep 100

WshShell.SendKeys "{TAB}"

WScript.Sleep 100

WshShell.SendKeys "<=15"

WScript.Sleep 500

WshShell.SendKeys "{ENTER}"

WScript.Sleep 2500
