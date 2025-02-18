# Zorg dat de Active Directory-module geladen is
Import-Module ActiveDirectory

# Haal alle servers op uit de specifieke OUs in AD
$OU1 = "OU=Servers,DC=Novoferm,DC=info"
$OU2 = "OU=Domain Controllers,DC=Novoferm,DC=info"

# Query Active Directory voor beide OUs en combineer de resultaten
$Servers = @()
$Servers += Get-ADComputer -Filter * -SearchBase $OU1 | Select-Object -ExpandProperty Name
$Servers += Get-ADComputer -Filter * -SearchBase $OU2 | Select-Object -ExpandProperty Name

# Sorteer de servers alfabetisch
$Servers = $Servers | Sort-Object

# Debug: Controleer of servers correct zijn opgehaald
Write-Host "Gevonden servers: $($Servers -join ', ')" -ForegroundColor Cyan

# Lege arrays voor resultaten en niet-bereikbare servers
$Results = @()
$FailedServers = @()

# Loop door alle servers
foreach ($Server in $Servers) {
    try {
        # Test eerst of de server online is (ping)
        if (-not (Test-Connection -ComputerName $Server -Count 2 -Quiet)) {
            Write-Host "❌ Server $Server is niet bereikbaar via netwerk (ping mislukt)." -ForegroundColor Red
            $FailedServers += $Server
            continue  # Ga verder met de volgende server
        }

        # Haal OS en hardware-informatie op via WMI
        $OS = Get-CimInstance -ComputerName $Server -ClassName Win32_OperatingSystem -ErrorAction Stop
        $ComputerSystem = Get-CimInstance -ComputerName $Server -ClassName Win32_ComputerSystem -ErrorAction Stop
        $BIOS = Get-CimInstance -ComputerName $Server -ClassName Win32_BIOS -ErrorAction Stop
        $CPU = (Get-CimInstance -ComputerName $Server -ClassName Win32_Processor -ErrorAction Stop | Select-Object -First 1).Name

        # Alleen het primaire IPv4-adres tonen
        $IPConfig = Get-WmiObject -ComputerName $Server -Class Win32_NetworkAdapterConfiguration -ErrorAction Stop | Where-Object { $_.IPEnabled -and $_.IPAddress -match '^\d{1,3}(\.\d{1,3}){3}$' }
        $IPAdres = if ($IPConfig) { $IPConfig.IPAddress[0] } else { "Geen IPv4-adres" }

        # Haal alleen de meest recente geïnstalleerde update op
        try {
            $LatestUpdate = Get-HotFix -ComputerName $Server -ErrorAction Stop | Sort-Object InstalledOn -Descending | Select-Object -First 1  
        } catch {
            Write-Host "⚠️ Kon KB-updates niet ophalen voor $Server." -ForegroundColor Yellow
            $LatestUpdate = $null
        }
        $LaatsteUpdateTekst = if ($LatestUpdate) { "$($LatestUpdate.HotFixID) ($($LatestUpdate.InstalledOn))" } else { "Geen updates gevonden" }

        # Uptime en laatste reboot ophalen
        $LastBootTime = $OS.LastBootUpTime
        $Uptime = (Get-Date) - $LastBootTime

        # Resultaat opslaan
        $Results += [PSCustomObject]@{
            "Servernaam"       = $Server
            "OS"               = "$($OS.Caption) ($($OS.Version))"
            "Buildnummer"      = $OS.BuildNumber
            "Fabrikant"        = $ComputerSystem.Manufacturer
            "Model"            = $ComputerSystem.Model
            "Processor"        = $CPU
            "Geheugen (GB)"    = [math]::Round($ComputerSystem.TotalPhysicalMemory / 1GB, 2)
            "IP Adres"         = $IPAdres
            "Laatste reboot"   = $LastBootTime
            "Uptime (Dagen)"   = [math]::Round($Uptime.TotalDays, 2)
            "Laatste update"   = $LaatsteUpdateTekst
        }

        Write-Host "✅ Server $Server succesvol verwerkt." -ForegroundColor Green
    }
    catch {
        Write-Host ("⚠️ Fout bij ophalen van gegevens voor {0}: {1}" -f $Server, $_.ToString()) -ForegroundColor Yellow
        $FailedServers += $Server
    }
}

# Sorteer de resultaten op servernaam vóór het genereren van het rapport
$Results = $Results | Sort-Object Servernaam

# HTML-styling met datum en tijd
$DatumTijd = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
$HTMLHeader = @"
<!DOCTYPE html>
<html lang="nl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Server Info</title>
    <style>
        body { font-family: Arial, sans-serif; background-color: #f4f4f4; margin: 10px; }
        .logo { width: 100px; display: block; margin-bottom: 10px; }
        .timestamp { text-align: center; font-size: 14px; margin-bottom: 10px; color: #555; }
        table { width: 100%; border-collapse: collapse; background: white; font-size: 12px; }
        th, td { padding: 6px; text-align: left; border-bottom: 1px solid #ddd; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        th { background: #0078D7; color: white; }
        tr:hover { background: #f1f1f1; }
    </style>
</head>
<body>
    <div class="logo">
        <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAMgAAAAsCAIAAACv72BdAAAYo0lEQVR42u1dB3wUVf5/M7O9l/RCICEJJFQhFCH0UG2IH1E44S9w9z9O9PSAg6BIOSli4VA8BfU8OQtwCigqNUAo0kMNBAiE9LrZ3man3Hszs5slIERIO+T3WTa7b2fevPJ9v/4eGMuy4AE9oMYm7H8FWAwDMIxrsfCvYXexwOMDHhp4SdJD0S4fcJLAQ8EX63LbSJqxe4GPAfAdksUDGPiHxW3OWooiMawhj8E0aiOBY4AFWinAcSAXAYUYyOGLAAqFWiISK8QsLIEvmQhIJHL4Dl9ivME9ZwE/Q3Ci8Ibf1dLUwsBiuHec/8B9gbMpzOdtp9XhZW0ka3V5a51UjRsz2SxVDqbKhVdZLGYnhAxm9QK70+720RywMC/pAbQPMFylNFc75YO9R5MWeFDdBxw0CFUATTsEL2AD34R3/m5MDAh/D2EJgUNgSTlgyUWsWqHTyjC1FGhlRJjOEKbEw5SMUaMJUYqNCqBVyDUyXCPF8Ds1BD2e8Y8h/yzhrSWpJYHFM6FfmkEvDcwu2uTwlVjdZVa6qNZSaMGLTaZqB1XjIG1ut9OHM27IZCjAYoCmAc5Bkq8RA8Jn4Mcpz+jqPwtrOPP7dcTe8KfuIxsEQQiHACYZ/jMLMBGCO4SjRKGUsmqZ1KCShyjYaIOxrZ6I1cljDIoYjShULTcoMcUv8T0WDUkT9ayB1ITAYrgVJHDvAGu6FZndTIWdKqxxXKom8yst+Say2GyrsJJWl4ciaeCj0O0Ytypx3I8b/7vQj5YdxsajOuSxAtT4d4ZnjdyygaJXJlFLReEaebRWlhCmTgiRJYcp40NUUVoCMrxbMDnubpybgttORaNREwILLT9QX6LBh1XaqKsmz4Vy5+kS88VK93WTrdLmdbm8HIA4Job7X8Fs5r6Bzr1QPdhBsPCYo1k0XCKRSAbRpmhjUCaHq7rG6FPDZe3D5DFaqZioVw/3akpwNRKwGGEd3Cza3D5wrcZ9ttR+rMRxurD2SrWt3OJhoOJDUWgsCB5AWJ1u9QBAd0f8PAZ4G/+CyqJMHKKWx4dpukSr09rou8WqEo1yvQKvfyvrR1ojoa0RgIU4UxAYYHeum70nipwH8s2Hr1VfqbLbrG5A+pDeQOB+VsQbeA8w1JQkQI37QDNoYigWcjWpWpoQqnkozjiovT6tjSopTAGNiUanuwGWoDxhdWIOluRVu/fnW7dfqMopMhebHACKNhx2A0daEX6fwgjOFsUgHRwSFDYEgQYCKoXNaQ/BUUWPbiCf8QtQhoMaFKBScYhO2SU2ZGiibkgHY5dIpUKC1V2LCZ2DbOzXMrJ74lgVdnrvFct3ZyoO5FeW1TiBlwQiTGBLt7H37g/y+iRqVUZyeBuDSiwmtpwqKaqo1mmUE3onMGxzmdoYkIvw704XFlTZGoytIApYBhBkPhqIxVqtvE98+COdwkek6BJDZHUX/grXob9pdxyDm/lTfg357emqb0+X5hTW0g4XpzZynAm7H9nSLcnrS42P2jmjV5RGmM6uy46czSseN6DDN1M7N3NbBr13IvtcMZDcszzj9TOKRhJTJk2KMYxJDXumR1SvNgr/BZxp2TAA/wqOVWzxbThZ+e/jJWcLTcDjQWAS4fc/Z7qZaIaQyPIXDW2rF/H6pdUDwubuJM22v09N//OA8GZuTsS8rEqoe4gaz8bjORnPxkTiuEjNU91jJvWO6RIla3gdtwYWsvo5/zNP2y9aV+27vv18CXB6gARHXPc3iKcAucnJQzv9a2IyhyokJbbn2Ue9swdOxg8zB6e3U3ooFmdBMwhDqFwVWenuS/cwFNlU4R5eUJIUEMu6JxheHJT4bI9QXtlnWUGZvOWTb8exIN/71/HqN7dfvlxYhfighPhN44knOFxeJitzyJD2GvjN4mbGfHjy56uVyPWPAalEwgrrvTkIQ0564CNJgBFN7mj3S0lDmHbGoMQ/D4w1KNAjIQ+6A7D4aF0AOV+fMs/dnFtUVIHwJCJ+63gSCC1fhVxZsSxDzYmFP27IX/PDKaCSCKMbHDdsJgpyATbHAHAMzEspDLq5I5Mzh8UiCcyiGGw9dN0ALJzT/vOqvc9/fuZIbgkQYy0JKdgBvm28Lx4a9jRnRxC3ahVSPBmkdfK3EH79r15VgHPJ8uUM59qpV8j6TXH+J1hOEIK7BKEKhTA7J8WczUzjKw7LzK6uNAMpJgCLFw+U//aAWXPDxPjBh3NzQnFRQjH3lOAuwxfkEAynxPGVsIy/Zq6EwFssHMiPtoeKjglbM6HzmBT9zZcgYPG8iu/+p0drpq07CjxeIBW1JJeiGb1GHaEmoOlu99JlFfbU+PB+CWoRxh4vch+/XMaNdZAbjaQNRn3/9iFtdayTwo8U2HOvVwkuWZoJ06uMChGsyseAa9AyR7ewaoU0Ri9HrgGAXa6wCT4er0+lV/dLCE0MEZE0OF7iOZVfhn4SIR+VXCbb+If09iGSDqEiPoh8rNhTZKEnfJxN0xSaZg9FqOT9E8JTI8TwIbkVvuzL5QCKKklAKwEJYWoRgUG6VGGDNtbApCijSrLtfInb5Q7VqY0qdGWNk6ypdQ9Mje4WLal2gK3nyu21VqBQjkiJ7BAmsnjAtgvVVTUWIBa3ZKiZ5YQjDV4a03XVk/HgRrHIAYsRvAmLtxcu2HgKwDEREffyxEZosdv38f8PmNYnFH7bfsldUuuY1jc08PuibQULvzkNpBzfQqsc+9u4HjOHRMuDLO6Np03PfnKYgT0n6VeffOiNR9oClHHFxGTudLq9wMf+OGvI6I5IT/rxovORd7MAAZkB9vKoLgtHtdXK6qZr92XHuLVHbHYHxESH2NCL8/sFNxM+P6fY0WPxbsTbfMzkwSlLH0sM+CAg5VZ6R//jeFFZNcIWXC1aVfnSoVJudKesL5k5KCQ1AsnUqMw95WW1W2ZlPN5ZC79+c84jJTyPpuj4Si5VeyavOw95Q9dIKV9i97A9Vxy6XFIDxC06U4BDk5se3bvdj9O7BxcjYPED9PHhqj+sOYjy01o8nYxTkA++OqxfvOrmX2BTvTQ073fb7E60Xhls0/SHx3YLubmaf5+omfThQYAxo3rE/zS9G3973IL9xSXVCXHh+Qv685f1efvY0bwSKJk+ndZvSp/Qm+vJuuwY9s5eQPnGPpy4aWpnIQXLbw2t+bnyj2sPQta48rm+Lw+K4G+5Wu2NNUgkBAJobrm3E0QeoCCwurSLOpPZGxbSLOvxYUoJurjKToXP2QHV/wuLR3QMl97YYc45eZPYhyUfHaqc/skhlFjY4uovbJDTO2l4588ndgiUCTpWpZ2KeXUX5fW2MK/iCbFQUfHSjBgtGngWsaji0yW1Sx5N5tc3pLj5e4uqrIBkFjydtnBULODyt6Z+lXfoSsWSxztP6GHkL4udn11SZmobbSxYPJAv6bHiUM6Zsg+mp/+pPwLBjjz7yLd2Q6k2ZViXTyck8te88J/8rWcLZw3r+NLAKOGuN4/mnC9cOKHvAvQsOGYovZRiWBGOTfny0mc/npo0qtvnz6FhrXLQQ947lnu5vENi1MnZfRQcdAa9dyr73HV435N92n87rUsgunrourPc7ChzEH/+/KBCIy9bkqGVI7Ta3OwrWy7F6uQLR8Xxbm+oPc7enE9SzMpxSRJuirZdsIx+ew9KGmxxYAHOavEy2/46ZGQHDV8gAGvZrtJ5XxwGcAW1hlZSdLheW7J0MO/zm/N94YqNRyHapj7e55PxbQFKHwWRr2Y5rDatTlW6JINf99P/U/DRd8ehDRsZaSxbMoivafjqE7tOFYmUsrKlI0JVaEIeW3v+p3NF1ndG83elvX3sxIViXCkvemNEtBZdsGRn6WvrDgMpLlWrTCuG85dN/erKP3edkWrVqZGaE3P68mP0yEdn918ps7t8ACdKlgq3Z/5Q/OG+XK1M4vYxO17q2z1aCQtnbsp/94ezEFiLxvd6fWQsj5UpX+d/tuMcYGkggcoHkxQbcun1dL7Zv1uX++W2c3iIzvzWCA0nl1dklcyB/EkkOboko1cbxMjXHTdN/nA/kLUag93jG9qtze4Xe/LfBGA9sebEd8eLkcLeGoikeibHHJ8lWF5xr+8vqqgFPnrh0z0XwBUMQF4l2XHBDuDzPd4nccvvUQjF7WMjMjnhCECIXl22ZCive2SsPrH7dDHs5tH5Gb3i0BxP/TKXxcX/fDYJft55yT5ixW5oKKclRR2b3Zt/XOKig/ml1SiZDicq3xweqkLImvL11c92n4Ef0ju32f8yGjtoIEbMzaoxIVOgZ2LE8b/2vU1Ebe73+W9uOYOk9sxBYzsbYMnBAnv6wt1A4TddSXpUj3a8vIaWX2Tm3hqTOSpcX7xkMK+XDFyVs/9sAZDJ8hdnJBhRkzK/L1i+OQfwLLE1EEVHhoaUvtFfyH7igfXs5+fX778MZOKWbh1HHt/EAUlfTO4EkAcSImaX1+MGHnrTzMFjOyPL9qdc85i398G2z360w4qxKbCkwEzHz9sJ5wfir09K3OG/9OBrSlh44Fq5CYrJT6f3n9IbBVtWHzL3ayvrHi2Hn/uvPH4oF8IOPN2vw4bnUT1WNx2ZudPtISGDbBcVcm2hwEIefufY4bxSqCS98lj3d59AFlCxmWrz2g6UGE3Szwzs9PWkJAgsL82eKvXinPKFjFDkWQAqqWjWlkt7ci5B+X5uUUanSPToN7NK5/4bigi/RuUmZz7a9e2x7eHH67VUO1gz6e2X2ubgKwIDQDK93GQM0RW/MYS3UcZ8dOanE9daCy+A5KO5ERM0VwFY607UTn5/L1CKWwVfdXpXPNd39tBo+PF8ubvzgl2AYAGNBWZl+e6SzC+OQOX5pdEdVz3VEXD4C/nrDhrZbvh//jL0qa7Itsop9vR4Yxfyznl9fxrd7YOnEgCCDouMPgwcvGpPX74HpWP4qCf6JG2ehnAMByMic29VeQ3kLiunpb88EGGxyOKLe20X2ovhob94IX1iT6Tg77liHbpsD5JEJPV0euqGycmwsMTqi529C/kXIMrEEqNBBpVvEY5X1TqgDSGTK8qXZeg40fa7dee/zPavZM4KXvenAc+loZr3XbUNXgKVfXba8G4fP4PaXO2go+btolyu7u3Dc+YKZmnyYmgVVre8VRg0a/+X0fmzicn8NwFYDBIBh64VVwJpSzMtNMr0tjmDR3ZAxvamM6ZxK7Ph/Mll0rLlI/hZmbjuwlfZedDCT0uKPjarF3/fhlM1aw9en5gWO6WPEAYe9N7J7LOFqEck1S8lFq3+IGk1cNXx/eeK0a8UHReuv+7X7nfkWd7efXVkStTMIZF8ybP/yl2/Pw+BgAJnFmZ0iULgfmdP+SzIciD3oJhgN8Sc7/Lf21sQqpWtGd9lVIqWa5jpmQ/2AxxrF6m/tkh4StqKYyfyS4HY79/ygWOvD0uLRXkEHx0qn772Z1i6emqfFwbEwJITJa60Rbsgdxzfv8P651Nhic3LRmXucrrcd5Mt0xRE0XAhXV6UkWgU8MMBCzndwalST49le1hoGN57Asa9EMrjxq4uyYg3IgNw4fbSRRsOwyYlRodefl2YvJ5vHTl5pRwZsBS75cXBj3fW3FzNS5sL3996UjCaKDrCqClfMiSAqoMFjnTIzHhPGBwBkvnHlH7TH76Fr2H57qLMr3KADIcNUyh4w40Hd+5X2ZcQ2tDt7PoZA8Z3ExzQbgoEPGp5Vb7eb+232ezw9hHd2mx/AcloCO/wzKzqWn9KAsMQYnHVsuEGJbrtjxvz12yH+hy+79VhAxPUsOTrkzUTVmfDYVkysfe8DMTIL1WTHV7bCXAG3HF3WDMQilKDT/7Qf2ofY6BMABbLhRCyr9pGrj7qsdnQOm65SI5WpcxbOMygQIkUT35ybvORfAiIIV3isl58yMtAjgs6Lt5XZbagYAtFi6TSf07q/bsehkB7Txa7Z26+mH26sM5iYhhMJL64YEQ7AwFVYwinoe+fyD5bVMee4dBgonfGd5vRPzqwrC5WkfO+v7rl8CUg4SI2Pjo5NiRvvqBDdF3689nrVYIkolH9q5/tNrVPlNQvmiwedt2xqtmbTpNOF1qrHt+MMQ+9/2RbkkFaVIcFO1iGFmBBM9FGbd6CQbAyCQEGrDx+EGp+Utn5BSOSQgkxBmZuzn936xnYl/UzBo3vboC4/Oas7elVWcLCaEGCTaEoqH6smtTvpfSQYIFQP7vhqoka93HOmbwitBbxlohGofZgBo0GOYtwzOxwsVC5AUAslmoVMprbl2J22ITgK+Dzg9mIcG37UD1BYMUmy7VyG9KHJEEOHi4CqFFrCP/6Nlutdan3PDEM8DIhoerkCKNYLK6wWPJKLQDyb5l/jXl9T/RO2vx7pIo5SDZ63m6bw1UniVD2EmMIUSeF6eUyuc1pzauwOy1OBEqC4OtXyJVyiRgq9RDcdntQA1BPcaNaQ8MFgOG1kL1xe2r1Gh1gGRzDrG4XRfrgZVq1iiAQ8N1e0u12tbAcROFCnyZEv3Fa2ohkTb00wBuAxTJCDtayXaULtp732e1oTTc/vPgYJ084HhQeFrYC1ccEGxQ2xrG6LPv6dQalHuC32usciEADbn8LEVwPC5zUGxN6vjq8DfxyoYpMfX0nt9UR+8XbcewW2A3Ewm/uQr0uB5dg/pYwwYH5lkMV7CZJAZFkcnrC359M0ikIITcrKDW+PscKxBEr7PSibVfWHrjGwHWJ4NVsmVgsSlxkWs2JErDTIhF699I/zBw8hovfbTpbO+7v+1qL17vZiIePF0KKGN41ZvkTKd2jFYHieji/daJfQFYWm31vZV3//EiRzWThcpEJDpNNN5oQVWynhOi2RiVFtzy2uEwnIutCgdfjBRhx5W8Z7UOQ52nm5qvvbj0N5K3GOdnUxOcCkTSmko99qO2coW16xSHv/232vd4ugzQAL4cXfJVTtXZ//skCE/CQXHYy0SRbJ2BjPNSB+SP7t1Pce2WNRdrZWTaLzahXly8fJsbByr1lf9lwCmBUy0frm5oCG3hwIi5a/3zf+Kl9o2K4yNUdN1bcYTNFvXM7ciu863OqNp4sQQkbJNn4+ynQ88TFS0fE6FqL3+9ilTdl/k5oDTycGnvolbQZ31z54KfzQuL/fUnBOykIIiJU+1i3mOd6RfZvp6r7nddmwe22TTd0l04gZ4sniLDvz1VuOVOZU1hDOTx1e1PvkY3RTLhefX3xEJmYoZm7r6axiMDxb0/XPPU+cm8+3S+ZwImv95xDErA1eI8akQInQdAocQ8awsnR+lEp4U91i+zdVnV323/udYt9pYM+kG/der7iQH5NQaUNuEkhb5MI5GL/mjlA+Sh4qFrM+JdFyxJcKU4v44KdwlHOJ4s2q7S066ixiPUfKEJxW6Il4lCjsm+7kDFdooYmavk4973Q3QOrHg+DdK2WPFxgycozH71uulJp9Tm9iKOKbjz5oyEdbj0mIfBn3ANQZ+f/j1IASTyYoBhTSOKMqrS2hkFJYenx6pRwRXAy3i9tv2nosDXRbnDIyS6Wuw8XWQ5crrpQ6So22RmXF3lp+Vx1nPN4PDgapIko+DgQht9mwqAtYnJJuFbeMULXO97QL17XJVIZZ2iq0HAjA0tQ9kF9AVhhp69Uu8+U2k4UWnPLHddrbTU2DzoelKG5k8QeHGZ0txSMobozjDhHsVikUctidcqUSFX3WH33WHXHMGWM4aajshqmjP9aapmjIqHBUW4jr5k8Fyrc50otl6rdxSZbuc3rhFwNxS64HVHBjO2+PLyv4XTLY/4Y7lAy3jUplkileKhG3kavSgyTp0YbUsOliaGqKJ1E2ULZKs0HLGFDrP9k0JvJ5mErHL5Ss/NqDZ1fbb5aSxVXmyudTI3d4fCwCHAULdSE+6McAcDdgDzQlC7cJqB6uAFBAGL8GiffX4kEmqR6lSpCTcQY9G11WEK4PilEFmNQRKpxg1J0i277a2pmp9v/wHHcDi9d6wYmp7fESpZbqRKzpdSKl9aaal1UrQtYnDanDwpVFu2F5FkdzfgPt+XPPw6cYew/juKGY5KD4dgYFDjWlr2xJHBs3g2n2fJzznmbYdPEMokUV4hYrVKllxNGJRauM8RqsRidOlIni9FioWqlXg40MhHR6hdO6wJWXdQV+A8fvdMIkjTj8GJ2H3C4nLVu1uzBrQ67yUGaPJjJxZjttTYP7SABd8I77XLbvBTm4w658DIYzTA06fJPMKib9YZvlK8LuvqbirgpixMykUgsIVhoq0hFQEywColCIZOpJKxKAlRioNMajXKRUcEaFbherdHKcL2M1SnkKimhkrLyO/qOeNz60dsKQwCtC1jNQDSL/scAaJ560JZ9liS9DMM6SQQLyPl4sQNRSNMN4mFwQtE2Hu50Dn6Lk4wAEhErEktFOCEh0H8UIMZbxZ66ZqbfHLAeUPPQfwFPas9OElOb/AAAAABJRU5ErkJggg=="alt="Novoferm Logo">
    </div>
    <div class='timestamp'>Serverrapport gegenereerd op $DatumTijd</div>
    <table>
        <tr>
            <th>Servernaam</th>
            <th>OS</th>
            <th>Buildnummer</th>
            <th>Fabrikant</th>
            <th>Model</th>
            <th>Processor</th>
            <th>Geheugen (GB)</th>
            <th>IP Adres</th>
            <th>Laatste reboot</th>
            <th>Uptime (Dagen)</th>
            <th>Laatste Update</th>
        </tr>
"@

# Data in HTML-tabel zetten
$HTMLBody = ""
foreach ($Result in $Results) {
    $HTMLBody += "<tr>"
    foreach ($Property in $Result.PSObject.Properties) {
        $HTMLBody += "<td>$($Property.Value)</td>"
    }
    $HTMLBody += "</tr>"
}

# HTML afsluiten en niet-bereikbare servers toevoegen
$HTMLFooter = "</table>"

if ($FailedServers.Count -gt 0) {
    $HTMLFooter += "<h3>Niet-bereikbare servers:</h3><ul>"
    foreach ($Server in $FailedServers) {
        $HTMLFooter += "<li>$Server</li>"
    }
    $HTMLFooter += "</ul>"
}

$HTMLFooter += "</body></html>"

# HTML-bestand opslaan
$HTMLFile = "C:\ServerInfo.html"
$HTMLContent = $HTMLHeader + $HTMLBody + $HTMLFooter
$HTMLContent | Out-File -FilePath $HTMLFile -Encoding utf8

# E-mail instellingen
$strFrom = "noreply@novoferm.nl"
$strTo = "IT.Monitoring@novoferm.nl"
$strSubject = "Server Monitoring Rapport - $DatumTijd"
$strSMTPServer = "exchange.novoferm.nl"

# E-mailbericht met inline HTML
$objEmailMessage = New-Object system.net.mail.mailmessage
$objEmailMessage.From = $strFrom
$objEmailMessage.To.Add($strTo)
$objEmailMessage.Subject = $strSubject
$objEmailMessage.IsBodyHTML = $true
$objEmailMessage.Body = $HTMLContent  

$objSMTP = New-Object Net.Mail.SmtpClient($strSMTPServer)
try { $objSMTP.Send($objEmailMessage) }
catch { Write-Host "Fout bij verzenden: $_" -ForegroundColor Red }

Start-Process $HTMLFile
