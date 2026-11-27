#requires -Version 5.1
<#
  FaturaAktarim.ps1
  "Gelen" klasorune atilan ZIP dosyalarini acar, icindeki UBL-TR e-Fatura / e-Arsiv XML'lerini okur
  ve SQL Server'daki tabloya yazar (her satir = 1 fatura kalemi, baslik alanlari her kalemde tekrar eder).

  - Tabloda ayni UUID'li fatura varsa tekrar eklenmez (ayni ZIP iki kez atilsa da cift kayit olmaz).
  - Her ZIP tek islemdir (transaction): ya tamami yazilir ya hic yazilmaz.
  - Basarili ZIP -> Islenen, hatali ZIP -> Hatali klasorune tasinir. Ayrinti -> Log klasoru.

  Kullanim:
    powershell -ExecutionPolicy Bypass -File FaturaAktarim.ps1            (normal calisma)
    powershell -ExecutionPolicy Bypass -File FaturaAktarim.ps1 -Deneme    (sadece oku/say, SQL'e yazma, ZIP'i tasima)
#>
param(
    [string]$AyarDosyasi = (Join-Path $PSScriptRoot 'ayarlar.json'),
    [switch]$Deneme
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Data
$INV = [Globalization.CultureInfo]::InvariantCulture
$NS_CBC = 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2'
$NS_CAC = 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2'
$KDV = '0015'; $OTV = '9077'                                   # GIB vergi kodlari
$SASI = [string][char]0x015E + 'AS' + [string][char]0x0130   # "SASI" (Turkce S-cedilla ve noktali I ile)

# Tablo kolonlari: ad ve tur (s=metin, d=tarih, t=saat, n=ondalik, i=tamsayi)
$KOLONLAR = @(
  'Klasor:s','DosyaAdi:s','FaturaNo:s','UUID:s','Profil:s','FaturaTipi:s','Tarih:d','Saat:t','ParaBirimi:s',
  'SatirSayisi:i','SiparisNo:s','SiparisTarihi:d','IcBelgeNo:s','SaseNo:s','MotorNo:s','PlakaNo:s','AracKm:s',
  'IadeFaturaNo:s','IadeFaturaTarihi:d','IstisnaKodu:s','IstisnaAciklama:s','Notlar:s','EkBilgiler:s',
  'SaticiVKN:s','SaticiUnvan:s','SaticiAdres:s','SaticiIlce:s','SaticiIl:s','SaticiVergiDairesi:s','SaticiTel:s','SaticiEposta:s',
  'AliciKimlikTipi:s','AliciVKN:s','AliciUnvan:s','AliciAdres:s','AliciIlce:s','AliciIl:s','AliciVergiDairesi:s','AliciTel:s','AliciEposta:s',
  'FatMalHizmetToplam:n','FatIskontoToplam:n','FatVergiHaricTutar:n','FatVergiDahilTutar:n','FatOdenecekTutar:n',
  'FatKDV:n','FatOTV:n','FatTevkifat:n','FatTevkifatOran:n','FatTevkifatKodu:s',
  'SatirNo:i','UrunKodu:s','UrunAdi:s','Miktar:n','Birim:s','BirimFiyat:n','SatirTutar:n','SatirIskontoOran:n','SatirIskonto:n',
  'SatirKDVOran:n','SatirKDV:n','SatirOTVOran:n','SatirOTV:n'
) | ForEach-Object { $p = $_.Split(':'); [pscustomobject]@{ Ad = $p[0]; Tur = $p[1] } }

# ---------------------------------------------------------------- ayarlar ve klasorler
if (-not (Test-Path -LiteralPath $AyarDosyasi)) {
    throw "Ayar dosyasi bulunamadi: $AyarDosyasi - ayarlar.ornek.json dosyasini ayarlar.json adiyla kopyalayip sunucu/veritabani bilgilerini doldurun."
}
$ayar = Get-Content -LiteralPath $AyarDosyasi -Raw -Encoding UTF8 | ConvertFrom-Json
$ana = if ($ayar.AnaKlasor) { $ayar.AnaKlasor } else { $PSScriptRoot }
$dir = @{}
foreach ($k in 'Gelen', 'Islenen', 'Hatali', 'Log', 'Gecici') {
    $dir[$k] = Join-Path $ana $k
    New-Item -ItemType Directory -Force -Path $dir[$k] | Out-Null
}
$tablo = [string]$ayar.Tablo
if ($tablo -notmatch '^[\w\.\[\] ]+$') { throw "Tablo adi gecersiz: $tablo" }

function Write-Log([string]$Mesaj, [string]$Seviye = 'BILGI') {
    $satir = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Seviye, $Mesaj
    $dosya = Join-Path $dir.Log ('aktarim_{0}.log' -f (Get-Date -Format 'yyyy-MM'))
    Add-Content -LiteralPath $dosya -Value $satir -Encoding UTF8
    Write-Host $satir
}

# ---------------------------------------------------------------- XML okuma yardimcilari
function Get-Txt($Node, [string]$XPath) {
    if ($null -eq $Node) { return $null }
    $x = $Node.SelectSingleNode($XPath, $script:ns)
    if ($null -eq $x) { return $null }
    $v = $x.InnerText.Trim()
    if ($v.Length -eq 0) { return $null }
    return $v
}

function Get-Num([string]$Metin) {
    if ([string]::IsNullOrWhiteSpace($Metin)) { return $null }
    $d = [decimal]0
    if ([decimal]::TryParse($Metin.Trim(), [Globalization.NumberStyles]::Float, $INV, [ref]$d)) { return $d }
    return $null
}

function Get-Party($Party) {
    $r = @{ tip = $null; vkn = $null; unvan = $null; adres = $null; ilce = $null; il = $null; vd = $null; tel = $null; mail = $null }
    if ($null -eq $Party) { return $r }
    $ids = @{}
    foreach ($i in $Party.SelectNodes('cac:PartyIdentification/cbc:ID', $script:ns)) {
        $sch = $i.GetAttribute('schemeID')
        if (-not $ids.ContainsKey($sch)) { $ids[$sch] = $i.InnerText.Trim() }
    }
    if ($ids.ContainsKey('VKN')) { $r.tip = 'VKN' } elseif ($ids.ContainsKey('TCKN')) { $r.tip = 'TCKN' }
    if ($r.tip) { $r.vkn = $ids[$r.tip] }
    $ad = Get-Txt $Party 'cac:PartyName/cbc:Name'
    if (-not $ad) {
        $parca = @((Get-Txt $Party 'cac:Person/cbc:FirstName'), (Get-Txt $Party 'cac:Person/cbc:FamilyName')) | Where-Object { $_ }
        if ($parca) { $ad = $parca -join ' ' }
    }
    $r.unvan = $ad
    $r.adres = Get-Txt $Party 'cac:PostalAddress/cbc:StreetName'
    $r.ilce  = Get-Txt $Party 'cac:PostalAddress/cbc:CitySubdivisionName'
    $r.il    = Get-Txt $Party 'cac:PostalAddress/cbc:CityName'
    $r.vd    = Get-Txt $Party 'cac:PartyTaxScheme/cac:TaxScheme/cbc:Name'
    $r.tel   = Get-Txt $Party 'cac:Contact/cbc:Telephone'
    $r.mail  = Get-Txt $Party 'cac:Contact/cbc:ElectronicMail'
    return $r
}

# vergi kodu -> @{Amt=toplam tutar; Pct=en yuksek oran}, ayrica istisna kod/aciklama listesi
function Get-Taxes($TaxTotal) {
    $v = @{}; $ist = New-Object System.Collections.ArrayList
    if ($null -ne $TaxTotal) {
        foreach ($s in $TaxTotal.SelectNodes('cac:TaxSubtotal', $script:ns)) {
            $code = Get-Txt $s 'cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode'
            if (-not $code) { $code = Get-Txt $s 'cac:TaxCategory/cac:TaxScheme/cbc:Name' }
            if (-not $code) { $code = '?' }
            if (-not $v.ContainsKey($code)) { $v[$code] = @{ Amt = [decimal]0; Pct = $null } }
            $a = Get-Num (Get-Txt $s 'cbc:TaxAmount')
            if ($null -ne $a) { $v[$code].Amt = $v[$code].Amt + $a }
            $p = Get-Num (Get-Txt $s 'cbc:Percent')
            if ($null -ne $p -and ($null -eq $v[$code].Pct -or $p -gt $v[$code].Pct)) { $v[$code].Pct = $p }
            $kc = Get-Txt $s 'cac:TaxCategory/cbc:TaxExemptionReasonCode'
            $kr = Get-Txt $s 'cac:TaxCategory/cbc:TaxExemptionReason'
            if ($kc) {
                $key = "$kc`t$kr"
                if (-not ($ist | Where-Object { $_.Key -eq $key })) { [void]$ist.Add(@{ Key = $key; Kod = $kc; Aciklama = $kr }) }
            }
        }
    }
    return @{ V = $v; Ist = $ist }
}
function Get-TaxVal($Taxes, [string]$Code, [string]$Alan) {
    if ($Taxes.V.ContainsKey($Code)) { return $Taxes.V[$Code][$Alan] }
    return $null
}

# Bir XML dosyasini okur; fatura degilse $null, fatura ise kalem satirlari (hashtable listesi) dondurur
function Read-Invoice([string]$Yol, [string]$Klasor) {
    $doc = New-Object System.Xml.XmlDocument
    $doc.XmlResolver = $null
    $doc.Load($Yol)
    $r = $doc.DocumentElement
    if ($r.LocalName -ne 'Invoice') { return $null }
    $script:ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $script:ns.AddNamespace('cbc', $NS_CBC); $script:ns.AddNamespace('cac', $NS_CAC)

    $supNode = $r.SelectSingleNode('cac:AccountingSupplierParty/cac:Party', $script:ns)
    $cusNode = $r.SelectSingleNode('cac:AccountingCustomerParty/cac:Party', $script:ns)
    $sup = Get-Party $supNode
    $cus = Get-Party $cusNode

    $ref = @{ SASI = $null; MOTOR = $null; PLAKA = $null; KM = $null }; $icno = $null; $ek = @()
    foreach ($d in $r.SelectNodes('cac:AdditionalDocumentReference', $script:ns)) {
        if ((Get-Txt $d 'cbc:DocumentType') -ceq 'XSLT') { continue }
        $code = Get-Txt $d 'cbc:DocumentTypeCode'; $id = Get-Txt $d 'cbc:ID'; $desc = Get-Txt $d 'cbc:DocumentDescription'
        if (-not $desc) { $desc = '' }
        $descU = $desc.ToUpperInvariant()
        $key = $null
        if ($descU.Contains($SASI)) { $key = 'SASI' } elseif ($descU.Contains('MOTOR')) { $key = 'MOTOR' } elseif ($descU.Contains('PLAKA')) { $key = 'PLAKA' } elseif ($descU.Contains('KM')) { $key = 'KM' }
        if ($key) { if (-not $ref[$key]) { $ref[$key] = $id } }
        elseif ($code -ceq 'CUST_INV_ID') { if (-not $icno) { $icno = $id } }
        elseif (@('OUTPUT_TYPE', 'TRANSPORT_TYPE', 'EREPSENDT', 'FATURA_TURU') -cnotcontains $code) {
            $ek += ('{0}{1}={2}' -f $code, $(if ($desc) { '/' + $desc } else { '' }), $id)
        }
    }

    $notlar = @()
    foreach ($n in $r.SelectNodes('cbc:Note', $script:ns)) {
        $tx = $n.InnerText
        if ($tx -and $tx.Trim()) { $notlar += $tx.Trim().Replace("`r", '').Replace("`n", ' / ') }
    }

    $ttNode = $r.SelectSingleNode('cac:TaxTotal', $script:ns); $tt = Get-Taxes $ttNode
    $lm = $r.SelectSingleNode('cac:LegalMonetaryTotal', $script:ns)

    # Tevkif edilen KDV = alt kirilim tutarlari (bazi sistemler ust TaxAmount'a saticida kalan KDV'yi yaziyor)
    $wht = $null; $whp = $null; $whk = $null
    $wh = $r.SelectSingleNode('cac:WithholdingTaxTotal', $script:ns)
    if ($null -ne $wh) {
        $whs = $wh.SelectNodes('cac:TaxSubtotal', $script:ns)
        if ($whs.Count -gt 0) {
            $wht = [decimal]0
            $kodlar = @()
            foreach ($s in $whs) {
                $a = Get-Num (Get-Txt $s 'cbc:TaxAmount'); if ($null -ne $a) { $wht = $wht + $a }
                $c = Get-Txt $s 'cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode'
                if ($c -and ($kodlar -cnotcontains $c)) { $kodlar += $c }
            }
            $whp = Get-Num (Get-Txt $whs[0] 'cbc:Percent')
            if ($kodlar.Count) { $whk = $kodlar -join ', ' }
        }
    }
    $br = $r.SelectSingleNode('cac:BillingReference/cac:InvoiceDocumentReference', $script:ns)
    $lines = $r.SelectNodes('cac:InvoiceLine', $script:ns)
    $saat = Get-Txt $r 'cbc:IssueTime'; if ($saat -and $saat.Length -gt 8) { $saat = $saat.Substring(0, 8) }
    $sipNo = Get-Txt $r 'cac:OrderReference/cbc:ID'

    $head = @{
        Klasor = $Klasor; DosyaAdi = [IO.Path]::GetFileName($Yol); FaturaNo = (Get-Txt $r 'cbc:ID'); UUID = (Get-Txt $r 'cbc:UUID')
        Profil = (Get-Txt $r 'cbc:ProfileID'); FaturaTipi = (Get-Txt $r 'cbc:InvoiceTypeCode'); Tarih = (Get-Txt $r 'cbc:IssueDate'); Saat = $saat
        ParaBirimi = (Get-Txt $r 'cbc:DocumentCurrencyCode'); SatirSayisi = [int]$lines.Count; SiparisNo = $sipNo
        SiparisTarihi = $(if ($sipNo) { Get-Txt $r 'cac:OrderReference/cbc:IssueDate' } else { $null })
        IcBelgeNo = $icno; SaseNo = $ref.SASI; MotorNo = $ref.MOTOR; PlakaNo = $ref.PLAKA; AracKm = $ref.KM
        IadeFaturaNo = (Get-Txt $br 'cbc:ID'); IadeFaturaTarihi = (Get-Txt $br 'cbc:IssueDate')
        IstisnaKodu = $(if ($tt.Ist.Count) { ($tt.Ist | ForEach-Object { $_.Kod }) -join ', ' } else { $null })
        IstisnaAciklama = $(if ($tt.Ist.Count) { $x = @($tt.Ist | Where-Object { $_.Aciklama } | ForEach-Object { $_.Aciklama }); if ($x.Count) { $x -join ' | ' } else { $null } } else { $null })
        Notlar = $(if ($notlar.Count) { $notlar -join ' | ' } else { $null })
        EkBilgiler = $(if ($ek.Count) { $ek -join '; ' } else { $null })
        FatMalHizmetToplam = (Get-Num (Get-Txt $lm 'cbc:LineExtensionAmount')); FatIskontoToplam = (Get-Num (Get-Txt $lm 'cbc:AllowanceTotalAmount'))
        FatVergiHaricTutar = (Get-Num (Get-Txt $lm 'cbc:TaxExclusiveAmount')); FatVergiDahilTutar = (Get-Num (Get-Txt $lm 'cbc:TaxInclusiveAmount'))
        FatOdenecekTutar = (Get-Num (Get-Txt $lm 'cbc:PayableAmount'))
        FatKDV = (Get-TaxVal $tt $KDV 'Amt'); FatOTV = (Get-TaxVal $tt $OTV 'Amt')
        FatTevkifat = $wht; FatTevkifatOran = $whp; FatTevkifatKodu = $whk
        AliciKimlikTipi = $cus.tip
    }
    foreach ($pair in @(@('Satici', $sup), @('Alici', $cus))) {
        $p = $pair[1]
        $head[$pair[0] + 'VKN'] = $p.vkn; $head[$pair[0] + 'Unvan'] = $p.unvan; $head[$pair[0] + 'Adres'] = $p.adres
        $head[$pair[0] + 'Ilce'] = $p.ilce; $head[$pair[0] + 'Il'] = $p.il; $head[$pair[0] + 'VergiDairesi'] = $p.vd
        $head[$pair[0] + 'Tel'] = $p.tel; $head[$pair[0] + 'Eposta'] = $p.mail
    }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($l in $lines) {
        $ltNode = $l.SelectSingleNode('cac:TaxTotal', $script:ns); $lt = Get-Taxes $ltNode
        $q = $l.SelectSingleNode('cbc:InvoicedQuantity', $script:ns)
        $ac = $l.SelectSingleNode('cac:AllowanceCharge', $script:ns)
        $row = $head.Clone()
        $row.SatirNo = [int](Get-Txt $l 'cbc:ID')
        $row.UrunKodu = Get-Txt $l 'cac:Item/cac:SellersItemIdentification/cbc:ID'
        $row.UrunAdi = Get-Txt $l 'cac:Item/cbc:Name'
        if ($null -ne $q) {
            $row.Miktar = Get-Num $q.InnerText
            $u = $q.GetAttribute('unitCode'); $row.Birim = $(if ($u) { $u } else { $null })
        } else { $row.Miktar = $null; $row.Birim = $null }
        $row.BirimFiyat = Get-Num (Get-Txt $l 'cac:Price/cbc:PriceAmount')
        $row.SatirTutar = Get-Num (Get-Txt $l 'cbc:LineExtensionAmount')
        $row.SatirIskontoOran = Get-Num (Get-Txt $ac 'cbc:MultiplierFactorNumeric')
        $row.SatirIskonto = Get-Num (Get-Txt $ac 'cbc:Amount')
        $row.SatirKDVOran = Get-TaxVal $lt $KDV 'Pct'; $row.SatirKDV = Get-TaxVal $lt $KDV 'Amt'
        $row.SatirOTVOran = Get-TaxVal $lt $OTV 'Pct'; $row.SatirOTV = Get-TaxVal $lt $OTV 'Amt'
        $rows.Add($row)
    }
    return , $rows
}

# ---------------------------------------------------------------- SQL yardimcilari
function New-SqlConnection {
    $csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $csb['Data Source'] = $ayar.Sunucu
    $csb['Initial Catalog'] = $ayar.Veritabani
    if ($ayar.WindowsKimlik) { $csb['Integrated Security'] = $true }
    else { $csb['User ID'] = $ayar.Kullanici; $csb['Password'] = $ayar.Sifre }
    $csb['Application Name'] = 'FaturaAktarim'
    $csb['Connect Timeout'] = 30
    $c = New-Object System.Data.SqlClient.SqlConnection($csb.ConnectionString)
    $c.Open()
    return $c
}

# Hedef tablonun kolonlarini ve tiplerini okur (kendi kurdugunuz tablo farkli tiplerde olabilir)
function Get-TargetSchema($Conn) {
    $cmd = $Conn.CreateCommand()
    $cmd.CommandText = "SELECT TOP 0 * FROM $tablo"
    $rd = $cmd.ExecuteReader([System.Data.CommandBehavior]::SchemaOnly -bor [System.Data.CommandBehavior]::KeyInfo)
    $st = $rd.GetSchemaTable(); $rd.Close()
    $res = @{}
    foreach ($c in $st.Rows) {
        $res[[string]$c.ColumnName] = @{ Type = [type]$c.DataType; Size = [int]$c.ColumnSize; Auto = [bool]($c.IsAutoIncrement -eq $true) }
    }
    return $res
}

# Okunan degeri hedef kolonun tipine kultur-bagimsiz cevirir
function Convert-ForColumn($Deger, [string]$Tur, $Kolon, [string]$KolonAdi) {
    if ($null -eq $Deger) { return [DBNull]::Value }
    $t = $Kolon.Type
    if ($t -eq [string]) {
        $s = if ($Deger -is [decimal]) { $Deger.ToString($INV) } else { [string]$Deger }
        if ($Kolon.Size -gt 0 -and $Kolon.Size -lt 1000000 -and $s.Length -gt $Kolon.Size) {
            throw "Kolon '$KolonAdi' kisa kaliyor: deger $($s.Length) karakter, tabloda $($Kolon.Size). Kolonu buyutun (or. NVARCHAR(MAX))."
        }
        return $s
    }
    if ($t -eq [datetime]) {
        if ($Tur -eq 't') { return [datetime]::ParseExact('1900-01-01 ' + $Deger, 'yyyy-MM-dd HH:mm:ss', $INV) }
        return [datetime]::ParseExact($Deger.Substring(0, 10), 'yyyy-MM-dd', $INV)
    }
    if ($t -eq [timespan]) { return [timespan]::ParseExact($Deger, 'hh\:mm\:ss', $INV) }
    if ($t -eq [datetimeoffset]) { return New-Object DateTimeOffset([datetime]::ParseExact($Deger.Substring(0, 10), 'yyyy-MM-dd', $INV)) }
    return [Convert]::ChangeType($Deger, $t, $INV)
}

function Get-ExistingUuids($Conn) {
    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $cmd = $Conn.CreateCommand(); $cmd.CommandTimeout = 300
    $cmd.CommandText = "SELECT DISTINCT UUID FROM $tablo WHERE UUID IS NOT NULL"
    $rd = $cmd.ExecuteReader()
    while ($rd.Read()) { [void]$set.Add([string]$rd.GetValue(0)) }
    $rd.Close()
    return , $set
}

# ---------------------------------------------------------------- ZIP isleme
function Test-FileReady($Dosya) {
    # Kopyalanmasi bitmemis dosyayi alma: son 20 sn icinde degistiyse veya baska program aciksa bekle
    if (((Get-Date) - $Dosya.LastWriteTime).TotalSeconds -lt 20) { return $false }
    try { $s = [IO.File]::Open($Dosya.FullName, 'Open', 'Read', 'None'); $s.Close(); return $true } catch { return $false }
}

function Expand-ZipDeep([string]$Zip, [string]$Hedef) {
    [IO.Compression.ZipFile]::ExtractToDirectory($Zip, $Hedef)
    # ZIP icinden ZIP cikarsa onlari da ac (en fazla 3 seviye)
    for ($seviye = 0; $seviye -lt 3; $seviye++) {
        $ic = @(Get-ChildItem -LiteralPath $Hedef -Recurse -File -Filter *.zip)
        if (-not $ic.Count) { break }
        foreach ($z in $ic) {
            $alt = Join-Path $z.DirectoryName ([IO.Path]::GetFileNameWithoutExtension($z.Name))
            if (Test-Path -LiteralPath $alt) { $alt = $alt + '_' + [guid]::NewGuid().ToString('N').Substring(0, 6) }
            [IO.Compression.ZipFile]::ExtractToDirectory($z.FullName, $alt)
            Remove-Item -LiteralPath $z.FullName -Force
        }
    }
}

function Invoke-Zip($Zip, $Conn, $Schema) {
    $calisma = Join-Path $dir.Gecici ('{0}_{1}' -f [IO.Path]::GetFileNameWithoutExtension($Zip.Name), [guid]::NewGuid().ToString('N').Substring(0, 8))
    try {
        Expand-ZipDeep $Zip.FullName $calisma
        $xmls = @(Get-ChildItem -LiteralPath $calisma -Recurse -File | Where-Object { $_.Extension -eq '.xml' } | Sort-Object FullName)
        if (-not $xmls.Count) { throw 'ZIP icinde XML dosyasi yok.' }

        $mevcut = Get-ExistingUuids $Conn
        $yeniSatir = New-Object System.Collections.Generic.List[object]
        $faturaYeni = 0; $faturaVar = 0; $faturaDegil = 0
        foreach ($x in $xmls) {
            $klasor = if ($x.DirectoryName.TrimEnd('\') -eq $calisma.TrimEnd('\')) { [IO.Path]::GetFileNameWithoutExtension($Zip.Name) } else { $x.Directory.Name }
            try { $satirlar = Read-Invoice $x.FullName $klasor }
            catch { throw "XML okunamadi: $($x.FullName.Substring($calisma.Length + 1)) - $($_.Exception.Message)" }
            if ($null -eq $satirlar) { $faturaDegil++; continue }
            if ($satirlar.Count -eq 0) { Write-Log "Kalemsiz fatura atlandi: $($x.Name)" 'UYARI'; continue }
            $uuid = $satirlar[0].UUID
            if (-not $uuid) { throw "UUID'siz fatura: $($x.Name)" }
            if ($mevcut.Contains($uuid)) { $faturaVar++; continue }
            [void]$mevcut.Add($uuid)            # ayni ZIP icinde iki kez gelirse de bir kez ekle
            $faturaYeni++
            foreach ($s in $satirlar) { $yeniSatir.Add($s) }
        }
        $ozet = "XML: $($xmls.Count), yeni fatura: $faturaYeni ($($yeniSatir.Count) kalem), zaten tabloda: $faturaVar, fatura olmayan XML: $faturaDegil"
        if ($Deneme) { Write-Log "[DENEME] $($Zip.Name) -> $ozet (SQL'e yazilmadi)"; return }

        if ($yeniSatir.Count -gt 0) {
            # Hedef tabloda olan kolonlarla DataTable kur
            $kullan = @($KOLONLAR | Where-Object { $Schema.ContainsKey($_.Ad) -and -not $Schema[$_.Ad].Auto })
            $dt = New-Object System.Data.DataTable
            foreach ($k in $kullan) { [void]$dt.Columns.Add($k.Ad, $Schema[$k.Ad].Type) }
            foreach ($s in $yeniSatir) {
                $dr = $dt.NewRow()
                foreach ($k in $kullan) {
                    try { $dr[$k.Ad] = Convert-ForColumn $s[$k.Ad] $k.Tur $Schema[$k.Ad] $k.Ad }
                    catch { throw "Fatura $($s.FaturaNo) kalem $($s.SatirNo), kolon $($k.Ad): $($_.Exception.Message)" }
                }
                $dt.Rows.Add($dr)
            }
            $tran = $Conn.BeginTransaction()
            try {
                $opt = [System.Data.SqlClient.SqlBulkCopyOptions]::CheckConstraints -bor [System.Data.SqlClient.SqlBulkCopyOptions]::FireTriggers
                $bc = New-Object System.Data.SqlClient.SqlBulkCopy($Conn, $opt, $tran)
                $bc.DestinationTableName = $tablo
                $bc.BulkCopyTimeout = 600
                foreach ($k in $kullan) { [void]$bc.ColumnMappings.Add($k.Ad, $k.Ad) }
                $bc.WriteToServer($dt)
                $tran.Commit()
            } catch { try { $tran.Rollback() } catch {}; throw }
        }
        $hedef = Join-Path $dir.Islenen ('{0}_{1}' -f (Get-Date -Format 'yyyyMMdd_HHmmss'), $Zip.Name)
        Move-Item -LiteralPath $Zip.FullName -Destination $hedef
        Write-Log "TAMAM $($Zip.Name) -> $ozet"
    } finally {
        if (Test-Path -LiteralPath $calisma) { Remove-Item -LiteralPath $calisma -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

# ---------------------------------------------------------------- ana akis
$kilitYolu = Join-Path $dir.Gecici '.calisiyor.kilit'
try { $kilit = [IO.File]::Open($kilitYolu, 'OpenOrCreate', 'ReadWrite', 'None') }
catch { Write-Host 'Baska bir aktarim zaten calisiyor, cikiliyor.'; exit 0 }

$hataVar = $false
try {
    $zipler = @(Get-ChildItem -LiteralPath $dir.Gelen -File | Where-Object { $_.Extension -eq '.zip' } | Sort-Object LastWriteTime)
    $hazir = @($zipler | Where-Object { Test-FileReady $_ })
    if ($zipler.Count -gt $hazir.Count) { Write-Host "$($zipler.Count - $hazir.Count) ZIP henuz kopyalaniyor, sonraki turda alinacak." }
    if (-not $hazir.Count) { Write-Host 'Islenecek ZIP yok.'; return }

    $conn = $null; $schema = $null
    try {
        $conn = New-SqlConnection
        $schema = Get-TargetSchema $conn
        if (-not $schema.ContainsKey('UUID')) { throw "Tabloda UUID kolonu yok; cift kayit kontrolu icin gerekli." }
        $eksik = @($KOLONLAR | Where-Object { -not $schema.ContainsKey($_.Ad) } | ForEach-Object { $_.Ad })
        if ($eksik.Count) { Write-Log ("Tabloda olmayan kolonlar yazilmayacak: " + ($eksik -join ', ')) 'UYARI' }
    } catch {
        Write-Log "SQL baglantisi / tablo hatasi: $($_.Exception.Message)" 'HATA'
        exit 2
    }

    foreach ($z in $hazir) {
        try { Invoke-Zip $z $conn $schema }
        catch {
            $hataVar = $true
            Write-Log "HATA $($z.Name): $($_.Exception.Message)" 'HATA'
            if (-not $Deneme -and (Test-Path -LiteralPath $z.FullName)) {
                $h = Join-Path $dir.Hatali ('{0}_{1}' -f (Get-Date -Format 'yyyyMMdd_HHmmss'), $z.Name)
                Move-Item -LiteralPath $z.FullName -Destination $h
                Set-Content -LiteralPath ($h + '.hata.txt') -Value $_.Exception.Message -Encoding UTF8
            }
        }
    }
    if ($conn) { $conn.Close() }
} finally {
    $kilit.Close()
    Remove-Item -LiteralPath $kilitYolu -Force -ErrorAction SilentlyContinue
}
if ($hataVar) { exit 1 }
