# CONFIGURATION
# Chemin vers le fichier CSV d'entrée
$csvPath = "\\casnas100\dfs_root\COMMUN\Commun47\Dossiers communs SRI\10-Gestion Crise Cyber\Astreinte\old\test_Isidore\exemple_contacts_outlook.csv"
# Dossier de sortie pour les fichiers VCF
$outputFolder = "\\casnas100\dfs_root\COMMUN\Commun47\Dossiers communs SRI\10-Gestion Crise Cyber\Astreinte\old\test_Isidore"


# CRÉATION DU DOSSIER DE SORTIE
New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null

# CHARGER LE CSV avec délimiteur point-virgule
$contacts = Get-Content $csvPath | ConvertFrom-Csv -Delimiter ';'

# BOUCLE SUR CHAQUE CONTACT
foreach ($contact in $contacts) {
    $firstName = $contact.Prenom
    $lastName = $contact.Nom
    $fullName = "$firstName $lastName".Trim()
    $company = $contact.Societe
    $title = $contact.Poste
    $email = $contact."E-mail"
    $mobile = $contact."Telephone mobile"
    $phone = $contact."Telephone fixe"
    $notes = $contact.Notes

    # GÉNÉRATION DU CONTENU VCF
    $vcf = @()
    $vcf += "BEGIN:VCARD"
    $vcf += "VERSION:3.0"
    $vcf += "N:$lastName;$firstName;;;"
    $vcf += "FN:$fullName"
    if ($company) { $vcf += "ORG:$company" }
    if ($title) { $vcf += "TITLE:$title" }
    if ($email) { $vcf += "EMAIL;TYPE=INTERNET:$email" }
    if ($mobile) { $vcf += "TEL;TYPE=CELL:$mobile" }
    if ($phone) { $vcf += "TEL;TYPE=WORK,VOICE:$phone" }
    if ($notes) { $vcf += "NOTE:$notes" }
    $vcf += "END:VCARD"

    # NETTOYER LE NOM DU FICHIER
    $safeName = ($fullName -replace '[^\w\d]', '_')
    $vcfPath = Join-Path -Path $outputFolder -ChildPath "$safeName.vcf"

    # ÉCRITURE DU FICHIER EN UTF-8 SANS BOM
    $utf8NoBom = New-Object System.Text.UTF8Encoding($False)
    [System.IO.File]::WriteAllLines($vcfPath, $vcf, $utf8NoBom)
}
