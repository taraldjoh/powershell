# Script for oppretting av standardisert Teams område
# Versjon 1.0
# Utgiver: Tarald Johansen
# Dependencies: Azure module, Teams Module, ExchangeOnline module, SharepointOnline psmodule, PnPOnline module
# Servicekonto som benyttest vil stå som teameier og må ha følgende roller i Azure:
# - Exchange Administrator
# - SharePoint Administrator
# - Teams Administrator
# - Directory Writers
# NB! Vil erstattes med app-basert autentisering i Azure
$SPOAdminSite = "https://site-admin.sharepoint.com"
$PnPSite = "https://site.sharepoint.com"
$TeamsPolicyName = "PolicyName"

$365GroupDisplayName = 'Teknisk (G-602)'                                        # Styrer visningsnavn til Azure gruppe/Team
$365GroupMailNickname = "Teams_G-602-Teknisk"                                   # Setter mailnick til Azure gruppen
$365GroupDescription = "Einingsgruppe for Teknisk"                              # Beskrivelse av type team, f.eks 'Einingsgruppe'
$365DynamicMembershipGroupObjectId = '""'   # ObjectId til AD-gruppe som brukes for automatisk tilgangsstyring 
$teamsImagePath = ""                                  # Filsti til logo som settes på Teamet

$svc = Get-Credential -Credential #UserPrincipalName#


$OutputEncoding = [ System.Text.Encoding]::UTF8
#

# Sjekker og evt etablerer nødvendige sesjoner scriptet er avhengig av
function connectionsCheck {
    try {
        # Sjekk Azure AD sesjon
        Get-AzureADCurrentSessionInfo -ErrorAction stop -WarningAction SilentlyContinue
        # Sjekk Exchange Online sesjon
        if ((Get-ConnectionInformation).State -ne "Connected") { throw("Ikke tilkoblet Exchange") }
        # Sjekk Teams sesjon
        Get-CsTenant -ErrorAction stop -WarningAction SilentlyContinue | Out-Null
        Write-Host -ForegroundColor Green "Sesjoner sjekket - OK"
        
    }
    catch {
        #Write-Host -ForegroundColor red "Ingen aktiv sesjon!"
        Connect-AzureAD -Credential $svc -InformationAction SilentlyContinue -ErrorAction stop  | Out-Null
        Connect-ExchangeOnline -Credential $svc -InformationAction SilentlyContinue -ErrorAction stop -ShowBanner:$false | Out-Null
        Connect-MicrosoftTeams -Credential $svc -InformationAction SilentlyContinue -ErrorAction Stop | Out-Null
        Connect-SPOService -Url $SPOAdminSite -Credential $svc | Out-Null
        Write-Host -ForegroundColor green "Alle sesjoner etablert"
    }


}
connectionsCheck

# Opprette dynamisk Azure gruppe og tilpasse denne for å opprette einings Team med egendefinerte innstilinger og mal
function opprettEiningsTeam {
    Write-Host -ForegroundColor yellow "Oppretter ny gruppe i Azure"
        
    try {
        # Oppretter ny M365 gruppe i Azure som benyttes videre for opprettelse av Team i Teams
        $res = New-AzureADMSGroup -DisplayName $365GroupDisplayName -Description $365GroupDescription  -MailEnabled $True -SecurityEnabled $True -MailNickname $365GroupMailNickname -GroupTypes "DynamicMembership", "Unified" -MembershipRule "user.memberOf -any (group.objectId -in [$365DynamicMembershipGroupObjectId])" -MembershipRuleProcessingState "On" -Visibility "Private" -ErrorAction stop
        Write-Host -ForegroundColor green "Opprette Azure gruppe status: OK"
        $azureGroupId = $res.id
        
    }
    catch {
        Write-Host -ForegroundColor red "Det skjedde en feil:"
        throw $_
        break
    }
        
    # SYNC CHECK LOOP BEFORE CONTINUE: Sjekker om mailobjektet er opprettet
    while ($True) {
        try {
            $mailGroup = Get-UnifiedGroup -Identity $365GroupMailNickname -ErrorAction Stop -Verbose
            break
        }
        catch {
            Write-Host -ForegroundColor Cyan "Synkroniserer, vennligst vent"
            Start-Sleep -s 10
        }
    }
    

    try {
        # Deaktiverer velkomstmail og skjuler gruppen fra Outlook/Adresseliste(GAL)
        Write-Host -ForegroundColor Yellow "Deaktiverer velkomstmail"
        Set-UnifiedGroup -Identity $365GroupMailNickname -UnifiedGroupWelcomeMessageEnabled:$false -HiddenFromExchangeClientsEnabled:$true -HiddenFromAddressListsEnabled:$true -Verbose
        Write-Host -ForegroundColor Green "Velkomstmail deaktivert"
    }
    catch {
        Write-Host -ForegroundColor red "Det skjedde en feil:"
        throw $_
        break
    }

    # Teams oppretting
    Write-Host -ForegroundColor yellow "Oppretter nytt Team"
    try {
        # Opprett Team basert på opprettet AD gruppe
        $newTeam = New-Team -GroupId $azureGroupId -ErrorAction stop 
        Write-Host -ForegroundColor green "Team opprettet"

    }
    catch {
        Write-Host -ForegroundColor red "Det skjedde en feil:"
        throw $_
    }
    
    # Teams konfigurasjon
    Write-Host -ForegroundColor Yellow "Konfigurerer Teams Policy og innstillinger"
    try {
        # Setter Teams Policy
        New-CsGroupPolicyAssignment -GroupId $newTeam.GroupId -PolicyType TeamsChannelsPolicy -PolicyName $TeamsPolicyName -ErrorAction stop | Out-Null
        # Setter standard bilde på Teamet
        Set-TeamPicture -GroupId $newTeam.GroupId -ImagePath $teamsImagePath -ErrorAction SilentlyContinue | Out-Null
    }
    catch {
        Write-Host -ForegroundColor red "Det skjedde en feil:"
        throw $_
    }
   
    Write-Host -ForegroundColor Green "Konfigurere Teams status: OK"
    
    # SYNC CHECK LOOP BEFORE CONTINUE: Sjekker at Sharepoint bibliotek er opprettet og etablerer sesjon
    Write-Host -ForegroundColor Cyan "Etablere kobling mot underliggende Sharepoint bibliotek"
    while ($True) {
        try {
            $pnpConnection = Connect-PnPOnline -Url "$PnPSite/sites/$365GroupMailNickname" -Credentials $svc 
            break
        }
        catch {
            Write-Host -ForegroundColor Cyan "Synkroniserer, vennligst vent"
            Start-Sleep -s 10
        }
    }

    # Sharepoint konfigurasjon 
    Write-Host -ForegroundColor Yellow "Konfigurerer sharepoint bibliotek"
    try {
        # Legger til 'Versjon' kolonnen
        Write-Host -ForegroundColor Yellow "Konfigurerer versjonsstyring og legger til versjonskolonne"
        $ListName = "Documents"
        $Context = Get-PnPContext 
        $Liste = $Context.Web.Lists.GetByTitle($ListName)
        $ViewFields = $Liste.DefaultView.ViewFields 
        $Context.Load($ViewFields) 
        $Context.ExecuteQuery() 
        
        $Liste.DefaultView.ViewFields.Add("Version")
        $Liste.DefaultView.Update()
        $Context.ExecuteQuery() 
        
        # Aktiverer versjonshåndtering
        Set-PnPList -Identity $ListName -EnableVersioning:$true -EnableMinorVersions:$true -MajorVersions 500 -MinorVersions 25 | Out-Null

        # Setter delingspolicy på sharepoint siten
        Set-SPOSite $Context.Url -DefaultSharingLinkType Internal -DefaultLinkPermission View -AnonymousLinkExpirationInDays 30 -SharingCapability Disabled | Out-Null
    }
    catch {
        Write-Host -ForegroundColor red "Noe gikk galt med konfigurasjonen:"
        Throw $_
    }

    # Endrer standardvisning av valg under 'Ny' knappen 
    Write-Host -ForegroundColor Yellow "Konfigurerer visning av valg under 'Ny' knappen"
    $iconVisibility = $true
    function Add-MenuItem {
        param(
            [Parameter(Mandatory)]$title,
            [Parameter(Mandatory)]$visible,
            [Parameter(Mandatory)]$templateId
        )
        
        $newChildNode = New-Object System.Object
        $newChildNode | Add-Member -type NoteProperty -name title -value $title
        $newChildNode | Add-Member -type NoteProperty -name visible -value $visible
        $newChildNode | Add-Member -type NoteProperty -name templateId -value $templateId
    
        return $newChildNode
    }
    
    function Get-DefaultMenuItems {
        $DefaultMenuItems = @()
        $DefaultMenuItems += Add-MenuItem -title "Folder" -templateId "NewFolder" -visible $iconVisibility 
        $DefaultMenuItems += Add-MenuItem -title "Word document" -templateId "NewDOC" -visible $iconVisibility
        $DefaultMenuItems += Add-MenuItem -title "Excel workbook" -templateId "NewXSL" -visible $iconVisibility 
        $DefaultMenuItems += Add-MenuItem -title "PowerPoint presentation" -templateId "NewPPT" -visible $iconVisibility 
        $DefaultMenuItems += Add-MenuItem -title "OneNote notebook" -templateId "NewONE" -visible $iconVisibility 
        $DefaultMenuItems += Add-MenuItem -title "Visio drawing" -templateId "NewVSDX" -visible $false
        $DefaultMenuItems += Add-MenuItem -title "Forms for Excel" -templateId "NewXSLForm" -visible $false
        $DefaultMenuItems += Add-MenuItem -title "Link" -templateId "Link" -visible $iconVisibility 
       
        return $DefaultMenuItems
    }
    $Context = Get-PnPContext
    $Liste = $Context.Web.Lists.GetByTitle($listName)
    $DefaultView = $Liste.DefaultView
    $MenuItems = Get-DefaultMenuItems
    $Context.Load($DefaultView)
    $Context.ExecuteQuery()


    $DefaultView.NewDocumentTemplates = $MenuItems | ConvertTo-Json
    $DefaultView.Update()
    $Context.ExecuteQuery()

    Write-Host -ForegroundColor Green "Konfigurasjon av sharepoint bibliotek ferdigstilt. Status: OK"

    # Lager Library ID for kopiering inn i Intune ADMX template
    $tenantId = Get-AzureADTenantDetail | Select -Property objectID
    $PnPSite = Get-PnPSite -Includes Id | Select Id, Url
    $PnPWeb = Get-PnPWeb -Includes Id | Select id
    $PnPList = Get-PnPList -Identity "Documents" -Includes Id | Select id 
    
    # Sammenfatter full URL
    $FULLURL = 'tenantId=' + $tenantId.ObjectID + '&siteId={' + $PnPSite.Id + '}&webId={' + $PnPWeb.Id + '}&listId=' + $PnPList.Id + '&webUrl=' + $PnPSite.Url + '&version=1' 

    # Script fullført
    Write-Host -ForegroundColor green "Opprettelsen av einingsteam er gjennomfort uten feilmeldinger."
    Write-Host -ForegroundColor DarkRed "Library ID til Azure:"
    Write-Output $FULLURL 
    Write-Host -ForegroundColor DarkRed "Azure GroupId:"
    Write-Output $azureGroupId     
}

opprettEiningsTeam
 


