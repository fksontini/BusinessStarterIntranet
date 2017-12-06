[CmdletBinding()]
Param(
	[Parameter(Mandatory=$True,Position=1)]
	[string]$SiteUrl,

	[Parameter(Mandatory=$True)]
	[string]$UserName,

	[Parameter(Mandatory=$True)]
	[string]$Password,
	
	[Parameter(Mandatory=$False)]
	[switch]$Prod=$false,
	
	[Parameter(Mandatory=$False)]
	[switch]$IncludeData=$false
)

# -----------------
# Global parameters
# -----------------

# Include utility functions
 . "./utility/Utility.ps1"

$0 = $myInvocation.MyCommand.Definition
$CommandDirectory = [System.IO.Path]::GetDirectoryName($0)

# Configuration file paths
$ProvisioningRootSiteTemplateFile = ($CommandDirectory  + ".\provisioning\RootSiteTemplate.xml")
$SearchConfigurationFilePath = ($CommandDirectory  + ".\provisioning\SearchConfiguration.xml")
$ImageRenditionsConfigurationFilePath = ($CommandDirectory + ".\provisioning\PublishingImageRenditions.xml")

# This name will be used to create a separated folder in the style library and the master page catalog.
# If you change this name, don't forget to update :
# - Links in the master page (CSS and JS files)
# - Web Parts XML contents on the provisioning template (display templates paths)
# - Display templates files (relative paths to hover panel display template)
$AppFolderName = "PnP"

# Connect to the site
$PasswordAsSecure = ConvertTo-SecureString $Password -AsPlainText -Force
$Credentials = New-Object System.Management.Automation.PSCredential ($UserName , $PasswordAsSecure)
Connect-PnPOnline -Url $SiteUrl -Credentials $Credentials

# Determine the SharePoint version
$ServerVersion = (Get-PnPContext).ServerLibraryVersion.Major

switch ($ServerVersion) 
{ 
	15 {$AssemblyVersion = "15.0.0.0"} 
	16 {$AssemblyVersion = "16.0.0.0"} 
    default {$AssemblyVersion = "16.0.0.0"}
}

# -------------------------------------------------------------------------------------
# Set the correct SharePoint assembly version in .aspx and .master files regarding the server version
# -------------------------------------------------------------------------------------
Get-ChildItem -Path ".\provisioning\artefacts" -Include "*.aspx","*.master" -Recurse | ForEach-Object {

    (Get-Content -Path $_.FullName) -replace "1[5|6]\.0\.0\.0",$AssemblyVersion | Out-File -FilePath $_.FullName
}

# -------------------------------------------------------------------------------------
# Upload files in the style library (folders are created automatically by the PnP cmdlet)
# -------------------------------------------------------------------------------------
Push-Location ".\app"

if ($Prod.IsPresent) {
		
	Write-Host "1# Bundling the application (production mode)..." -ForegroundColor Magenta
	
	# Bundle the project in production mode (the '2>$null' is to avoid PowerShell ISE errors)
	webpack -p 2>$null
		
} else {

	Write-Host "1# Bundling the application (development mode)..." -ForegroundColor Magenta
	
	# Bundle the project in dev mode
	webpack 2>$null
}

Pop-Location

# Get Webpack output folder and upload all files in the style library (eventually will be replaced by CDN in the future)
$DistFolder = $CommandDirectory + "\app\dist"

Write-Host "2# Uploading all files (non optimized)..." -ForegroundColor Magenta

Push-Location $DistFolder 

Get-ChildItem -Recurse $DistFolder -File | ForEach-Object {

    $TargetFolder = "Style Library\$AppFolderName\" + (Resolve-Path -relative $_.FullName) | Split-Path -Parent

	$varFile = Add-PnPFile -Path $_.FullName -Folder ($TargetFolder.Replace("\","/")).Replace("./","").Replace(".","") -Checkout
}

Pop-Location

# -------------------------------------------------------------------------------------
# Apply root site template
# -------------------------------------------------------------------------------------
Write-Host "3# Apply the provisioning template to the root site..." -ForegroundColor Magenta

# Create news and events folders in the "Pages" library
Ensure-PnPFolder -SiteRelativePath "Sidor/News" | Out-Null
Ensure-PnPFolder -SiteRelativePath "Sidor/Events" | Out-Null

# Apply the root site provisioning template
Apply-PnPProvisioningTemplate -Path $ProvisioningRootSiteTemplateFile -Parameters @{ "CompanyName" = $AppFolderName; "AssemblyVersion" = $AssemblyVersion; }

# Enable Item Scheduling feature on the "Pages" library
Enable-CustomItemScheduling -Web (Get-PnPWeb) -PagesLibraryName "Sidor"

# Content Types order
$ContentTypesOrderRoot = @(

	[PSCustomObject]@{FolderName="Sidor";ContentTypes=@("Home Page","Static Page","Search Page")},
	[PSCustomObject]@{FolderName="Sidor/News";ContentTypes=@("News Page")}
	[PSCustomObject]@{FolderName="Sidor/Events";ContentTypes=@("Event Page")}
)

$ContentTypesOrderRoot | Foreach-Object { Set-FolderContentTypesOrder -FolderRelativePath $_.FolderName -ContentTypes $_.ContentTypes }

# Set up the search configuration
# Be careful, in SharePoint Online, we can't update an automatically created managed property to be sortable. We have to use Refinable<Type>XX predefined property.
# For example, for the news list on the front page, we use the RefinableDate00 property for the publishing date. Use an alias instead of using the default name.
Set-PnPSearchConfiguration -Path $SearchConfigurationFilePath -Scope Site

Write-Host "4# Publishing artefacts..." -ForegroundColor Magenta

# Publishing artefacts
$Site = Get-PnPSite
$SiteServerRelativeUrl = Get-PnPProperty -ClientObject $Site -Property ServerRelativeUrl

$FilesToPublish = @(
	# Master pages
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/$AppFolderName/portal.master"},
			
	# Page Layouts
    [PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/$AppFolderName/EventPageLayout.aspx"},
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/$AppFolderName/NewsPageLayout.aspx"},
    [PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/$AppFolderName/StaticPageLayout.aspx"},
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/$AppFolderName/HomePageLayout.aspx"},
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/$AppFolderName/SearchPageLayout.aspx"},

	# Display Templates
    [PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Content Web Parts/$AppFolderName/Item_Intranet-News.html"},
    [PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Content Web Parts/$AppFolderName/Item_Intranet-News-Tile.html"},		
    [PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Content Web Parts/$AppFolderName/Item_Intranet-Event.html"},	
    [PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Content Web Parts/$AppFolderName/Item_Intranet-Document.html"},
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Content Web Parts/$AppFolderName/Item_Intranet-Contact.html"},
    [PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Content Web Parts/$AppFolderName/Control_Intranet-List_Paging.html"},	
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Content Web Parts/$AppFolderName/Control_Intranet-List_NoPaging.html"},	
    [PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Content Web Parts/$AppFolderName/Control_Intranet_Tiles_List.html"},			
    [PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Search/$AppFolderName/Item_Intranet-News_Search.html"},
    [PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Search/$AppFolderName/Item_Intranet-Event_Search.html"},	
    [PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Search/$AppFolderName/Item_Intranet-Page_Search.html"},
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Search/$AppFolderName/Control_Intranet-SearchResults.html"},
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Filters/$AppFolderName/Filter_Intranet-Item.html"},
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Filters$AppFolderName/Filter_Intranet-SliderBarGraph.html"},
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Filters/$AppFolderName/Control_Intranet-Refinement.html"},
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Search/$AppFolderName/Item_Intranet-Document_HoverPanel.html"},  
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Search/$AppFolderName/Item_Intranet-Document_Search.html"},
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Search/$AppFolderName/Item_Intranet_CommonHoverPanel_Actions.html"},  
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Search/$AppFolderName/Item_Intranet_CommonHoverPanel_Body.html"},  
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Search/$AppFolderName/Item_Intranet_CommonHoverPanel_Header.html"},  		
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/_catalogs/masterpage/display templates/Search/$AppFolderName/Item_Intranet_WebPage_HoverPanel.html"},

	# Pages  	
    [PSCustomObject]@{Url="$SiteServerRelativeUrl/Sidor/Home.aspx"},  
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/Sidor/Search.aspx"},  
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/Sidor/SearchDocuments.aspx"},
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/Sidor/Hem.aspx"},  
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/Sidor/Sok.aspx"},  
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/Sidor/SokDokument.aspx"}    	
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/Sidor/SokPersoner.aspx"},  
	[PSCustomObject]@{Url="$SiteServerRelativeUrl/Sidor/SearchPeople.aspx"}   
)

$FilesToPublish | ForEach-Object {

    Set-PnPFileCheckedOut -Url $_.Url
    Set-PnPFileCheckedIn -Url $_.Url -CheckinType MajorCheckIn
}

# Approve all items 
Get-PnPListItem -List Sidor | ForEach-Object { 
    $_["_ModerationStatus"] = 0
    $_.Update()
}

Execute-PnPQuery

# Reset the theme
Set-PnPTheme

# Set the theme
$Web = Get-PnPWeb
$bgImageUrl = Out-Null
$fontScheme = Out-Null
$Web.ApplyTheme("$SiteServerRelativeUrl/_catalogs/theme/15/intranet.spcolor", $fontScheme, $bgImageUrl, $true)

Execute-PnPQuery

# -------------------------------------------------------------------------------------
# 3) Taxonomy setup
# -------------------------------------------------------------------------------------
Write-Host "5# Set up taxonomy..." -ForegroundColor Magenta

# Get the site collection term group name
$CurrentSite = Get-PnPSite
$Session = Get-PnPTaxonomySession
$TermStore = $Session.GetDefaultSiteCollectionTermStore();
$SiteCollectionTermGroup = $TermStore.GetSiteCollectionGroup($CurrentSite, $false)
$IntranetTermGroupName = Get-PnPProperty -ClientObject $SiteCollectionTermGroup -Property Name 

$SiteMapTermSetName_EN = "Site Map EN"
$SiteMapTermSetName_SV = "Site Map SV"

$HeaderLinksTermSetName_EN = "Header Links EN"
$HeaderLinksTermSetName_SV = "Header Links SV"

$FooterLinksTermSetName_EN = "Footer Links EN"
$FooterLinksTermSetName_SV = "Footer Links SV"

# Get navigation term sets for each language (SV & EN)
$SiteMapTermSet_EN = Get-PnPTaxonomyItem -Term "$IntranetTermGroupName|$SiteMapTermSetName_EN"
$SiteMapTermSetId_EN = $SiteMapTermSet_EN.Id

$SiteMapTermSet_SV = Get-PnPTaxonomyItem -Term "$IntranetTermGroupName|$SiteMapTermSetName_SV"
$SiteMapTermSetId_SV = $SiteMapTermSet_SV.Id

# Duplicate the Site Map EN into Site Map FR to have a mirror structure (i.e pin terms with children)
$SiteMapTermSetTerms_SV = Get-PnPProperty -ClientObject $SiteMapTermSet_SV -Property Terms

$SiteMapTermSetTerms_SV | ForEach-Object {

	$NavTerm = Get-PnPTaxonomyItem -Term ("$IntranetTermGroupName|$SiteMapTermSetName_EN|" + $_.Name)

    if ($NavTerm -eq $null) {

		$Reuse = $SiteMapTermSet_EN.ReuseTermWithPinning($_)

		Execute-PnPQuery
	}
}

# Do the same thing for header links term set
$HeaderLinksTermSet_EN = Get-PnPTaxonomyItem -Term "$IntranetTermGroupName|$HeaderLinksTermSetName_EN"
$HeaderLinksTermSetId_EN = $HeaderLinksTermSet_EN.Id

$HeaderLinksTermSet_SV = Get-PnPTaxonomyItem -Term "$IntranetTermGroupName|$HeaderLinksTermSetName_SV"
$HeaderLinksTermSetId_SV = $HeaderLinksTermSet_SV.Id

$HeaderLinksTermSetTerms_SV = Get-PnPProperty -ClientObject $HeaderLinksTermSet_SV -Property Terms

$HeaderLinksTermSetTerms_SV | ForEach-Object {

	$NavTerm = Get-PnPTaxonomyItem -Term ("$IntranetTermGroupName|$HeaderLinksTermSetName_EN|" + $_.Name)

    if ($NavTerm -eq $null) {

		$Reuse = $HeaderLinksTermSet_EN.ReuseTermWithPinning($_)

		Execute-PnPQuery
	}
}

# ...and for the footer links term set also
$FooterLinksTermSet_EN = Get-PnPTaxonomyItem -Term "$IntranetTermGroupName|$FooterLinksTermSetName_EN"
$FooterLinksTermSetId_EN = $FooterLinksTermSet_EN.Id

$FooterLinksTermSet_SV = Get-PnPTaxonomyItem -Term "$IntranetTermGroupName|$FooterLinksTermSetName_SV"
$FooterLinksTermSetId_SV = $FooterLinksTermSet_SV.Id

$FooterLinksTermSetTerms_SV = Get-PnPProperty -ClientObject $FooterLinksTermSet_SV -Property Terms

$FooterLinksTermSetTerms_SV | ForEach-Object {

	$NavTerm = Get-PnPTaxonomyItem -Term ("$IntranetTermGroupName|$FooterLinksTermSetName_EN|" + $_.Name)

    if ($NavTerm -eq $null) {

		$Reuse = $FooterLinksTermSet_EN.ReuseTermWithPinning($_)

		Execute-PnPQuery
	}
}


# -------------------------------------------------------------------------------------
# Setup the configuration list
# -------------------------------------------------------------------------------------
Write-Host "6# Setup the configuration list..." -ForegroundColor Magenta

$ConfigurationList = Get-PnPList -Identity "Konfiguration"

$ConfigurationItems = @(

	@{ "Title"="Default EN";"ForceCacheRefresh"=1;"SiteMapTermSetId"=$SiteMapTermSetId_EN;"HeaderLinksTermSetId"=$HeaderLinksTermSetId_EN;"FooterLinksTermSetId"=$FooterLinksTermSetId_EN;"IntranetContentLanguage"="EN" },
	@{ "Title"="Default SV";"ForceCacheRefresh"=1;"SiteMapTermSetId"=$SiteMapTermSetId_SV;"HeaderLinksTermSetId"=$HeaderLinksTermSetId_SV;"FooterLinksTermSetId"=$FooterLinksTermSetId_SV;"IntranetContentLanguage"="SV" }
)

# Create the configuration item for each language
$ConfigurationItems | ForEach-Object {

    # We create items in two steps because of a bug with the Add-PnPListItem since the February release https://github.com/SharePoint/PnP-PowerShell/issues/778
    $Item = Add-PnPListItem -List $ConfigurationList
    $Item = Set-PnPListItem -Identity  $Item.Id -List $ConfigurationList -Values $_ -ContentType "Item"
}

# -------------------------------------------------------------------------------------
# Add image renditions
# -------------------------------------------------------------------------------------
Write-Host "7# Configure image renditions..." -ForegroundColor Magenta

# Thanks to http://www.eliostruyf.com/provision-image-renditions-to-your-sharepoint-2013-site/
$File = Add-PnPFile -Path $ImageRenditionsConfigurationFilePath -Folder "_catalogs\masterpage\" -Checkout

# -------------------------------------------------------------------------------------
# Add sample data
# -------------------------------------------------------------------------------------
if ($IncludeData.IsPresent) {
	$CarouselItemsList = Get-PnPList -Identity "Carousel Items"

    $ConfigurationItemsEN = @(

	    @{ "Title"="Our offers: What we as a company can offer";"CarouselItemURL"="http://www.dizparc.se/vara-erbjudanden";"CarouselItemImage"="$SiteServerRelativeUrl/Style Library/PnP/img/kundvard.jpg";"IntranetContentLanguage"="EN" },
	    @{ "Title"="Customers: Customer testimonials and real life cases";"CarouselItemURL"="http://www.dizparc.se/kunder";"CarouselItemImage"="$SiteServerRelativeUrl/Style Library/PnP/img/kundlista.jpg";"IntranetContentLanguage"="EN" },
        @{ "Title"="Work at dizparc: How to get in touch";"CarouselItemURL"="http://www.dizparc.se/jobba-hos-oss";"CarouselItemImage"="$SiteServerRelativeUrl/Style Library/PnP/img/arenden.jpg";"IntranetContentLanguage"="EN" }
		  
    )

    $ConfigurationItemsSV = @(

	    @{ "Title"="Våra erbjudanden: Vad kan vi som företag erbjuda";"CarouselItemURL"="http://www.dizparc.se/vara-erbjudanden";"CarouselItemImage"="$SiteServerRelativeUrl/Style Library/PnP/img/kundvard.jpg";"IntranetContentLanguage"="SV" },
	    @{ "Title"="Kunder: De kunder och case vi arbetar med";"CarouselItemURL"="http://www.dizparc.se/kunder";"CarouselItemImage"="$SiteServerRelativeUrl/Style Library/PnP/img/kundlista.jpg";"IntranetContentLanguage"="SV" },
        @{ "Title"="Arbeta på dizparc: Hur kommer du i kontakt med oss?";"CarouselItemURL"="http://www.dizparc.se/jobba-hos-oss";"CarouselItemImage"="$SiteServerRelativeUrl/Style Library/PnP/img/arenden.jpg";"IntranetContentLanguage"="SV" }
	  )

    # $CarouselItemsList = Get-PnPList -Identity "Carousel Items"

    # $ConfigurationItemsEN = @(

	#     @{ "Title"="Part 1: Functional overview (How to use the solution?)";"CarouselItemURL"="http://thecollaborationcorner.com/2016/08/22/part-1-functional-overview-how-to-use-the-solution";"CarouselItemImage"="http://thecollaborationcorner.com/wp-content/uploads/2016/08/part1.png";"IntranetContentLanguage"="EN" },
	#     @{ "Title"="Part 2: Frameworks and libraries used (How it is implemented?)";"CarouselItemURL"="http://thecollaborationcorner.com/2016/08/25/part-2-frameworks-and-libraries-used-how-it-is-implemented";"CarouselItemImage"="http://thecollaborationcorner.com/wp-content/uploads/2016/08/part2.png";"IntranetContentLanguage"="EN" },
    #     @{ "Title"="Part 3: Design and mobile implementation";"CarouselItemURL"="http://thecollaborationcorner.com/2016/08/29/part-3-design-and-mobile-implementation";"CarouselItemImage"="http://thecollaborationcorner.com/wp-content/uploads/2016/08/part3.png";"IntranetContentLanguage"="EN" },
    #     @{ "Title"="Part 4: The navigation implementation";"CarouselItemURL"="http://thecollaborationcorner.com/2016/08/31/part-4-the-navigation-implementation";"CarouselItemImage"="http://thecollaborationcorner.com/wp-content/uploads/2016/08/part4.png";"IntranetContentLanguage"="EN" },    
    #     @{ "Title"="Part 5: Localization";"CarouselItemURL"="http://thecollaborationcorner.com/2016/09/02/part-5-localization";"CarouselItemImage"="http://thecollaborationcorner.com/wp-content/uploads/2016/09/part5.png";"IntranetContentLanguage"="EN" },  
    #     @{ "Title"="Part 6: The search implementation";"CarouselItemURL"="http://thecollaborationcorner.com/2016/09/08/part-6-the-search-implementation";"CarouselItemImage"="http://thecollaborationcorner.com/wp-content/uploads/2016/09/part6.png";"IntranetContentLanguage"="EN" }  
    # )

    # $ConfigurationItemsSV = @(

	#     @{ "Title"="Del 1: Aperçu fonctionel (Comment utiliser cette solution?)";"CarouselItemURL"="http://thecollaborationcorner.com/2016/08/22/part-1-functional-overview-how-to-use-the-solution";"CarouselItemImage"="http://thecollaborationcorner.com/wp-content/uploads/2016/08/part1.png";"IntranetContentLanguage"="SV" },
	#     @{ "Title"="Del 2: Frameworks et librairies utilisées";"CarouselItemURL"="http://thecollaborationcorner.com/2016/08/25/part-2-frameworks-and-libraries-used-how-it-is-implemented";"CarouselItemImage"="http://thecollaborationcorner.com/wp-content/uploads/2016/08/part2.png";"IntranetContentLanguage"="SV" },
    #     @{ "Title"="Del 3: Identité visuelle et implémentation mobile";"CarouselItemURL"="http://thecollaborationcorner.com/2016/08/29/part-3-design-and-mobile-implementation";"CarouselItemImage"="http://thecollaborationcorner.com/wp-content/uploads/2016/08/part3.png";"IntranetContentLanguage"="SV" },
    #     @{ "Title"="Del 4: Implémentation de la navigation";"CarouselItemURL"="http://thecollaborationcorner.com/2016/08/31/part-4-the-navigation-implementation";"CarouselItemImage"="http://thecollaborationcorner.com/wp-content/uploads/2016/08/part4.png";"IntranetContentLanguage"="SV" },    
    #     @{ "Title"="Del 5: Multilinguisme";"CarouselItemURL"="http://thecollaborationcorner.com/2016/09/02/part-5-localization";"CarouselItemImage"="http://thecollaborationcorner.com/wp-content/uploads/2016/09/part5.png";"IntranetContentLanguage"="SV" },  
    #     @{ "Title"="Del 6: Implémentation de la recherche";"CarouselItemURL"="http://thecollaborationcorner.com/2016/09/08/part-6-the-search-implementation";"CarouselItemImage"="http://thecollaborationcorner.com/wp-content/uploads/2016/09/part6.png";"IntranetContentLanguage"="SV" }  
    # )

    Write-Host "8# Add carousel data..." -ForegroundColor Magenta

    # Create the configuration item for each language
    $ConfigurationItemsEN | ForEach-Object {

		$Item = Add-PnPListItem -List $CarouselItemsList
    	$Item = Set-PnPListItem -Identity  $Item.Id -List $CarouselItemsList -Values $_ -ContentType "Carousel Item"
    }

    $ConfigurationItemsSV | ForEach-Object {

		$Item = Add-PnPListItem -List $CarouselItemsList
    	$Item = Set-PnPListItem -Identity  $Item.Id -List $CarouselItemsList -Values $_ -ContentType "Carousel Item"
    }

    # Add promoted links
    $PromotedLinksList = Get-PnPList -Identity "Länkar"
    $PromotedLinks = @(

	    @{ "Title"="Presentationer";"BackgroundImageLocation"="$SiteServerRelativeUrl/Style Library/PnP/img/kundvard.jpg";"LinkLocation"="http://www.dizparc.se"},
	    @{ "Title"="HR";"BackgroundImageLocation"="$SiteServerRelativeUrl/Style Library/PnP/img/kundlista.jpg";"LinkLocation"="http://www.dizparc.se"},
	    @{ "Title"="Marknadsföring";"BackgroundImageLocation"="$SiteServerRelativeUrl/Style Library/PnP/img/arenden.jpg";"LinkLocation"="http://www.dizparc.se"},
	    @{ "Title"="Leverantörsavtal";"BackgroundImageLocation"="$SiteServerRelativeUrl/Style Library/PnP/img/tidredovisning.jpg";"LinkLocation"="http://www.dizparc.se"}
		
    )
	
    $PromotedLinks | ForEach-Object {

		$Item = Add-PnPListItem -List $PromotedLinksList
    	$Item = Set-PnPListItem -Identity  $Item.Id -List $PromotedLinksList -Values $_ -ContentType "Item"
    }
}

Write-Host "Done!" -ForegroundColor Green

# Close the connection to the server
Disconnect-PnPOnline


