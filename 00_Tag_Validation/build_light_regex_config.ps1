param(
	[Parameter(Mandatory=$true)]
	[string]$ism_file 
)
$ism_file = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\RegisterGateway\source data\class library\clib.xml"
Set-Location $PSScriptRoot
# $ism_file = Get-ChildItem -Path .\ -Filter *ISM*.xml | Sort-Object | Select-Object -Last 1
[xml]$XmlDocument = Get-Content $ism_file
$XmlNamespace = @{ a = "http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01"
                   b = "http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Nomenclature/2015/09"}
$NamingTemplates = Select-Xml -Xml $XmlDocument  -XPath "/a:ClassLibrary/a:ReferenceData/a:NamingAndNumbering/a:Templates" -Namespace $XmlNamespace 


$regex_list = @{}
#$regex_list.add("Regexp","Naming_template_ID")
$congig_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\06_Regexp_configs\Light_regex.csv"
Write-Host "Updating Light REGEXP configuration file in $config_path" -BackgroundColor DarkYellow
"Regexp;Naming_template_ID;Discipline" |  Out-File $congig_path 

foreach($namingTemplate_xml in $namingTemplates.Node.ChildNodes)
{

	if($namingTemplate_xml.obsolete)
	{
		$mess = $namingTemplate_xml.id + " skipped as obsolete" 
		continue
	}
	if([string]::IsNullOrEmpty($namingTemplate_xml.InnerXML))
	{
		$mess = $namingTemplate_xml.id + " skipped as no Naming elements" 
		continue		
	}
	if(Select-Xml -XML $namingTemplate_xml -XPath "./a:Elements/a:Element[@id='EI000087']" -Namespace $XmlNamespace)
	{
		$mess = $namingTemplate_xml.id + " skipped as because of reference" 
		continue
	}
    $discipline = $namingTemplate_xml.Extension.Discipline
	$regex = ""
	$NT_Elems_ids = Select-Xml -Xml $namingTemplate_xml -XPath "./a:Elements/a:Element" -Namespace $XmlNamespace
    $NT_Elems_ids = $NT_Elems_ids | Sort-Object {$_.Node.sortOrder}

    $last_Naming_Element = $NT_Elems_ids | Sort-Object {$_.Node.sortorder} | Select-Object -Last 1
    
	foreach($NT_naming_element in $NT_Elems_ids)
	{

		if($NT_naming_element.node.regEx)
        {
            $regex_val= $NT_naming_element.node.regEx
        }
        else
        {
		    $elem_regex_xpath = "//a:NamingAndNumbering/a:Elements/a:Element[@id='" + $NT_naming_element.node.id + "']" 
		    $elem_regex = Select-Xml -Xml $namingTemplate_xml -XPath $elem_regex_xpath -Namespace $XmlNamespace
            $regex_val = $elem_regex.Node.regEx
        }
		if(!$regex_val)
        {
            $regex = ''
			$message = "No regular expression found on " + $NT_naming_element.Node.id + ", Namint template " + $namingTemplate_xml.id + " will be skipped..."
            echo $message
			break
        }
		if($regex_val -Match '\|'){$regex_val = '('+$regex_val+')'}
		
		$elem_suffix = $NT_naming_element.node.suffix
		$elem_prefix = $NT_naming_element.node.prefix
        if([string]::IsNullOrEmpty($NT_naming_element.Node.mandatory)){
            $elem_mandatory = [System.Convert]::ToBoolean("True")
        }
        else{
            $elem_mandatory = [System.Convert]::ToBoolean($NT_naming_element.node.mandatory) 
        }

        if($elem_mandatory)
		{
			$val = $elem_prefix + $regex_val + $elem_suffix
		}
		else
		{
			$val = '(' + $elem_prefix + $regex_val + $elem_suffix + ')?'
		}

		$regex = $regex + $val
        if($NT_naming_element.Node.sortOrder -eq  $last_Naming_Element.node.sortorder){
            $regex = "^" + $regex + "$"
            
        }
	}
	if(!$regex){continue}
    if(!$regex_list[$regex]){$regex_list.add($regex,$namingTemplate_xml.id + ';' + $discipline)
}

}


foreach($key in $regex_list.Keys){
    $key + ";" + $regex_list[$key] | Out-File $congig_path -Append
}