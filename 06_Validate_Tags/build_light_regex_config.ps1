cd $PSScriptRoot

#$xmlWriter = New-Object System.XMl.XmlTextWriter("\\QAMV3-SFIL102\Home\mch107\My Documents\14_PS\NOC_AIM\light_regex_config.xml", $Null)
#$xmlWriter.Formatting = 'Indented'
#$xmlWriter.Indentation = 1
#$XmlWriter.IndentChar = "`t"
#$xmlWriter.WriteStartDocument()
#$xmlWriter.WriteStartElement("Patterns")
#$xmlWriter.WriteAttributeString("version","5.0")# catalog Start Node
#
#	$xmlWriter.WriteStartElement("Pattern")
#	$xmlWriter.WriteAttributeString("Regexp",$regex)	
#	$xmlWriter.WriteAttributeString("NT_ID",$namingTemplate_xml.id)
#	$xmlWriter.WriteEndElement()
#$xmlWriter.WriteEndElement()  # catalog end node.
#$xmlWriter.Flush()
#$xmlWriter.Close()

$ism_file = Get-ChildItem -Path .\ -Filter *ISM*.xml | Sort-Object | Select-Object -Last 1
[xml]$XmlDocument = Get-Content $ism_file.FullName
$XmlNamespace = @{ a = "http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01"}
$NamingTemplates = Select-Xml -Xml $XmlDocument  -XPath "/a:ClassLibrary/a:ReferenceData/a:NamingAndNumbering/a:Templates" -Namespace $XmlNamespace 
$regex_list = @()
"Regexp;Naming_template_ID" |  Out-File .\Light_regex.csv

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
    $regex_list += $regex + ";" + $namingTemplate_xml.id

}

$regex_list |  Out-File .\Light_regex.csv -Append