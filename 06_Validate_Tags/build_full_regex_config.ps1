cd $PSScriptRoot



#Read RDL for Tag validation
$ism_file = Get-ChildItem -Path .\ -Filter *ISM*.xml | Sort-Object | Select-Object -Last 1
[xml]$XmlDocument = Get-Content $ism_file.FullName
$XmlNamespace = @{ a = "http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01"}

$NamingTemplates = Select-Xml -Xml $XmlDocument  -XPath "/a:ClassLibrary/a:ReferenceData/a:NamingAndNumbering/a:Templates" -Namespace $XmlNamespace 

$regexList = @()
"Regexp;Naming_template_ID" |  Out-File .\Full_regex.csv



foreach ($namingTemplate_xml in $NamingTemplates.Node.ChildNodes){
    $final_regexp = ""
	
    $NT_Elements = Select-Xml -Xml $namingTemplate_xml -XPath "./a:Elements/a:Element" -Namespace $XmlNamespace
    $NT_Elements = $NT_Elements | Sort-Object {$_.Node.sortorder}

    $last_Naming_Element = $NT_Elements | Sort-Object {$_.Node.sortorder} | Select-Object -Last 1

    #$NT_Elements.Node |  Out-File .\sortorder.csv -Append
        foreach($NT_Element in $NT_Elements){
            $regexp = ""
            #Go to naming element and select validation source 
            $validation_element_XPath = "/a:ClassLibrary/a:ReferenceData/a:NamingAndNumbering/a:Elements/a:Element[@id='" + $NT_Element.Node.id + "']"
            $validation_element = Select-Xml -Xml $XmlDocument  -XPath $validation_element_XPath  -Namespace $XmlNamespace 
            $source_att_ID = $validation_element.Node.source
            
            #Go to attribute and select validation type, assuming all of them are enumerations
            $validation_attribute_XPath = "/a:ClassLibrary/a:Attributes/a:Attribute[@id='" + $source_att_ID + "']"
            $validation_attribute = Select-Xml -Xml $XmlDocument  -XPath $validation_attribute_XPath  -Namespace $XmlNamespace
                
            #Check parent Tag class, if ye then select its regular expression
            if($source_att_ID -eq "ALS-A0000002877"){
                $regexp =  $validation_element.Node.regEx
            }
            #Next check fo enumeration
            elseif($validation_attribute.Node.validationType -eq "Enumeration" ){
                $validationRule_id = $validation_attribute.Node.validationRule
                $enumeration_list_Xpath = "/a:ClassLibrary/a:ReferenceData/a:Enumerations/a:List[@id='" + $validationRule_id + "']/a:Items/a:Item"
                $validation_enum = Select-Xml -Xml $XmlDocument  -XPath $enumeration_list_Xpath  -Namespace $XmlNamespace
                $regexp = $validation_enum.Node.name -join "|"
               $regexp = "(" + $regexp + ")"
            }
            #and last in validation type regex (not enum) then use it straight away
            elseif($validation_attribute.Node.validationType -eq "RegularExpression"){
                $regexp = $validation_attribute.Node.validationRule
            }

            #Process dot suffixes
            $suffix = $NT_Element.Node.suffix
            if($NT_Elems.node.suffix -eq "."){$suffix = "\."}

		    $elem_suffix = $NT_Element.node.suffix
		    $elem_prefix = $NT_Element.node.prefix

            if([string]::IsNullOrEmpty($NT_Element.Node.mandatory)){
                $elem_mandatory = [System.Convert]::ToBoolean("True")
            }
            else{
                $elem_mandatory = [System.Convert]::ToBoolean($NT_Element.node.mandatory) 
            }
            if($elem_mandatory)
		    {
			    $regexp = $elem_prefix + $regexp + $elem_suffix 
		    }
		    else
		    {
			    $regexp = '(' + $elem_prefix + $regexp + $elem_suffix + ')?'
		    }

            if($NT_Element.Node.sortOrder -eq  $last_Naming_Element.node.sortorder){
                $regexp = $regexp + "$"
            }
            $asd = $validation_element.Node.name -replace '\W'
            $regexp = "(?<" + $asd + ">" + $regexp + ")"

            
            $final_regexp = $final_regexp + $regexp
        }
        $regexList += $final_regexp + ";" + $namingTemplate_xml.id
        
    }
  $regexList|  Out-File .\Full_regex.csv -Append

     
#Echo ($watch.Elapsed.TotalSeconds) #this at the end)

#$watch.reset()

# Define the DataTable Columns  
#$table = New-Object system.Data.DataTable 'TestDataTable'  
#$newcol = New-Object system.Data.DataColumn Regexp,([string]); $table.columns.add($newcol)  
#$newcol = New-Object system.Data.DataColumn NT_ID,([string]); $table.columns.add($newcol)  
#
#foreach($regex in $regexList){
#$row = $table.NewRow()  
#$row.Regexp= ($regex)  
#$row.NT_ID= ("EMPTY")  
#$table.Rows.Add($row)   
#}


#Add-Type -Path "\\QAMV3-SFIL102\Home\mch107\My Documents\14_PS\NOC_AIM\EPPlusFree.dll"
#$file = New-Object System.IO.FileInfo("\\QAMV3-SFIL102\Home\mch107\My Documents\14_PS\NOC_AIM\book.xlsx")
#
#$pkg = New-Object OfficeOpenXml.ExcelPackage($file)
#foreach($sheet in $pkg.Workbook.Worksheets){
#    if($sheet.Name -eq "MTR"){
#    $i = 1
#    foreach($regex in $regexList){
#        $sheet.Cells[$i,1].Value  = $regex
#        $i++
#    }
#        
#   }
#}
# #[void]$pkg.Workbook.Worksheets.Add("MTR").Cells.LoadFromArrays($regexList )
#$pkg.Save()

