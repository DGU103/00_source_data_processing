
<#
v6 Modeification list
1. Validation with functional code have been removed, now only class search will be performed
2. Added mapping to the disciplines in validation report
3. Added mapping to description in validation report


#>

param(
    [Parameter(Mandatory=$true)]
 #   [string]$Path = '.\2024-06-09_ASBH-LTM-MTGL-LST.xlsx',
   [string]$Path
    #[Parameter(Mandatory=$true,HelpMessage="UpdateRegexConfigurations (Yes/No)")]
    # [ValidateSet('Y','N','y','n')]
    # [string]$UpdateRegexConfigurations

)

$Path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\MTR\WeeklyReports\2025-04-13_EPCI 11-Master Tag Register-13-Apr-25.xlsx"

Clear-Host
Set-Location $PSScriptRoot
Import-Module .\importexcel.7.8.6\EPPlus.Interfaces.dll
Import-Module .\importexcel.7.8.6\ImportExcel.psd1
Import-Module .\importexcel.7.8.6\EPPlus.System.Drawing.dll

class Validation_Message{
            [String] $Error_code
            [String] $Tag_number
            [String] $Validation_Status
            [String] $Validation_message
            [String] $Service_message
            [String] $NNG_Tag_format
            [String] $Line_sequence_control
            [String] $Tag_class
            [String] $Tag_discipline
            [String] $Tag_description
            [String] $Status
            [String] $Action
            [String] $Location
            [String] $sub_system

}

Function Get-FunctionCodeClass{
    param(
        [Parameter(Mandatory=$true)]
        [string]$TagNumber
    )
 # -- Function code validation -- #
    $tag = $TagNumber
    $FUNCTIONAL_CODE = ""
    $FUNCTIONAL_CODE = $tag.Split('-')[1] 
    $ENUM_IDs = @()

    $xpath=""
    $xpath = "/a:ClassLibrary/a:ReferenceData/a:Enumerations/a:List[@aspect='true']/a:Items/a:Item[@name='" + $FUNCTIONAL_CODE + "']"
    $list_XML = Select-Xml -Xml $XmlDocument  -XPath $xpath -Namespace $XmlNamespace 
    $ENUM_IDs += $list_XML.Node.ParentNode.ParentNode.id #| Where-Object{$_ -match "VAL"}

    if([String]::IsNullOrEmpty($ENUM_IDs)){
        return "Functinal code """ +$FUNCTIONAL_CODE+ """ does not exist in list of approved codes"
    }

    foreach($ENUM_ID in $ENUM_IDs){
            
        # $Tag_validation_messages = @()
            
        # Reverse searching for a Naming Template base on FUNCTION CODE 
        #Attribute
        $xpath = "/a:ClassLibrary/a:Attributes/a:Attribute[@validationRule='" + $ENUM_ID + "']"
        $attribute_XML = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
        $attribute_ID = $attribute_XML.Node.id

        #Element
        $xpath = "/a:ClassLibrary/a:ReferenceData/a:NamingAndNumbering/a:Elements/a:Element[@source='" + $attribute_ID + "']"
        $ELEMENT_XML = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
        $ELEMENT_ID = $ELEMENT_XML.Node.id

        #Naming template
        $xpath = "/a:ClassLibrary/a:ReferenceData/a:NamingAndNumbering/a:Templates/a:Template/a:Elements/a:Element[@id='" + $ELEMENT_ID + "']"
        $NAMING_TEMPLATE_XML = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
        $NAMING_TEMPLATE = $NAMING_TEMPLATE_XML.Node.ParentNode.ParentNode
        # END of Reverse searching for a Naming Template base on FUNCTION CODE
        
        if([String]::IsNullOrEmpty($NAMING_TEMPLATE.id)){
            return "Cannot identify NAMING TEMPLATE base on FUNCTIONAL CODE in list of valid values for Tag: " + $tag
        }
        $xpath = "/a:ClassLibrary/a:Functionals/a:Class/a:NamingTemplates/a:NamingTemplate[@id='" + $NAMING_TEMPLATE.id + "']"
        $CLASS = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
        $temp_class = ""
        foreach($text in $CLASS.Node.ParentNode.ParentNode.Extension.ELEMENT_TYPE){$temp_class = $temp_class+"'"+ $text.replace('{geicl:notderivable}','') + "' or "}
        #$tag_discipline = $NAMING_TEMPLATE.Extension.Discipline.replace('{geicl:notderivable}','')

        return "Applicable equipmnet classes based on Function Code Lookup " + $temp_class,""
    }



    #$validation_report_messages += $Tag_validation_messages
}

Function Validate_Tag{
    param( 
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$Tag,
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [System.Xml.XmlNode]$Naming_Template

    )
    $tag_validation_string = $Tag
    # $Validation_Message = @()

    if([string]::IsNullOrEmpty($Pipeline_sequence_number)){$Pipeline_sequence_number = "NA"}


    $last_Naming_Element = $NAMING_TEMPLATE.Elements.ChildNodes | Sort-Object {$_.sortorder} | Select-Object -Last 1
    foreach($ELEMENT in $NAMING_TEMPLATE.Elements.ChildNodes | Sort-Object {$_.sortorder}){
        #$tag_is_valid = 1
        $regexp = ""
        $suffix = ""
        $prefix = ""
        $suffix = $ELEMENT.suffix
        $prefix = $ELEMENT.prefix

        if(-not[string]::IsNullOrEmpty($suffix)){$suffix = '\' + $suffix}
        if(-not[string]::IsNullOrEmpty($prefix)){$prefix = '\' + $prefix}

        $mandatory = $ELEMENT.mandatory
                    
        #Go to naming element and select validation source 
        $validation_element_XPath = "/a:ClassLibrary/a:ReferenceData/a:NamingAndNumbering/a:Elements/a:Element[@id='" + $ELEMENT.id + "']"
        $validation_element = Select-Xml -Xml $XmlDocument  -XPath $validation_element_XPath  -Namespace $XmlNamespace 
        $source_att_ID = $validation_element.Node.source
                    
        #Go to attribute and select validation type, assuming all of them are enumerations
        $validation_attribute_XPath = "/a:ClassLibrary/a:Attributes/a:Attribute[@id='" + $source_att_ID + "']"
        $validation_attribute = Select-Xml -Xml $XmlDocument  -XPath $validation_attribute_XPath  -Namespace $XmlNamespace

        #Check parent Tag class, if yes then select its regular expression
        if($source_att_ID -eq "ALS-A0000002877"){
            $regexp =  $validation_element.Node.regEx
        }
        # Check fo enumeration #
        elseif($validation_attribute.Node.validationType -eq "Enumeration" ){

            $validationRule_id = $validation_attribute.Node.validationRule
            $enumeration_list_Xpath = "/a:ClassLibrary/a:ReferenceData/a:Enumerations/a:List[@id='" + $validationRule_id + "']/a:Items/a:Item"
            $validation_enum = Select-Xml -Xml $XmlDocument  -XPath $enumeration_list_Xpath  -Namespace $XmlNamespace
        
            $enum_validation_flag = "false"
            $enum_regexp =  $validation_element.Node.regEx

            #$validation_value = $tag_validation_string - $enum_regexp
            $validation_value =  [Regex]::Match($tag_validation_string, $enum_regexp).value
            if($mandatory -ne "false"){
                 foreach($enum_value in $validation_enum.Node.name){
                    if($validation_value -eq $enum_value){
                        $enum_validation_flag = "true"
                        $regexp = $enum_value
                        break
                    }
                }       
            }
            else{
                if(-not([String]::IsNullOrEmpty($validation_value))){
                     foreach($enum_value in $validation_enum.Node.name){
                        if($enum_value -eq $validation_value){
                            $enum_validation_flag = "true"
                            $regexp = $enum_value
                            break
                        }
                    }             
                }
                else{
                    $enum_validation_flag = "true"
                }
            }

            if($enum_validation_flag -ne "true"){
                return @{"#0005"="For naming template: """+  $NAMING_TEMPLATE.name + """ with ID: """ + $NAMING_TEMPLATE.id + """ validation FAILED for Naming Element """ + $validation_element.Node.name + """ with pattern """ +  $enum_regexp + """, captured the value '"+$validation_value +" is not in list of valid values. Tag part '" + $tag_validation_string + "'" }
            }
        
        }

        # If validation type is Regular expression (not enumeration) then use this regular expression
        elseif($validation_attribute.Node.validationType -eq "RegularExpression"){
            $regexp = $validation_attribute.Node.validationRule
        }
    
        if($mandatory -ne "false")
	    {
		    $regexp = $prefix + $regexp + $suffix 
	    }
	    else
	    {
		    $regexp = '(' + $prefix + $regexp + $suffix + ')?'
	    }
                        
        if($ELEMENT.sortOrder -eq  $last_Naming_Element.sortorder){
            $regexp = $regexp + "$"
        }

        if($tag_validation_string -cmatch "^"+$regexp) {
            $tag_validation_string = $tag_validation_string -replace ("^"+$regexp) , ""
            # $tag_is_valid = 1
        }
        else{
            return @{"#0007"="For naming template: """+$NAMING_TEMPLATE.name+""" with ID: """+$NAMING_TEMPLATE.id + """ Validation FAILED for """ + $validation_element.Node.name + """ with expected pattern """ + $regexp + """ value captured '" + $tag_validation_string + "'"}
        }
    }

        $class_Xpath = "/a:ClassLibrary/a:Functionals/a:Class/a:NamingTemplates/a:NamingTemplate[@id='" + $NAMING_TEMPLATE.id + "']"
        $classes = Select-Xml -Xml $XmlDocument  -XPath $class_Xpath  -Namespace $XmlNamespace
        $classes_name = $classes.Node.ParentNode.ParentNode.Name -join "|"
        return @{"#0000"=$classes_name.split("|")[0] + " | Naming template: '" +$NAMING_TEMPLATE.id  + ";"}
}
Function Import-MTR-XLSX{
    param($Path, [int]$StartRow=1,[int]$EndRow=130000 )
# $final_data_table= [System.Data.DataTable]::new()
$data_table = [System.Data.DataTable]::new()

$datacol = [System.Data.DataColumn]::new()
$datacol.DataType = [string]
$datacol.ColumnName = "Tag_Number"
#$datacol.Unique = $true
$data_table.Columns.Add($datacol)

$datacol = [System.Data.DataColumn]::new()
$datacol.DataType = [string]
$datacol.ColumnName = "Tag_class"
$data_table.Columns.Add($datacol)

$datacol = [System.Data.DataColumn]::new()
$datacol.DataType = [string]
$datacol.ColumnName = "Tag_status"
$data_table.Columns.Add($datacol)

$datacol = [System.Data.DataColumn]::new()
$datacol.DataType = [string]
$datacol.ColumnName = "Tag_description"
$data_table.Columns.Add($datacol)

$datacol = [System.Data.DataColumn]::new()
$datacol.DataType = [string]
$datacol.ColumnName = "Tag_subsystem"
$data_table.Columns.Add($datacol)

$Path = (Resolve-Path $Path).ProviderPath
$Stream = New-Object -TypeName System.IO.FileStream -ArgumentList $path, 'Open', 'Read', 'ReadWrite'
$Excel = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Stream

foreach($Worksheet in $Excel.Workbook.Worksheets){
        
    $Tag_col = $Worksheet.Cells["A1:AW5"] | Where-Object{$_.text -eq "Tag Number"}
    $Tag_class = $Worksheet.Cells["A1:AW5"] | Where-Object{$_.text -eq "Type of Item"}
    $Tag_status = $Worksheet.Cells["A1:AW5"] | Where-Object{$_.text -eq "Logically Deleted"}
    $Tag_description = $Worksheet.Cells["A1:AW5"] | Where-Object{$_.text -eq "DESCRIPTION"}
    $Tag_subsystem = $Worksheet.Cells["A1:AW5"] | Where-Object{$_.text -eq "SUBSYSTEM"}

    if($Tag_col.Length -eq 0){throw "Tag number column not found"}
    if($Tag_class.Length -eq 0){throw "Type of Item column not found"}
    if($Tag_status.Length -eq 0){throw "Logically Deleted column not found"}
       
    $first_data_row = [int][Regex]::Match($Tag_col.Address, "[0-9]{1,2}").value + 1
    if($StartRow -gt $first_data_row ){$first_data_row = $StartRow}
    $Tag_col = [Regex]::Match($Tag_col.Address, "[A-Z]{1,2}").value
    $Tag_class  = [Regex]::Match($Tag_class.Address, "[A-Z]{1,2}").value
    $Tag_status  = [Regex]::Match($Tag_status.Address, "[A-Z]{1,2}").value
    $Tag_description  = [Regex]::Match($Tag_description.Address, "[A-Z]{1,2}").value
    $Tag_subsystem  = [Regex]::Match($Tag_subsystem.Address, "[A-Z]{1,2}").value
    
    for($i=$first_data_row; $i -le $EndRow; $i++){

        $data_row = $data_table.NewRow()

        $data_range = $Tag_col + $i
        $cell_value = $Worksheet.cells[$data_range].Text
        $data_row["Tag_Number"] = $cell_value
            
        if($cell_value -eq ""){break}

        $data_range = $Tag_class + $i
        $cell_value = $Worksheet.cells[$data_range].Text
        $data_row["Tag_class"] = $cell_value
            
        $data_range = $Tag_status + $i
        $cell_value = $Worksheet.cells[$data_range].Text
        $data_row["Tag_status"] = $cell_value
        
        $data_range = $Tag_description + $i
        $cell_value = $Worksheet.cells[$data_range].Text
        $data_row["Tag_description"] = $cell_value
            
        $data_range = $Tag_subsystem + $i
        $cell_value = $Worksheet.cells[$data_range].Text
        $data_row["Tag_subsystem"] = $cell_value
            
        $data_table.Rows.Add($data_row)
    }
    break
}

$Stream.Close()
$Stream.Dispose()
$Excel.Dispose()
$Excel = $null
#$final_data_table = $data_table |  Where-Object{$_.Tag_status -eq "False"}

return $data_table |  Where-Object{$_.Tag_status -ne "True"}
}
Function Import-MTR-CSV{

    param(
        [Parameter(Mandatory=$true)]
        $Path 
    )
    $data_table = [System.Data.DataTable]::new()

    $datacol = [System.Data.DataColumn]::new()
    $datacol.DataType = [string]
    $datacol.ColumnName = "Tag_Number"
    #$datacol.Unique = $true
    $data_table.Columns.Add($datacol)

    $datacol = [System.Data.DataColumn]::new()
    $datacol.DataType = [string]
    $datacol.ColumnName = "Tag_class"
    $data_table.Columns.Add($datacol)

    $datacol = [System.Data.DataColumn]::new()
    $datacol.DataType = [string]
    $datacol.ColumnName = "Tag_status"
    $data_table.Columns.Add($datacol)

    $Path = (Resolve-Path $Path).ProviderPath

    $csv = Import-Csv -Delimiter ";" -Path $Path 
    foreach($record in $csv){
        $data_row = $data_table.NewRow()
        $data_row["Tag_Number"] = $record.'TAG NUMBER'
        $data_table.Rows.Add($data_row)
    }
    return $data_table
}

$ism_file = Get-ChildItem -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\RegisterGateway\source data\class library\" -Filter clib.xml | Sort-Object | Select-Object -Last 1


if($UpdateRegexConfigurations -eq 'Y' -or $UpdateRegexConfigurations -eq 'y'){
    
    .\build_full_regex_config.ps1 -ism_file $ism_file.FullName


    .\build_light_regex_config.ps1 -ism_file $ism_file.FullName
}

if($Path -match "xlsx$"){
 $Tag_data_table = Import-MTR-XLSX -Path $Path -StartRow 1 -EndRow 150000
   
}
elseif($Path -match "csv$"){
 $Tag_data_table = Import-MTR-CSV -Path $Path
 $Path = $Path -replace ".csv" , "_validation_report.csv"
}

$config_directory = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\06_Regexp_configs"

$light_regex_csv = Import-Csv -Path ($config_directory + "\Light_regex.csv") -Delimiter ";"
# $full_regex_csv = Import-Csv -Path ($config_directory + "\Full_regex.csv") -Delimiter ";"
# $all_light_regexps = $light_regex_csv | Select-Object {$_.Regexp} -Unique




Write-Output ""
Write-Output ""
Write-Output ""

[xml]$XmlDocument = Get-Content $ism_file.FullName
$XmlNamespace = @{ a = "http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01" 
 nmcltr ="http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Nomenclature/2015/09"}
    
 Write-Host ("[Class library Path]	") -ForegroundColor Gray -NoNewline
 Write-Host ($ism_file.FullName) -ForegroundColor Cyan
 Write-Host ("[Class library ID]		")  -ForegroundColor Gray -NoNewline
 Write-Host ($XmlDocument.ClassLibrary.id)  -ForegroundColor Cyan
 Write-Host ("[Class library Version]	") -ForegroundColor Gray -NoNewline
 Write-Host ($XmlDocument.ClassLibrary.version)  -ForegroundColor Cyan
 


if($Tag_data_table[5].Tag_Number.split('-')[0] -match 'BH|BE|BD'){
    $package_code =  "EPCIC12"
}
elseif($Tag_data_table[5].Tag_Number.split('-')[0] -match 'BJ'){
    $package_code =  "EPCIC13"
}
else{
    $package_code =  "EPCIC11"
}
$Host.UI.RawUI.WindowTitle = $Path

$Tag_table = @{}
$Line_table = @()
# $ii = 0

#$validation_report_size = $count * 10
#$validation_report = @(100000)

$validation_report = New-Object Validation_Message[] 300000
$vri = 0


$count = $Tag_data_table.Count
$i = 0
<# Duplicate Tag search #>
foreach($rowee in $Tag_data_table){
    $z = [Math]::Round($i+ 100 / $count,2)
    Write-Progress -Activity "Duplicated Tag search in progress... " -Status "$z% Complete: from $count" -PercentComplete $i
    $i = $i+ 100 / $count

    if(-not $Tag_table[$rowee.Tag_Number] ){
        $Tag_table.Add($rowee.Tag_Number,$rowee)
    }
    if ($rowee.Tag_Number.split('-').count -ge 5)
    {
        $Line_table+=$rowee.Tag_Number.split('-')[0] + '-' + $rowee.Tag_Number.split('-')[3]
    }
}

$duplicates = $Tag_data_table | Group-Object -Property Tag_Number | Where-Object { $_.count -ge 2 } 

$Line_table_duplicates = $Line_table | Group-Object | Where-Object { $_.count -ge 2 }


$watch = New-Object System.Diagnostics.Stopwatch
$watch.Start()

Write-Output "Main validation"

$count = $Tag_data_table.Count
$i = 0
foreach($row in $Tag_data_table){
    $z = [Math]::Round($i+ 100 / $count,2)
     Write-Progress -Activity "Processing in Progress" -Status "$z% Complete: from $count" -PercentComplete $i
     $i = $i+ 100 / $count
    
    $tag = $row.Tag_Number
    $tag_class = $row.Tag_class
    $tag_description = "Empty"
    $tag_description = $row.Tag_description -replace "\n", " "
    $tag_subsystem = $row.Tag_subsystem
    $Tag_light_validation = "Not Valid"
    # $Tag_Light_NTs = @()

   
    if($duplicates | Where-Object {$_.name -match $tag} ){
        $vm = New-Object -TypeName Validation_Message
        $vm.Tag_number = $tag 
        $vm.Validation_Status = "Not Valid"
        $vm.Error_code = "#0012"
        $vm.Validation_message = "Duplicated Tag"
        $vm.Service_message = "Duplicated Tag"
        $vm.NNG_Tag_format = "NA"
        $vm.Tag_discipline = "Not identified"
        # $vm.Location = $location
        
        
        $validation_report[$vri] = $vm
        $vri++
        CONTINUE
    }


    # First iteration is to check any Light Pattern validation
    foreach($light_NT in $light_regex_csv){
        if($tag -match $light_NT.Regexp){
            $Tag_light_validation = "Valid"
            #$Tag_Light_NTs += $light_NT.Naming_template_ID
            BREAK

        }
    }
    if( -not($Tag_light_validation)){
        
        $vm = New-Object -TypeName Validation_Message
        $vm.Tag_number = $tag 
        $vm.Tag_class = $tag_class
        $vm.Tag_description = $tag_description
        $vm.Validation_Status = "Not Valid"
        $vm.Error_code = "#0002"
        $vm.Validation_message  = "Tag does not match any pattern"
        $vm.Service_message = "Light regex validation"
        $vm.NNG_Tag_format = "NA"
        $vm.Tag_discipline = "Not identified"
        # $vm.Location = $location
        $validation_report[$vri] = $vm
        $vri++
        continue
    }
    <# If Tag did not pass Light regexp validation then location may not be defined #>
    if ($row.Tag_Number.split('-')[0].Length -ge 4) {
        $location = $row.Tag_Number.split('-')[0].substring(0,4) | Where-Object {$_ -match '^AS[A-Z]{2}'}
    }
    else {
        $location = 'Cannot be extracted from Tag ID'
    }
    # Then check is this a signal Tag with dot
    IF($tag -match '\.[0-9]{1,2}[A-Z]?$'){
        $PARENT_TAG = $tag.Split('.')[0]
        $ANOTHER_PARENT_TAG = [regex]::Match($tag, "[A-Z]{5}-N-[0-9]{4}")
        $ANOTHER_PARENT_TAG = $ANOTHER_PARENT_TAG.value

        if(-not ($Tag_table[$PARENT_TAG] -or $Tag_table[$ANOTHER_PARENT_TAG])) {
                $vm = New-Object -TypeName Validation_Message
                $vm.Tag_number = $tag 
                $vm.Tag_class = $tag_class
                $vm.Tag_description = $tag_description
                
                $vm.Validation_Status = "Not Valid"
                $vm.Error_code = "#0001"
                $vm.Validation_message ="Parent Tag "+ $PARENT_TAG.Text + $ANOTHER_PARENT_TAG.text +" not found in current register"
                $vm.Service_message = "Parent Tag"
                $vm.NNG_Tag_format = "Instrumentation"
                $vm.Tag_discipline = "Instrumentation"
                $vm.Location = $location
                $validation_report[$vri] = $vm
                $vri++
            CONTINUE
        }
           
    }

    # Tag CLASS from MTR based validation #
    if([string]::IsNullOrEmpty($tag_class)){
                                $vm = New-Object -TypeName Validation_Message
                                $vm.Tag_number = $tag 
                                $vm.Tag_class = $tag_class
                                $vm.Tag_description = $tag_description

                                $vm.Validation_Status = "Not Valid"
                                $vm.Error_code = "#0011"
                                $vm.Validation_message = "Tag class have not been provided"
                                $vm.Service_message = "Mapping procedure"
                                $vm.NNG_Tag_format = "NA"
                                $vm.Tag_discipline = "Not identified"
                                $vm.Location = $location
                                $validation_report[$vri] = $vm
                                $vri++
        CONTINUE
    }

    #$xpath = "/a:ClassLibrary/a:Functionals/a:Class[@aspect='"+$tag_class +"']"
    $xpath = "/a:ClassLibrary/a:Physicals/a:Class/a:Extension[@nmcltr:AENG='"+$tag_class +"']"
    $TAG_CLASS_XML = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
    
    ## -- Attempt to map class with functional class -- ##
    if(-not($TAG_CLASS_XML.Node)){
        $xpath = "/a:ClassLibrary/a:Functionals/a:Class/a:Extension[@nmcltr:AENG='"+$tag_class +"']"
        $TAG_CLASS_XML = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
    }

    if(-not($TAG_CLASS_XML.Node)){
                                $vm = New-Object -TypeName Validation_Message
                                $vm.Tag_number = $tag 
                                $vm.Tag_class = $tag_class
                                $vm.Tag_description = $tag_description

                                $vm.Validation_Status = "Not Valid"
                                $vm.Error_code = "#0009"
                                $vm.Validation_message = "Class "+ $tag_class +" cannot be mapped"
                                $vm.Service_message = "Mapping procedure"
                                $vm.NNG_Tag_format = "NA"
                                $vm.Tag_discipline = "Not identified"
                                $vm.Location = $location
                                $validation_report[$vri] = $vm
                                $vri++
        CONTINUE
    }
    if($TAG_CLASS_XML.Node.ParentNode.NamingTemplates.ChildNodes.Count -eq 0){
                                $vm = New-Object -TypeName Validation_Message
                                $vm.Tag_number = $tag 
                                $vm.Tag_class = $tag_class
                                $vm.Tag_description = $tag_description

                                $vm.Validation_Status = "Not Valid"
                                $vm.Error_code = "#0010"
                                $vm.Validation_message = "There is no Tag Naming Template defined for class " + $tag_class
                                $vm.Service_message = "Mapping procedure"
                                $vm.NNG_Tag_format = "NA"
                                $vm.Tag_discipline = "Not identified"
                                $vm.Location = $location
                                $validation_report[$vri] = $vm
                                $vri++
        CONTINUE
    }

    
    <#Class based validation#>   
    $Tag_validation_messages = @()
    # $tvm_i = 0
    foreach($CLASS_NAMING_TEMPLATE in $TAG_CLASS_XML.Node.ParentNode.NamingTemplates.ChildNodes){
        $xpath = "/a:ClassLibrary/a:ReferenceData/a:NamingAndNumbering/a:Templates/a:Template[@id='"+ $CLASS_NAMING_TEMPLATE.id +"']"
        $NAMING_TEMPLATE_CLASS_XML = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
        

        $vm = New-Object -TypeName Validation_Message
        $vm.Tag_number = $tag 
        $vm.Tag_class = $tag_class
        $vm.Tag_description = $tag_description
        
        $vm.Service_message = "Class based validation"

        $tag_discipline = $NAMING_TEMPLATE_CLASS_XML.Node.Extension.Discipline.Replace('{geicl:notderivable}','')
        $vm.Tag_discipline = $tag_discipline

        $NNG_mapping = $NAMING_TEMPLATE_CLASS_XML.Node.Extension.NNG_TAG_FormatID.Replace('{geicl:notderivable}','')
        $vm.NNG_Tag_format = $NNG_mapping

        $message = Validate_tag -Tag $tag -Naming_Template $NAMING_TEMPLATE_CLASS_XML.Node 
        $vm.Error_code =     $message.keys[0]
        $vm.Validation_message = $message.Values[0] 
        $vm.Validation_Status = "Not Valid"
        $vm.Location = $location


        if($message['#0000'] -and $validation_report){
            $vm.Validation_message = $message['#0000']
            $vm.Error_code = '#0000'
            $vm.Validation_Status = "Valid"
            $vm.sub_system = $tag_subsystem
            $Tag_validation_messages = @()
            $Tag_validation_messages += $vm
            break
        }
        else{
            #$temp_message = Get-FunctionCodeClass -TagNumber $tag
            $Tag_validation_messages += $vm
            #$tvm_i++
        }

    }

    foreach($Tag_validation_message in $Tag_validation_messages){
            $validation_report[$vri] = $Tag_validation_message;
            $vri++
    }    
<#  
    END ofClass based validation 
    All next checks will be applied
#>   

    IF($tag -match "[A-Z]{5}-(FCV|ICD|LCV|PCV|PSV|SOV|TCV)-[0-9]{6}[A-Z]"){
        $targetTag = [Regex]::Match($tag, "[A-Z]{5}-(FCV|ICD|LCV|PCV|PSV|SOV|TCV)-[0-9]{6}").value
        if ($Tag_table[$targetTag])
        {

            $vm = New-Object -TypeName Validation_Message
            $vm.Tag_number = $tag 
            $vm.Tag_class = $tag_class
            $vm.Tag_description = $tag_description

            $vm.Validation_Status = "Not Valid"
            $vm.Error_code = "#0013"
            $vm.Validation_message = "There are $tag and $targetTag identified in the register. The sub-sequence index (A, B, C etc.) allowed only if a control loop contains two or more instruments with the same Functional Code"
            $vm.Service_message = "Advanced logic validation"
            $vm.NNG_Tag_format = $validation_report[$vri-1].NNG_Tag_format
            $vm.Tag_discipline = $validation_report[$vri-1].Tag_discipline

            # In case if Tag validated succesfully based on previous rules, 
            # we need to replace that message with an error,
            # otherwise add an additional one.
            
            #Write-Host "Error $targetTag cannot find pair $tag" -ForegroundColor 12 

        }
    }
    <# Pipleine sequence validation #>
    
    if(<#$tag.split('-').count -ge 5) #> $tag_class -eq ":PipelinePhysical" -and ($tag -notmatch "^PL-.{1,}")){
        $pipeline_seq = $tag.Split('-')[0] + "-" + $tag.Split('-')[3]  
        if ($Line_table_duplicates | Where-Object{$_.Name -eq $pipeline_seq})
        {
            $vm = New-Object -TypeName Validation_Message
            $vm.Tag_number = $tag 
            $vm.Tag_class = $tag_class
            $vm.Tag_description = $tag_description
            $vm.Validation_Status = "Not Valid"
            $vm.Error_code = "#0014"
            $vm.Validation_message = "Duplicated PipeLine sequence number found for: " + $pipeline_seq.Replace('-','*')
            $vm.Service_message = "Duplicated PipeLine sequence number"
            $vm.NNG_Tag_format = $validation_report[$vri-1].NNG_Tag_format
            $vm.Tag_discipline = $validation_report[$vri-1].Tag_discipline
            $vm.Location = $location
        
            if ($validation_report[$vri-1].Validation_Status -eq "Valid")
            {
                $validation_report[$vri-1] = $vm
            }
            else
            {
                $validation_report[$vri] = $vm
                $vri++
            }
        }
    }
    
                
} 


<# End Tag validation #>
 $watch.Stop()
 $message =  "Validation is: " +[Math]::Round($watch.Elapsed.TotalSeconds / $count, 3) + " seconds per Tag"
 Write-Output $message

# $WarningAction = "SilentlyContinue"


# --- Basic report Export --- #
$csv_path = $Path.Replace('.xlsx', '_validation_report_v6.6.csv')
$validation_report | Where-Object {$_.Tag_number} | Export-Csv -Path $csv_path -NoTypeInformation -force


# --- AIM-A Export --- #
# $gdp_staging_area = "\\QAMV3-SFIL102\Home\mch107\My Documents\14_PS\00_Tag_control\"
$gdp_staging_area = "\\Qamv3-sapp243\gdp\GDP_StagingArea\MP\MTR\"

if (Test-Path -Path $gdp_staging_area) {
   $aim_export_path = [System.IO.Path]::Combine($gdp_staging_area, 'MTR_' + $package_code+ "_Tag_Load_to_AIM-A.csv")
   $validation_report | Where-Object {$_.Validation_Status -eq "Valid"} | Select-Object ("Tag_number", "Tag_class", "Tag_description", "Action", "Status", @{Name = "Discipline"; Expression = {$_.Tag_discipline}}, @{Name = "Platform"; Expression = {$_.Location}}, @{Name = 'Sub System'; Expression = {$_.'Sub_System'}} ) | Export-Csv -Path $aim_export_path -NoTypeInformation -Force
}

# --- Validation report in XLSX format Export --- #
$csv = Import-Csv $csv_path -Delimiter ","
$xlsx_path = $csv_path.Replace(".csv",".xlsx")
if (Test-Path $xlsx_path) {Remove-Item $xlsx_path -WarningAction silentlyContinue}
$csv | Export-Excel $xlsx_path -WorksheetName "Sheet1" -FreezeTopRow -AutoFilter -AutoNameRange -BoldTopRow -NoNumberConversion *
$csv | Select-Object ("Tag_number", "Validation_Status") -Unique | Export-Excel $xlsx_path -WorksheetName "Sheet2" -FreezeTopRow -AutoFilter -AutoNameRange -BoldTopRow -NoNumberConversion *



# --- Extract for consistency report Export --- #
$consistency_report_path = [System.IO.Path]::Combine("\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\MTR", $package_code+ "_validation_report.xlsx")
if (Test-Path $consistency_report_path) {Remove-Item $consistency_report_path -WarningAction silentlyContinue}
$csv | Export-Excel $consistency_report_path -WorksheetName "Sheet1"  

# --- NNG file format Export --- #
#$xlsx_path = $csv_path.Replace(".csv","_for_NNG.xlsx")
#$csv | select ("Tag_number", "Validation_Status") -Unique |Where-Object{$_.'Tag format' -ne "NA"} | Export-Excel $xlsx_path -AutoSize
