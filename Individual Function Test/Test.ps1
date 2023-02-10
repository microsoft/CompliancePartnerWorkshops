
    $dlppolicyrule = Get-DlpCompliancerule  -Policy "DataRiskCheck-DLP-Teams"
    $sitlist = @()
    $labellist = @()
    
    foreach($rule in $dlppolicyrule){
        $rulename = $rule.DistinguishedName.split(",")[0] -replace 'CN='
        $ruledisabled

        if($rule.Disabled -like 'True'){
            $ruledisabled="No"
        }
        else{
            $ruledisabled="Yes"
        }
        ### new content to deal with advaned rules

        if($rule.IsAdvancedRule -like 'True'){
            #### do all the code for the advanced processing
            $rulejson = $rule.AdvancedRule | convertfrom-json | Select-Object -ExpandProperty Condition | Select-Object -ExpandProperty SubConditions

            foreach ($ruleitem in $rulejson){
                if ($ruleitem.ConditionName -like "ContentContainsSensitiveInformation"){
                    $sitgroups = $ruleitem | Select-Object -ExpandProperty Value | Select-Object -ExpandProperty Groups 

                    foreach($sitgroup in $sitgroups){
                    $sitruledetails = $sitgroup | Select-Object -ExpandProperty SensitiveTypes

                        foreach ($sitrule in $sitruledetails){
                            $shash = [ordered]@{
                            RuleGroup = $rulename
                            Name = $sitrule.name
                            RuleEnabled = $ruledisabled
                            ClassifierType = if($sitrule.ClassifierType){$sitrule.ClassifierType}else{"Content"}
                            MinCount = $sitrule.Minconfidence
                            MaxCount = $sitrule.Maxconfidence
                            ConfidenceLevel = $sitrule.Confidencelevel
                            }
                        $sobject = New-Object PSObject -Property $shash
                        $sitlist += $sobject
                        }
                    
                    }

                }

            }

        }

        elseif($null -ne $rule.ContentContainsSensitiveInformation){
        
            #check if more than one group of Sit's exsits in the policy
            if($null -ne $rule.ContentContainsSensitiveInformation.groups){
                $sitgroup = $rule.ContentContainsSensitiveInformation.groups
                
                foreach($group in $sitgroup){
                 $sensitivetypes = $group.sensitivetypes
                 $labels = $group.labels
                     foreach($sitg in $sensitivetypes){
                         $shash = [ordered]@{
                            RuleGroup = $rulename
                            Name = $sitg.name
                            RuleEnabled = $ruledisabled
                            ClassifierType = $sitg.ClassifierType
                            MinCount = $sitg.mincount
                            MaxCount = $sitg.maxcount
                            ConfidenceLevel = $sitg.confidencelevel
                         }
                         $sobject = New-Object PSObject -Property $shash
                         $sitlist += $sobject
                     }
                     foreach($label in $labels){
                        $lhash = [ordered]@{
                            RuleGroup = $rulename
                            LabelName = (get-label $label.name).DisplayName
                        }
                         $labellist += $lhash
                     }   
                }
             }
                 else{
                 $sits = $rule.ContentContainsSensitiveInformation
                    foreach($sit in $sits){
                        $shash = [ordered]@{
                            RuleGroup = $rulename
                            Name = $sit.name
                            RuleEnabled = $ruledisabled
                            ClassifierType = $sit.ClassifierType
                            MinCount = $sit.mincount
                            MaxCount = $sit.maxcount
                            ConfidenceLevel = $sit.confidencelevel
                         }
                         $sobject = New-Object PSObject -Property $shash
                         $sitlist += $sobject
                     }
                 }
        }      
    }
    
    $sitlist
    