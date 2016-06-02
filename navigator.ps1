Function Invoke-ISENavigator
{
<#
.Synopsis
ISENavigator provides a quick navigation menu which displays a list of functions, regions, and the main comment based help block.

.Description
ISENavigator is a useful tool for navigator large scripts quickly.  Particularly those that have a significant number of functions.
It was designed to be installed as an add-on in ISE and combined wtih a hot key provides a terrific ease of use feature.  This works against
the current open file within ISE only.

.Example
Invoke-ISENavigator

.Notes
Author : Jason Brisbin
Date : 6/2/2016
Version : 1
License : The MIT License

.link
https://github.com/jasonbrisbin/ISEIce
#>





#region Main
 if(-not($psise))
        {
            write-information -MessageData "Function not available outside of ISE." -InformationAction Continue
            return        
        }

        #The ISE Object model includes a script parser.....let's fire that bad boy up and make some magic.
        #We only care (for nav purposes) about a few types of scriptlines which will be called out in Token Filter
        #once the entire script has been parsed, let's filter it for just what we care about.
        $scriptbody=$psise.CurrentFile.Editor.Text
        $token_filter="Attribute","Keyword","Comment"
        $token_master=[System.Management.Automation.PSParser]::Tokenize($scriptbody,[ref]$null) | Sort-Object -Property startline,startcolumn
        $token_list= $token_master | where{($token_filter -match $_.type)}

        #There could be Many help blocks such as inside functions but we can discard those since we can jump to the top of a function
        #so let's display the first one only, we can use the following Flag to track that.
        $firsthelp=$false
        
        #Same thing here, let's just display the first cmdletbinding and we will track it with this flag
        $firstcmdletbinding=$false
        
        $jump_points=foreach($token in $token_list)
            {

                #Token_list includes ONLY lines that contain what we care about....time to party!
                switch($token)
                    {
                        #Is the Token a Function?
                        {($_.content -eq "function") -and ($_.type -eq "keyword")}
                            {
                                
                                #Alright so there can be some weirdness in the Tokenizer
                                #Once in a while the name of the function does not appear immediately after the function token and we need to search for it.
                                
                                #Get the current position in the Token Collection
                                $current_index=$token_master.indexof($token)
                                
                                #Keep checking the next item in the Token Collection until we find a CommandArgument (name of the function)
                                do
                                    {
                                
                                
                                        $next_token=$token_master[$current_index]
                                        $current_index++
                                    }until($next_token.type -eq "CommandArgument")
                        
                                #This will contain the function name
                                $name=$next_token.content
                        
                                #Return an object to the $jump_points collection with Content and Line Number
                                    $record=new-object psobject -Property @{
                                        "Content"="Function $name"
                                        "Line"=$token.startline
                                    }
                                $record
                                break
                            }
                
                        #Is the token the First Help block for the script?
                        {($_.type -eq "comment") -and ($firsthelp -eq $false) -and ($_.content -match ".Description")}
                            {
                                    $record=new-object psobject -Property @{
                                        "Content"="Comment Based Help"
                                        "Line"=$token.startline
                                    }
                                $record
                                $firsthelp=$true
                                break
                            }
                
                        #Is the token the First Help block for the script?
                        {($_.type -eq "comment") -and ($firsthelp -eq $false) -and ($_.content -match ".Synopsis")}
                            {
                                    $record=new-object psobject -Property @{
                                        "Content"="Comment Based Help"
                                        "Line"=$token.startline
                                    }
                                $record
                                $firsthelp=$true
                                break
                            }
                
                        #Is the token a Region block used for script organization?
                        {($_.type -eq "comment") -and ($_.content -like "#region*")}
                            {
                                    $content=(Get-Culture).TextInfo.ToTitleCase($token.content.replace("#",""))
                                    $record=new-object psobject -Property @{
                                        "Content"= $content
                                        "Line"=$token.startline
                                    }
                                $record
                                break
                            }
                
                        #Is the token the First Cmdletbinding for the script?
                        {($_.type -eq "attribute") -and ($firstcmdletbinding -eq $false) -and ($_.content -eq "CmdletBinding")}
                            {
                                    $record=new-object psobject -Property @{
                                        "Content"="Cmdletbinding"
                                        "Line"=$token.startline
                                    }
                                $record
                                $firstcmdletbinding=$true
                                break
                            }

                        #Default is to discard the result.
                        default
                            {
                                break
                            }
                    }
        
            }
        $jump_points=$jump_points|Sort-Object -Property line

        #Create the Windows Form to select where to Jump To in the current ISE window
        Add-Type -AssemblyName System.Windows.Forms
        $form = New-Object Windows.Forms.Form
        $form.Size = New-Object Drawing.Size @(300,200)        
        $form.StartPosition = "CenterScreen"
        $form.text = "Navigator"
        
        #Create the listbox control
        $listbox = New-Object System.Windows.Forms.listbox
        $listBox.Width = 300
        $listBox.Height = 200
        $form.AutoSize = $true
        
        #Binds the Listbox control to the Content and Line properties of objects when a new item is added
        $listbox.DisplayMember = "Content"
        $listbox.ValueMember = "Line"

        #Add new items to the listbox
        foreach($entry in $jump_points)
        {
            $listbox.items.add($entry) | out-null
        }

        #Add the listbox to the form
        $form.Controls.Add($listbox)

        #Create an event handler
        #This script block gets executed everytime something is selected in the Windows Form
        $listBox.add_SelectedIndexChanged(
            {
                #Get the object selected from the dialog box
                $result = $listBox.SelectedItem             
                
                #Find out the last line in the script
                #This is a visual trick only.  We will scroll to the end of the file first
                #Then scroll to the selected file second, this should ensure the selected line is near the top of the editor screen
                #To answer the obvious question: Yes it's kind of a hack.   
                $last_line=$psise.CurrentPowerShellTab.Files.SelectedFile.Editor.LineCount
                $psise.CurrentPowerShellTab.files.SelectedFile.Editor.EnsureVisible($last_line)
                $psise.CurrentPowerShellTab.files.SelectedFile.Editor.EnsureVisible($result.line)
                $psise.CurrentPowerShellTab.Files.SelectedFile.Editor.SetCaretPosition($result.line,1)
                $psIse.CurrentPowerShellTab.files.SelectedFile.Editor.SelectCaretLine()
                
                #Close the form immediately after the making the jump
                $form.close()
            })
    
        #Display the form
        $null = $form.ShowDialog()
#endregion
}
