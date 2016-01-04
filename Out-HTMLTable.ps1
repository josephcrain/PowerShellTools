function Out-HTMLTable
<#
.Synopsis
   Outputs an array of PSObjects to an HTML table
.EXAMPLE
    # Create some data
   $Object = 1..12 | % { [pscustomobject]@{'Month'=(get-date -Month $_ -Day 1 -Format 'MMMM');'Days'=([DateTime]::DaysInMonth((Get-Date).Year, $_))}}
   # Output to HTML Table
   $Object | Out-HTMLTable -Title "Days in each month of $((Get-Date).Year)"
.INPUTS
   An array of PS Objects to be output. Note that ordering is often not as expected, either use -Columns or pass an [ordered][pscustomobject]@{}
.OUTPUTS
   HTML String
.NOTES
   Contains code for an old version of this function that creates custom HTML tags in the data cells by passing additional columns to $ItemList formatted with HTML_COLNAME_ATTRIBUTENAME_ATTRIBUTEVALUE.  For this reason, any column name that starts with HTML_ will be ommited
#>
{
    [CmdletBinding()]
    Param
    (
    # Omitting title skips its output
    [string]$Title,
    [Parameter(ValueFromPipeline=$true)][psobject[]]$ItemList,
    [string[]]$Columns = '',
    [string]$EmptyMessage = "No records",
    [string]$TableStyleDefault = "font-size:8pt;font-family:Arial,sans-serif",
    [string]$TableStyle = "",
    [string]$TitleBG = "blue",
    [string]$TitleFG = "white",
    [string]$HeaderBG = "#000099",
    [string]$HeaderFG = "#FFFFFF",
    [string]$RowBG1 = "#FFFFFF",
    [string]$RowBG2 = "#E8E8E8",
    <#
    Array of row/cell formatting custom PS Object definitions, each with the following properties:

    row = '','row1', or 'row2' (use '' for all rows and row1/2 to apply to rows alternately)
    col = '','COLUMN_NAME' (use ''  for all rows or name the column  to apply formatting to)
    property = CSS property to modify
    value = value to apply to this property
    
    Examples:
    @{'row'='';'col'='';property='color';value='black'} # Change text of all data cells to black
    @{'row'='row1';'col'='';property='color';value='black'} # Change text of all row1 data cells to black
    @{'row'='row1';'col'='Month';property='color';value='black'} # Change text of all row1 data cells named 'Month' to black
    #>
    [pscustomobject[]]$CellFormatting
    )

    begin
    {
        # Set variable for row counting
        $RowCount = 0;

        # Flag for outputting header on first item
        $IsFirst = $true;
    }

    process
    {
        # Determine Column order and generate table header
        if ($IsFirst)
        {
            # Column Headers
            if ($ItemList.Count -gt 0)
            {
                # Get array of header names   
                if ($ItemList[0] -is [System.Data.DataRow])
                {
                    # DataTable Rows have properties named "Property" not "NoteProperty"
                    $Headers = @(($ItemList[0] | gm -MemberType Property).Name);
                }
                else
                {
                    $Headers = @(($ItemList[0] | gm -MemberType NoteProperty).Name);
                }

                # If a custom array of columns to use was passed, generate validated list to use in table creation
                $HeaderOrder = @();
                if ($Columns)
                {
                    # Loop each passed column
                    foreach($Column in $Columns)
                    {
                        # Validate column
                        if ($Headers -contains $Column)
                        {
                            $HeaderOrder += $Column;
                        }
                    }
                }
                else # Otherwise use all headers in order returned by the PSObject
                {
                    $HeaderOrder = @($Headers);
                }
    
                # Remove any HTML formatting properties (for backwards compatibility)
                $HeaderOrder = @($HeaderOrder -notlike 'HTML_*');
                $NumColumns = $HeaderOrder.Count
            }
            else
            {
                $NumColumns = 1;
            }

            # Table Title
            if ($TableStyle)
            {
                if ($TableStyleDefault) { $TableStyleDefault += ";$TableStyle"; }
                else { $TableStyleDefault = $TableStyle; }
            }
            $HTMLTable = "<table border=""1"" cellpadding=""4"" cellspacing=""0"" style=""$TableStyleDefault"">`n"
            if ($Title) { $HTMLTable += "<tr bgcolor=""$TitleBG""><td align=""center"" colspan=""$NumColumns""><strong><font color=""$TitleFG"">$Title</font></strong></td></tr>'`n" }
            $HTMLTable += "<tr bgcolor=""$HeaderBG"">`n";

            # Table Column Headers
            if ($ItemList.Count -gt 0)
            {
                foreach ($Column in $HeaderOrder)
                {
                    $HTMLTable += "<td align=""center""><strong><font color=""$HeaderFG"">$Column</font></strong></td>";
                }
                $HTMLTable += "</tr>";
            }

            $IsFirst = $false;
        }    

        if ($ItemList.Count -eq 0)
        {
            $HTMLTable += "<tr>
		            <td align=""center"">
		            <strong><font color=""#000099"">$EmptyMessage
		            </font></strong>
		            </td>
		            </tr>"
        }

        $ItemList | % `
        {

            # Add data records

            # Determine Even\Odd Row to alternate coloring
            if($RowCount%2)
            {
                $HTMLTable += "<tr bgcolor=""$RowBG2"">";
            }
            else
            {
                $HTMLTable += "<tr bgcolor=""$RowBG1"">";
            }

            foreach ($Column in $HeaderOrder)
            {
                # String for custom formatting
                $Format = '';
                $AlignFormat = "";

                # Loop all HTML formatting properties and apply formatting (for backwards compatibility)
                foreach ($p in @($Headers -like 'HTML_*'))
                {
                    # Strip HTML_ identifier
                    $s = $p -replace 'HTML_', '';

                    # Is this only supposed to apply to a specific column?
                    if ($s.IndexOf('_') -gt -1)
                    {
                        $HTMLAtt = $s.substring(0, $s.indexof("_"));
                        $HTMLAttCol = $s.substring($s.indexof("_") + 1);
                        # Does the specified column match the current column being output?  If so, apply custom HTML attribute
                        if ($Column -eq $HTMLAttCol) { $Format += " $HTMLAtt=""$($_."$p")"""; }
                    }
                    else
                    {
                        # This custom HTML attribute should apply to the entire row
                        $Format += " $s=""$($_."$p")""";
                    }
                }

                # Lookup custom formatting in $CellFormatting array
                $FormatMatches = $CellFormatting | ? { ($_.row -eq '' -or ($_.row -eq 'row1' -and !($RowCount%2)) -or ($_.row -eq 'row2' -and $RowCount%2)) -and ($_.col -eq '' -or $_.col -eq $Column) }
                $FormatStyle = "";
                foreach ($FormatMatch in $FormatMatches) { $FormatStyle += "$($FormatMatch.property):$($FormatMatch.value);" }

                # Center if this data is any type but a string
                if ([string]::IsNullOrEmpty($AlignFormat) -and $_."$Column" -isnot [string])
                {
                    $Format += " align=""center""";
                }
                if ($FormatStyle) { $Format += " style=""$FormatStyle"""; }
                $HTMLTable += "<td$Format>$($_."$Column")</td>`n";
            }
            $HTMLTable += "</tr>";

            # Increment row count
            $RowCount++;
        }
    }

    # Close table
    end
    {
        $HTMLTable += "</table>";
        $HTMLTable;
    }
}
