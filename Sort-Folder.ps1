<#
.SYNOPSIS
    Sorts files in a folder into subfolders based on the initial characters of their filenames.

.DESCRIPTION
    This script takes a folder path as input and organizes the files in that folder into subfolders.
    Subfolders are created based on the initial characters of the filenames, and files are moved
    into their respective subfolders. The script provides various options for customizing the sorting
    process, including combining and splitting folders based on file count and applying prefix/suffix
    to folder names.

.PARAMETER Path
    Path to the folder to sort (Default: Current file path).

.PARAMETER Force
    Force switch (If set, user will not be prompted to confirm).

.PARAMETER Empty
    Empty switch (If set, empty folders will also be created).

.PARAMETER Combine
    Combine switch (If set, folders containing fewer items nearby to each other can be combined).

.PARAMETER Split
    Split switch (If set, folders containing too many items can be separated).

.PARAMETER WhatIf
    WhatIf switch (If set, no changes will actually be made, and the expected output will be printed to the console).

.PARAMETER Recurse
    If this switch is set to true, the function will be called recursively on all folders created after they are created.
    This will occur 'depth' times, by default this will be set to two layers.

.PARAMETER Depth
    The depth for which the application should run recursively if the recursive switch is set to true.
    By default, this is set to two layers.

.PARAMETER CurrentDepth
    The depth for which the script is currently operating at. This is intended for use with recursive function calls
    of the script and should not be supplied manually.

.PARAMETER Upper
    If this switch is set to true, the folder names will be converted to uppercase before being created.
    Otherwise, they will be written as lowercase instead.

.PARAMETER IncludeCount
    If this switch is set to true, the number of files will be included in the name of the folder when created.

.PARAMETER Prefix
    Prefix for the folder, which will be added.

.PARAMETER Suffix
    Suffix for the folder, which will be added.

.PARAMETER Threshold
    Threshold for a folder to be broken up into multiple folders (i.e., the maximum size for any folders).
    This will be set to 10% of the sample size by default.

.NOTES
    File sorting script created by Damon Murdoch (@SirScrubbington).
    Date: 2023/10/09
    Version: 1.0

#>

Param(
  # Path to the folder to sort (Default: Current file path)
  [Alias('p')][Parameter(Mandatory = $False)][String]$Path = '.', 

  # Force switch (If set, user will not be prompted to confirm)
  [Alias('f')][Parameter(Mandatory = $False)][Switch]$Force, 

  # Empty switch (If set, empty folders will also be created)
  [Alias('e')][Parameter(Mandatory = $False)][Switch]$Empty, 

  # Combine switch (If set, folders containing fewer 
  # items nearby to each other can be combined (e.g. x-y))
  [Alias('c')][Parameter(Mandatory = $False)][Switch]$Combine,

  # Split switch (If set, folders containing too many items
  # can be seperated (e.g. a1, a2))
  [Alias('s')][Parameter(Mandatory = $False)][Switch]$Split,

  # WhatIf switch (If set, no changes will actually be made, 
  # and the expected output will be printed to the console.)
  [Alias('w')][Parameter(Mandatory = $False)][Switch]$WhatIf,

  # If this switch is set to true, the function wll be called
  # recursively on all folders created after they are created.
  # This will occur 'depth' times, by default this will be set
  # to two layers.
  [Alias('r')][Parameter(Mandatory = $False)][Switch]$Recurse,

  # The depth for which the application should run recursively, 
  # if the recursive switch is set to true. By default, this is
  # set to two layers.
  [Alias('d')][Parameter(Mandatory = $False)][Int]$Depth = 2,

  # The depth for which the script is currently operating at. This is
  # intended for use with recursive function calls of the script, and
  # should not be supplied manually.
  [Alias('cd')][Parameter(Mandatory = $False)][Int]$CurrentDepth = 0,

  # If this switch is set to true, the folder names will be
  # converted to upper case before being created. Otherwise, 
  # they will be written as lower case instead.
  [Alias('u')][Parameter(Mandatory = $False)][Switch]$Upper, 

  # If this switch is set to true, the number of files will 
  # be included in the name of the folder when created.
  [Alias('ic')][Parameter(Mandatory = $False)][Switch]$IncludeCount, 

  # Prefix for the folder which will be added
  [Alias('pre')][Parameter(Mandatory = $False)][String]$Prefix,

  # Suffix for the folder which will be added
  [Alias('suf')][Parameter(Mandatory = $False)][String]$Suffix,
  
  # Threshold for a folder to be broken up into multiple folders
  # (i.e. the maximum size for any folders). This will be set to
  # 10% of the sample size by default.
  [Alias('t')][Parameter(Mandatory = $False)][Int]$Threshold = 0
);

Function Show-Table {
  Param(
    # Display the hash table data
    [Alias()][Parameter(Mandatory = $True)][hashtable]$Hash
  );

  # Get the sorted keys from the table
  $Keys = $Hash.Keys | Sort-Object;

  # Loop over the keys
  ForEach ($Key in $Keys) {
      
    # Upper switch is set
    If ($Upper) {
      # Convert key to upper case
      $Key = $Key.ToUpper()
    }

    # Get the number of files
    $Count = $Hash[$Key].Count;

    # Include Count is set
    If ($IncludeCount) {
      $Key = "$Key [$Count]";
    }

    # Add prefix / suffix
    $Key = "$Prefix $Key $Suffix";

    Write-Output "Folder: $Key ($Count Files)";
  }
}

Try {
  Write-Output "Sorting folder '$Path' ...";

  # List of allowed keys
  $Keys = 0..9 + ('a'..'z')

  # Sort Table
  $Hash = @{};

  # Empty switch is set
  If ($Empty) {
    # Loop over the keys
    ForEach ($Key in $Keys) {
      # Add an empty array list to the hashtable (This is so jank lol)
      $Hash[([String]$Key)] = [System.Collections.ArrayList]@();
    }
  }

  # Get the child items in the folder
  $Children = Get-ChildItem -Path $Path;

  # Total file count
  $Count = 0;

  # Loop over all of the children
  ForEach ($Child in $Children) {

    # Get the full path for the file
    $FullName = $Child | Select-Object -ExpandProperty FullName;

    # Get the name for the file
    $Name = $Child | Select-Object -ExpandProperty Name;

    # Get the first character
    $StartChar = ([String](($Name.ToLower())[0]));

    # First character is NOT in the table
    If ($Hash.Keys -notcontains $StartChar) {

      # Add an empty array list to the hashtable
      $Hash[$StartChar] = [System.Collections.ArrayList]@();
    }

    # Add the name to the array
    $Hash[$StartChar] += $FullName;

    # Increment Counter
    $Count++;
  }

  Write-Output "Total Files Found: $Count ...";

  # Threshold is not set
  If ($Threshold -Eq 0) {
    # Set threshold to 10% of the count
    $Threshold = [Math]::Ceiling($Count / 10);

    Write-Output "Threshold set to 10% of count $($Count) [$($Threshold)] ...";
  }

  # Split switch is set
  If ($Split) {

    Write-Output "Split switch set: Splitting into folders of $Threshold or less ...";

    # Loop over the number of keys
    For ($i = 0; $i -lt ($Hash.Keys).Count; $i++) {

      # Get the keys from the table
      $Keys = [String[]]($Hash.Keys);

      # Get the key from the table
      $Key = $Keys[$i];

      # Dereference the array
      $Array = $Hash[$Key];

      # Get the hash array length
      $ArrayLength = $Array.Count;

      # If the array length exceeds the threshold
      If ($ArrayLength -Gt $Threshold) {
        # Split the array into two parts using slicing
        $StartHalf = $Array[0..($Threshold - 1)];
        $EndHalf = $Array[$Threshold..($ArrayLength - 1)];

        # Index Key Placeholders
        $StartKey = $EndKey = $Null;

        # If key is one character
        If ($Key.Length -eq 1) {
          $StartKey = "$($Key)1"; # e.g. a1
          $EndKey = "$($Key)2"; # e.g. a2
        }
        Else {
          # Key is multiple characters
          # Get the number from the key
          $Number = [Int]($Key[1]);

          # Increment the ending key number
          $EndKey = "$Key$($Number++)";
        }

        # If both the start key and end keys are not null
        If (($Null -ne $StartKey) -and ($Null -ne $EndKey)) {

          # Remove the original 
          # from the table
          $Hash.Remove($Key);

          # Add the starting half to the table
          $Hash[$StartKey] = $StartHalf;

          # Add the ending half to the table
          $Hash[$EndKey] = $EndHalf;

          # Return to previous index
          $i--;
        }
      }
    }
  }

  # Combine switch is set
  If ($Combine) {

    Write-Output "Combine switch set: Combining small adjacent into folders of $Threshold or less ...";

    # Loop over the number of keys (Minus one)
    For ($i = 0; $i -lt ($Hash.Keys).Count - 1; $i++) {

      # Get the keys from the table
      $Keys = $Hash.Keys | Sort-Object;

      # Get the key for both indexes
      $Key = $Keys[$i];
      $KeyNext = $Keys[$i + 1];

      # Get the data for both indexes
      $Table = $Hash[$Key];
      $TableNext = $Hash[$KeyNext];

      # If the combined size of both tables is within the threshold
      If (($Table.Count + $TableNext.Count) -lt $Threshold) {
        $KeyLeft = $Key.Split('-')[0]; # First key
        $KeyRight = $KeyNext.Split('-')[-1]; # Last key

        $KeyNew = "$($KeyLeft)-$($KeyRight)";

        # Delete the existing tables
        $Hash.Remove($Key);
        $Hash.Remove($KeyNext);

        # Combine the tables and add to the list
        $Hash[$KeyNew] = ($Table + $TableNext);

        # Start back at the beginning of the table
        $i = -1;
      }
    }
  }

  # At last two folders in the list
  If ($Hash.Keys.Length -gt 1) {

    Write-Output "Folder processing completed. Proposed Folder Structure:";

    # Show the table structure
    Show-Table -Hash $Hash;
  
    # WhatIf / Force switch are NOT set
    If (-Not ($WhatIf -Or $Force)) {
      Read-Host "Please press enter to confirm, or ctrl + c to quit.";
    }
  
    # Loop over the hash keys
    ForEach ($Key in $Hash.Keys) {
  
      # Create the folder name
      $FolderName = $Key;
  
      # Upper switch is set
      If ($Upper) {
        # Convert key to upper case
        $FolderName = $FolderName.ToUpper()
      }
  
      # Include Count is set
      If ($IncludeCount) {
        $FolderName = "$FolderName [$($Hash[$Key].Count)]";
      }
  
      # Join the path for the new folder
      $FolderPath = Join-Path -Path $Path -ChildPath $FolderName;
  
      Write-Output "Creating folder '$FolderPath' ...";
  
      If (-Not $WhatIf) {
        # Create a new directory for the files
        New-Item -Path $FolderPath -ItemType Directory -Force;
      }
      
      # Loop over all of the files
      ForEach ($FilePath in $Hash[$Key]) {
        Write-Output "Moving file '$FilePath' to '$FolderPath' ...";
  
        # WhatIf switch not set
        If (-Not $WhatIf) {
          Try {
            # Create a new directory for the files
            Move-Item -Path $FilePath -Destination $FolderPath;
          }
          Catch { # Failed to move item
            Write-Output "Failed to move file '$FilePath'! $($_.Exception.Message)";
          }
        }
      }
  
      Write-Output "All files moved to folder '$FolderPath' successfully.";
  
      # Recursive is true, and we have not exceeded the depth limit
      If ($Recurse -And ($CurrentDepth -Lt $Depth)) {
        ."$PSScriptRoot/Sort-Folder.ps1" -Path $FolderPath -Force:$True -Empty:$Empty -Combine:$Combine -Split:$Split -WhatIf:$WhatIf -Recurse:$Recurse -Depth $Depth -CurrentDepth ($CurrentDepth + 1) -Upper:$Upper -IncludeCount:$IncludeCount -Prefix $Prefix -Suffix $Suffix -Threshold $Threshold;
      }
  
      Write-Output "Folder processing completed.";
    }
  }
  Else { # 1-2 folders
    Write-Output "Output would result in only one child folder - No modifications made.";
  }
}
Catch {
  # Failed to sort folder
  Write-Output "Failed to sort folder! $($_.Exception.Message)";
}