#----------------------------------------------------------------
#Parameters 
#----------------------------------------------------------------
param($server = "", #ServerName parameter to connect 
      $user = "", #UserName parameter  to connect
      $password = "", #Password Parameter  to connect
      $Db = "", #DBName Parameter  to connect
      $Provider = "") #Provider

#----------------------------------------------------------------
#Connection SQLClient
#----------------------------------------------------------------

Function GiveMeConnectionSourceSQLClient()
{ 
  for ($i=1; $i -lt 10; $i++)
  {
   try
    {
      logMsg( "Connecting to the database...Attempt #" + $i) (1)
      $SQLConnection = New-Object System.Data.SqlClient.SqlConnection 
      $SQLConnection.ConnectionString = "Server="+$server+";Database="+$Db+";User ID="+$user+";Password="+$password+";Connection Timeout=60" 
      $SQLConnection.Open()
      logMsg("Connected to the database...") (1)
      return $SQLConnection
      break;
    }
  catch
   {
    logMsg("Not able to connect - Retrying the connection..." + $Error[0].Exception) (2)
    Start-Sleep -s 5
   }
  }
}

#----------------------------------------------------------------
#Connection OleDB
#----------------------------------------------------------------

Function GiveMeConnectionSourceOleDB()
{ 
  for ($i=1; $i -le 2; $i++)
  {
   try
    {
      logMsg( "Connecting to the database...Attempt #" + $i) (1)
      $InfoProvider = "SQLOLEDB"
        if($Provider -eq "2" )
        { 
          $InfoProvider = "MSOLEDBSQL"
        }
      $SQLConnection = New-Object System.Data.OleDb.OleDbConnection
      $SQLConnection.ConnectionString = "Provider="+$InfoProvider+";Server="+$server+";Database="+$Db+";UID="+$user+";PWD="+$password+";Timeout=60" 
      $SQLConnection.Open()
      logMsg("Connected to the database...") (1) 
      return $SQLConnection
      break;
    }
  catch
   {
    logMsg("Not able to connect - Retrying the connection..." + $Error[0].Exception) (2)
    Start-Sleep -s 5
   }
  }
}

#----------------------------------------------------------------
#Connection ADODB
#----------------------------------------------------------------

Function GiveMeConnectionSourceADODB()
{ 
  for ($i=1; $i -le 2; $i++)
  {
   try
    {
      logMsg( "Connecting to the database...Attempt #" + $i) (1)
      $InfoProvider = "SQLOLEDB"
        if($Provider -eq "5" )
        { 
          $InfoProvider = "MSOLEDBSQL"
        }
      $SQLConnection = New-Object -comobject ADODB.Connection
      $SQLConnection.ConnectionString = "Provider="+$InfoProvider+";Server="+$server+";Database="+$Db+";UID="+$user+";PWD="+$password+";Timeout=60" 
      $SQLConnection.Open()
      logMsg("Connected to the database...") (1) 
      return $SQLConnection
      break;
    }
  catch
   {
    logMsg("Not able to connect - Retrying the connection..." + $Error[0].Exception) (2)
    Start-Sleep -s 5
   }
  }
}

#--------------------------------
#Log the operations
#--------------------------------
function logMsg
{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $msg,
         [Parameter(Mandatory=$false, Position=1)]
         [int] $Color
    )
  try
   {
    $Fecha = Get-Date -format "yyyy-MM-dd HH:mm:ss"
    $msg = $Fecha + " " + $msg
    $Colores="White"
    If($Color -eq 1 )
     {
      $Colores ="Cyan"
     }
    If($Color -eq 3 )
     {
      $Colores ="Yellow"
     }

     if($Color -eq 2)
      {
        Write-Host -ForegroundColor White -BackgroundColor Red $msg 
      } 
     else 
      {
        Write-Host -ForegroundColor $Colores $msg 
      } 
   }
  catch
  {
    Write-Host $msg 
  }
}

#--------------------------------
#Validate Param
#--------------------------------

function TestEmpty($s)
{
if ([string]::IsNullOrWhitespace($s))
  {
    return $true;
  }
else
  {
    return $false;
  }
}

##Let's run the process

cls

if (TestEmpty($server)) { $server = read-host -Prompt "Please enter a Server Name" }
if (TestEmpty($user))  { $user = read-host -Prompt "Please enter a User Name"   }
    if (TestEmpty($password))  
    {  
       $passwordSecure = read-host -Prompt "Please enter a password"  -assecurestring  
       $password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($passwordSecure))
    }
if (TestEmpty($Db))  { $Db = read-host -Prompt "Please enter a Database Name"  }
if (TestEmpty($Provider))  { $Provider = read-host -Prompt "Please enter the provider to use - (1) - OLEDB SQLOLEDB or (2) - OLEDB MSOLEDBSQL or (3) .Net SQL Client or (4) ADO - SQLOLEDB or (5) ADO - MSOLEDBSQL"  }

$QueryText = read-host -Prompt "Please enter the query to execute"

if (TestEmpty($QueryText)) 
{
  logMsg("Query Text didn't provide." ) (2)
  exit;
}

$iProvider = $Provider -as [int] 

if($iProvider -le 0 -Or $iProvider -ge 6  ) 
{
  logMsg("Incorrect provider specified." ) (2)
  exit;
}

if(TestEmpty($Provider))
{
  logMsg("Incorrect provider specified." ) (2)
  exit;
}
  
try 
{
  logMsg("-------------------------") (1)
  $sw = [diagnostics.stopwatch]::StartNew()     

  if( $Provider -eq "1" -or $Provider -eq "2" )
  {
    $connection = GiveMeConnectionSourceOleDB
    if (TestEmpty($connection)) 
    {
     logMsg("Retry-Logic was not working - It is not possible to continue with the program." ) (2)
     exit;
    }

    $command = New-Object -TypeName System.Data.OleDb.OleDbCommand
    $command.Connection=$connection
  }
  if( $Provider -eq "3"  )
  {
    $connection = GiveMeConnectionSourceSQLClient
    if (TestEmpty($connection)) 
    {
     logMsg("Retry-Logic was not working - It is not possible to continue with the program." ) (2)
     exit;
    }

    $command = New-Object System.Data.SqlClient.SqlCommand
    $command.Connection=$connection
  }

  if( $Provider -eq "4" -or $Provider -eq "5" )
  {
    $connection = GiveMeConnectionSourceADODB
    if (TestEmpty($connection)) 
    {
     logMsg("Retry-Logic was not working - It is not possible to continue with the program." ) (2)
     exit;
    }

    $command = New-Object -comobject ADODB.Command
    $command.ActiveConnection=$connection
  }

    $command.CommandTimeout = 60

    $start = get-date
    logMsg("------------------------- Executing the query ------------------- ")
    logMsg("Query                 :  " + $QueryText)
    $command.CommandText = $QueryText
    $CntRow=0
    Try
     {
      if(  $Provider -eq "1" -or $Provider -eq "2" -or $Provider -eq "3" )
      {
       try
       {
       $result = $command.ExecuteReader()  
       while ($result.Read())
       {
        $CntRow++
        $LngData=""
        logMsg("Row #" + $CntRow ) (3)
        for($i=0;$i -le $result.FieldCount-1;$i++) 
         {
          $LngData= $LngData + " - Field #" + ($i+1) + " - Data:" + $result.GetValue($i)
         }   
          logMsg($LngData) (1)
         }
        }
        catch
        {
         logMsg("Error Executing the Query:" + $Error[0].Exception) (2)
        }
      }
      if( $Provider -eq "4" -or $Provider -eq "5" )
      {
       try
       {
       $result = $command.Execute()  
       while ($result.EOF -eq $False)
       {
        $CntRow++
        $LngData=""
        logMsg("Row #" + $CntRow ) (3)
        for($i=0;$i -le $result.Fields.Count-1;$i++) 
         {
          $LngData= $LngData + " - Field #" + ($i+1) + " - Data:" + $result.Fields($i).Value
         }   
          logMsg($LngData) (1)
          $result.MoveNext()
        }
       }
       catch
      {
       logMsg("Error Executing the Query:" + $Error[0].Exception) (2)
      }
      }
      
     }
    catch
      {
       logMsg("Error Executing the Query:" + $Error[0].Exception) (2)
      }
     $end = get-date
     logMsg("Time required (ms)    :  " + (New-TimeSpan -Start $start -End $end).TotalMilliseconds)
     $connection.Close()
     logMsg("Time spent (ms) Proccess:  " +$sw.elapsed) (1)
  }
catch 
  {
    logMsg("Error Executing the Query:" + $Error[0].Exception) (2)
   }
