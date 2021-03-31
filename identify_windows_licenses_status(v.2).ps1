Add-Type -AssemblyName PresentationFramework

#Build the GUI
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="Window" Title="Windows Servers License Status Report" WindowStartupLocation = "CenterScreen" 
    Width = "1000" Height = "400" ShowInTaskbar = "True">

<StackPanel>
  <Button Name="getlsstatus" Content="Get License Status" Width="100" HorizontalAlignment="Left" Margin="10"/>
  <Button Name="exportexcel" Content="Export to Excel" Width="100" HorizontalAlignment="Left" Margin="10"/>

<ListView Name="datagrid" Grid.Column="4" Grid.Row="0" Margin="30,30,30,30" Height="600" ScrollViewer.VerticalScrollBarVisibility="Visible"  >

            <ListView.View>
            
                <GridView>
                    <GridViewColumn Header="OS" Width="210" DisplayMemberBinding="{Binding os}" />
                    <GridViewColumn Header="License" Width="210" DisplayMemberBinding="{Binding license}" />
                    <GridViewColumn Header="Computer Name" Width="150" DisplayMemberBinding="{Binding Computername}"/>
                     <GridViewColumn Header="License Status" Width="150" DisplayMemberBinding="{Binding licensestatus}" />
                    <GridViewColumn Header="Computer Status" Width="200" DisplayMemberBinding="{Binding status}"/>

                    
                </GridView>
               
            </ListView.View>
        </ListView>
  
</StackPanel>

</Window>
"@ 


$reader=(New-Object System.Xml.XmlNodeReader $xaml)
$Window=[Windows.Markup.XamlReader]::Load( $reader )

$getlsstatus = $Window.FindName("getlsstatus")
$datagrid = $Window.FindName("datagrid")
$exportexcel = $Window.FindName("exportexcel")


$cred=Get-Credential "domain\administrator"

#When click the Button Get License Status
$getlsstatus.Add_Click({

$servernames = Get-ADComputer -Filter 'OperatingSystem -like "Windows Server*"'

$srvstatus=@()

#Create an arraylist with objects to identify which Servers are online/offline
ForEach($name in $servernames){

$servernames=$name.Name

if (Test-Connection -ComputerName $servernames  -Count 1 -Quiet )
  {
        $o = new-object  psobject
        $o | add-member -membertype noteproperty -name computername -value (Get-CIMInstance Win32_Computersystem -computername $servernames).Caption | Out-Null
        $o | add-member -membertype noteproperty -name status -value "Online"| Out-Null

  }

else
  {
        $o = new-object  psobject
        $o | add-member -membertype noteproperty -name computername -value  $servernames | Out-Null
        $o | add-member -membertype noteproperty -name status -value "Offline"| Out-Null
  }
    $srvstatus+=$o

}

$lsstatus=@()

#Run the Invoke-Command to identify License Status base on the results of the previous arraylist $srvstatus 
ForEach($servers in $srvstatus){

$status=$servers.status

$names=$servers.computername
   
        $ScriptBlock={
            $licenses=@{
             os=(Get-CIMInstance Win32_Operatingsystem ).Caption
             license= ( Get-CimInstance -ClassName SoftwareLicensingProduct | where {$_.productkeyid} ).Description
             Computername=$env:Computername
             licensestatus= ( Get-CimInstance -ClassName SoftwareLicensingProduct | where {$_.productkeyid} ).LicenseStatus
             status=$status
            }
            $localresult=New-Object psobject -Property $licenses
            $localresult

      
            }

   
if ($status -eq "Online"){
   $lsstatus+= Invoke-Command -cn $names -Credential $cred -ScriptBlock $Scriptblock -ArgumentList $names,$status
    }
    
if ($status -eq "Offline"){
            $licenses=@{
            os="Not Available"
            license= "Not Available"
            Computername=$names
            licensestatus= "Not Available"
            status=$status
            }
            $r=New-Object psobject -Property $licenses
            $r
}

$lsstatus+=$r



}

$lsstatus| Select-Object -property os,license,Computername,licensestatus,status


#Bind the data in DatagridView

$datagrid=$Window.FindName("datagrid").ItemsSource = $lsstatus

})

#Export data from Datagrid to Csv in the Location that you will select
$exportexcel.Add_Click({


$OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null


$datagrid.ItemsSource| Select-Object -property os,license,Computername,licensestatus,status |export-csv $OpenFileDialog.FileName -NoType
    

})
$Window.ShowDialog()
