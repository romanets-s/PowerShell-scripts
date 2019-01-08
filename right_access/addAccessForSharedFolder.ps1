Param(
   [string]$users,
   [string]$pathToShared
)
try {
    if($users.Contains(";")){
        foreach($user in $users.Split(";")){
            $acl = Get-Acl $pathToShared
            $ar = New-Object  system.security.accesscontrol.filesystemaccessrule($user.Trim(), 'Modify', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
            $acl.SetAccessRule($ar)
            Set-Acl $pathToShared $acl
        }
    }
    else{
        $acl = Get-Acl $pathToShared
        $ar = New-Object  system.security.accesscontrol.filesystemaccessrule($users.Trim(), 'Modify', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
        $acl.SetAccessRule($ar)
        Set-Acl $pathToShared $acl
    }

} catch {
    return "false"
}
return $env:computername
