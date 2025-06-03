# $procs = Get-Process *AVEVA*

# foreach ($proc in $procs) {
#     try {
#         $proc.kill()
#     }
#     catch {
#         Write-Error "FAILED for " $proc
#         CONTINUE
#     }
# # }
# Get-Process | Where-Object {$_.name -match 'AVEVA.NET.Gateways.IED'}
# for ($i = 0; $i -lt 100; $i++) {
#     $a = Get-Process | Where-Object {$_.name -match 'AVEVA'} | Select-Object -Property Id
# foreach ($b in $a) {
#     try {
#             Stop-Process -Id $b.Id            
#         }
#     catch {
#         CONTINUE
#     }
#         }


#     $a = Get-Process | Where-Object {$_.name -match 'visio'} | Select-Object -Property Id
# foreach ($b in $a) {
#     try { Stop-Process -Id $b.Id }
#     catch { CONTINUE }
# }
    # Start-Sleep -Milliseconds 150
# }
# $id = ((quser | Where-Object {$_ -match 'SVC-AVEVA'}) -split ' +')[2]
# Write-Host $id
# logoff $id
# foreach ($b in $a) {
#     $b.Split(' +')[2]
# }
#asd
Get-Job | Stop-Job
Get-Job |  Remove-Job