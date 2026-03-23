if((Test-Path -LiteralPath "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\Authentication\EnableWebSignIn") -ne $true) {  New-Item "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\Authentication\EnableWebSignIn" -force};
New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\PolicyManager\default\Authentication\EnableWebSignIn' -Name 'Behavior' -Value 32 -PropertyType DWord -Force
New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\PolicyManager\default\Authentication\EnableWebSignIn' -Name 'highrange' -Value 2 -PropertyType DWord -Force
New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\PolicyManager\default\Authentication\EnableWebSignIn' -Name 'lowrange' -Value 0 -PropertyType DWord -Force
New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\PolicyManager\default\Authentication\EnableWebSignIn' -Name 'mergealgorithm' -Value 3 -PropertyType DWord -Force
New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\PolicyManager\default\Authentication\EnableWebSignIn' -Name 'policytype' -Value 4 -PropertyType DWord -Force
New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\PolicyManager\default\Authentication\EnableWebSignIn' -Name 'value' -Value 0 -PropertyType DWord -Force
#New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\Credential Providers\{60b78e88-ead8-445c-9cfd-0b87f74ea6cd}' -Name 'Disabled' -Value 1 -PropertyType DWORD -Force
