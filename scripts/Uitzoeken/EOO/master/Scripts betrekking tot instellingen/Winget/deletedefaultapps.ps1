$model = (get-computerinfo).csmanufacturer
if ($model -eq "HP"){
winget uninstall -e "HP Wolf Security" --silent
winget uninstall -e "HP Notifications" -i
winget uninstall -e "HP Connection Optimizer" --silent
winget uninstall -e "HP Audio Control" --silent
winget uninstall -e "HP System Information" --silent
winget uninstall -e "HP Easy Clean" --silent
winget uninstall -e "HP PC Hardware Diagnostics Windows" --silent
winget uninstall -e "HP Privacy Settings" --silent
winget uninstall -e "HP QuickDrop" --silent
winget uninstall -e "HP System Information" --silent
winget uninstall -e "Microsoft 365 - nl-nl" --silent
winget uninstall -e "Microsoft 365 - en-us" --silent
}
elseif ($model -eq "Lenovo") {
    winget uninstall -e "Microsoft 365 - de-de" 
    winget uninstall -e "Microsoft 365 - nl-nl" 
    winget uninstall -e "Microsoft 365 - en-us"
    winget uninstall -e "Microsoft 365 - fr-fr"
    winget uninstall -e "Microsoft 365 - it-it"
    winget uninstall -e "WebAdvisor van McAfee"
   
}
 