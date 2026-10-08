@ECHO OFF
:: Copy to start.local.cmd next to start.bat and fill in. start.local.cmd is git-ignored:
:: it holds the site's LocalAdmin password and stays on the USB stick, never in the repo.
:: Leave a value empty and start.bat asks for it when it is needed (the password hidden).
:: Avoid % in the password - batch expands it.
SET "LOCALADMIN_PASSWORD="
SET "INSTALL_SHARE=\\server\Software"
