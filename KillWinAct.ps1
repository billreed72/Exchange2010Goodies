#==============================================================
# KILLING WINDOWS ACTIVATION: General Instructions
#==============================================================
# 1. Disable SL UI Notification service, THEN REBOOT!
# 2. Take Ownership, Set Full Control Permissions & rename files:
# - c:\Windows\System32\SLLUA.exe
# - c:\Windows\System32\SLUI.exe
# - c:\Windows\System32\SLUINotify.dll
# 3. Change DATE/TIME to 1 year in future > REBOOT
# 4. Install IIS ROLE
# 5. Install Windows Process Activation service FEATURE
# 6. Change DATE/TIME back to correct DATE/TIME > REBOOT
#==============================================================
sc config sppuinotify start= disabled
# MUST REBOOT HERE!


TAKEOWN /F c:\Windows\System32\SLUI.exe
ICACLS c:\Windows\System32\SLUI.exe /grant administrator:F
REN c:\Windows\System32\SLUI.exe SLUI.OLD

TAKEOWN /F c:\Windows\System32\SPPUINotify.dll
ICACLS c:\Windows\System32\SPPUINotify.dll /grant administrator:F
REN c:\Windows\System32\SPPUINotify.dll SPPUINotify.OLD

#==============================================================
#                i couldn't find sllua.exe                   #
#==============================================================
#TAKEOWN /F c:\Windows\System32\SLLUA.exe
#ICACLS c:\Windows\System32\SLLUA.exe /grant administrator:F
#REN c:\Windows\System32\SLLUA.exe SLLUA.OLD
#==============================================================
