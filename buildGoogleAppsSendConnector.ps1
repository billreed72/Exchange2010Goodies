
new-SendConnector 
  -Name 'GoogleApps Sending' 
  -Usage 'Internal' 
  -AddressSpaces 'SMTP:*;1' 
  -IsScopedConnector $false 
  -DNSRoutingEnabled $false 
  -SmartHosts 'smtp.gmail.com' 
  -SmartHostAuthMechanism 'BasicAuthRequireTLS' 
  -UseExternalDNSServersEnabled $false 
  -AuthenticationCredential 'System.Management.Automation.PSCredential' 
  -SourceTransportServers 'DEX'


