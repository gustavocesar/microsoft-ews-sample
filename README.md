# Microsoft EWS - Exchange Web Service (Sample)

## Configure:

 - Open file Config/ExchangeServiceConfig.cs
 - Set `Username` e `Password` of the Mail Box
 - Set `FindSubjects` to be search on the Mail Box

## Run:

```
dotnet restore
dotnet watch run
```

## Restults
 - Attachments found: `Download`
 - Exported message: `Export`

## Useful Links
 - https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/start-using-web-services-in-exchange
 - https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications
 - https://docs.microsoft.com/en-us/exchange/policy-and-compliance/ediscovery/message-properties-and-search-operators?view=exchserver-2019